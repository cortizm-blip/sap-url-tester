const express = require("express");
const fetch = require("node-fetch");
const https = require("https");
const path = require("path");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const os = require("os");
const { execFile } = require("child_process");

const app = express();
app.use(express.json({ limit: "20mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ storage: multer.memoryStorage() });
const httpsAgent = new https.Agent({ rejectUnauthorized: false });

let tempExcelPath = null;
let tempExcelName = "resultado.xlsx";

// ── Test URL ──────────────────────────────────────────────────
async function testUrl({ url, username, password, type }) {
  const startTime = Date.now();
  const result = { url, type: type || "apirest", timestamp: new Date().toISOString() };
  try {
    const headers = {};
    if (username && password)
      headers["Authorization"] = "Basic " + Buffer.from(`${username}:${password}`).toString("base64");
    if (type === "soamanager") headers["Accept"] = "text/xml,application/xml";

    const response = await fetch(url, {
      method: "GET", headers,
      agent: url.startsWith("https") ? httpsAgent : undefined,
      timeout: 10000,
    });
    const elapsed = Date.now() - startTime;
    const bodyText = await response.text();
    result.status = response.status;
    result.statusText = response.statusText;
    result.elapsed_ms = elapsed;
    result.reachable = true;
    result.content_type = response.headers.get("content-type") || "";
    result.body_preview = bodyText.substring(0, 500);

    if (type === "soamanager") {
      result.has_wsdl = bodyText.includes("wsdl:definitions") || bodyText.includes("definitions xmlns");
      if (result.has_wsdl) result.wsdl_info = parseWsdl(bodyText, url);
    }

    if (response.status === 401) { result.auth_status = "FAIL - Credenciales inválidas"; result.ok = false; }
    else if (response.status === 403) { result.auth_status = "FAIL - Sin autorización"; result.ok = false; }
    else if (response.status >= 200 && response.status < 400) { result.auth_status = username ? "OK - Autenticado" : "OK - Sin auth"; result.ok = true; }
    else { result.auth_status = `HTTP ${response.status}`; result.ok = false; }
  } catch (err) {
    result.elapsed_ms = Date.now() - startTime;
    result.reachable = false; result.ok = false; result.error = err.message;
    if (err.message.includes("ECONNREFUSED")) result.error_type = "CONEXIÓN RECHAZADA";
    else if (err.message.includes("ENOTFOUND") || err.message.includes("getaddrinfo")) result.error_type = "DNS - Host no encontrado";
    else if (err.message.includes("ETIMEDOUT") || err.message.includes("timeout")) result.error_type = "TIMEOUT";
    else if (err.message.includes("CERT") || err.message.includes("SSL")) result.error_type = "SSL/TLS";
    else result.error_type = "ERROR DE RED";
  }
  return result;
}

function parseWsdl(xml, sourceUrl) {
  const info = { source_url: sourceUrl, operations: [], service_name: "", port_name: "", soap_address: "" };
  const svcMatch = xml.match(/wsdl:service[^>]+name=["']([^"']+)["']/);
  if (svcMatch) info.service_name = svcMatch[1];
  const portMatch = xml.match(/wsdl:port[^>]+name=["']([^"']+)["']/);
  if (portMatch) info.port_name = portMatch[1];
  const addrMatch = xml.match(/soap[^:]*:address[^>]+location=["']([^"']+)["']/);
  if (addrMatch) info.soap_address = addrMatch[1];
  const opRegex = /wsdl:operation[^>]+name=["']([^"']+)["']/g;
  let m;
  while ((m = opRegex.exec(xml)) !== null)
    if (!info.operations.includes(m[1])) info.operations.push(m[1]);
  try {
    const pub = new URL(sourceUrl);
    info.public_host = pub.hostname; info.public_url = sourceUrl;
    if (info.soap_address) {
      const internal = new URL(info.soap_address);
      info.internal_host = internal.hostname; info.internal_url = info.soap_address;
    }
  } catch (e) {}
  return info;
}

app.post("/api/test-url", async (req, res) => {
  const { url, username, password, type } = req.body;
  if (!url) return res.status(400).json({ error: "URL requerida" });
  res.json(await testUrl({ url, username, password, type }));
});

// ── Parse Excel ───────────────────────────────────────────────
app.post("/api/parse-excel", upload.single("file"), (req, res) => {
  try {
    const buffer = req.file.buffer;
    tempExcelPath = path.join(os.tmpdir(), "sap_tester_upload.xlsx");
    tempExcelName = req.file.originalname;
    fs.writeFileSync(tempExcelPath, buffer);

    const workbook = XLSX.read(buffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    const firstRow = rows[0] || [];
    const hasHeader = firstRow.some(c =>
      ["nombre","name","url","tipo","type","estatus","status"].includes(String(c).toLowerCase().trim())
    );
    const dataRows = hasHeader ? rows.slice(1) : rows;
    const urls = [];

    for (const row of dataRows) {
      if (!row || row.length === 0) continue;
      const cells = row.map(c => String(c).trim());
      let name, url, type;
      if (cells[0] && cells[0].startsWith("http")) {
        url = cells[0]; name = cells[0]; type = cells[1] || "apirest";
      } else if (cells.length >= 2) {
        name = cells[0]; url = cells[1]; type = cells[2] || "apirest";
      }
      if (url && url.startsWith("http")) {
        urls.push({ name: name || url, url, type: type.toLowerCase().includes("soa") ? "soamanager" : "apirest" });
      }
    }
    res.json({ count: urls.length, urls, fileName: req.file.originalname });
  } catch (e) {
    res.status(400).json({ error: "Error leyendo Excel: " + e.message });
  }
});

// ── Export Excel con colores via Python/openpyxl ──────────────
app.post("/api/export-excel", (req, res) => {
  try {
    const { results, xlsxBase64 } = req.body;
    if (!results || !results.length) return res.status(400).json({ error: "Sin resultados" });

    // Restaurar Excel desde base64 del browser si el temp fue borrado
    if (xlsxBase64 && (!tempExcelPath || !fs.existsSync(tempExcelPath))) {
      tempExcelPath = path.join(os.tmpdir(), "sap_tester_upload.xlsx");
      fs.writeFileSync(tempExcelPath, Buffer.from(xlsxBase64, "base64"));
    }

    if (!tempExcelPath || !fs.existsSync(tempExcelPath))
      return res.status(400).json({ error: "Sube el Excel nuevamente e intenta otra vez." });

    const outPath = path.join(os.tmpdir(), "sap_tester_output.xlsx");
    const resultsJson = JSON.stringify(results);
    const scriptPath = path.join(__dirname, "color_excel.py");

    execFile("python3", [scriptPath, tempExcelPath, outPath, resultsJson], { timeout: 30000 }, (err, stdout, stderr) => {
      if (err) {
        console.error("Python error:", stderr);
        return res.status(500).json({ error: "Error coloreando Excel: " + (stderr || err.message) });
      }
      try {
        const pyResult = JSON.parse(stdout.trim());
        if (!pyResult.ok) return res.status(500).json({ error: "Error en script Python" });
      } catch(e) {}

      const outBuffer = fs.readFileSync(outPath);
      const outName = tempExcelName.replace(/\.xlsx$/i, "") + "_estatus.xlsx";
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
      res.send(outBuffer);

      // Cleanup
      try { fs.unlinkSync(outPath); } catch(e) {}
    });
  } catch (e) {
    res.status(500).json({ error: "Error: " + e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SAP URL Tester en puerto ${PORT}`));
