const express = require("express");
const fetch = require("node-fetch");
const https = require("https");
const path = require("path");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const os = require("os");

const app = express();
app.use(express.json({ limit: "20mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ storage: multer.memoryStorage() });
const httpsAgent = new https.Agent({ rejectUnauthorized: false });

// Guardar Excel en temp
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
      timeout: 8000,
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

// ── Test single URL ───────────────────────────────────────────
app.post("/api/test-url", async (req, res) => {
  const { url, username, password, type } = req.body;
  if (!url) return res.status(400).json({ error: "URL requerida" });
  res.json(await testUrl({ url, username, password, type }));
});

// ── Parse Excel — guarda archivo en temp ─────────────────────
app.post("/api/parse-excel", upload.single("file"), (req, res) => {
  try {
    const buffer = req.file.buffer;
    // Guardar en temp para usar después en export
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

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      if (!row || row.length === 0) continue;
      const cells = row.map(c => String(c).trim());
      let name, url, type;
      if (cells[0] && cells[0].startsWith("http")) {
        url = cells[0]; name = cells[0]; type = cells[1] || "apirest";
      } else if (cells.length >= 2) {
        name = cells[0]; url = cells[1]; type = cells[2] || "apirest";
      }
      if (url && url.startsWith("http")) {
        urls.push({
          name: name || url, url,
          type: type.toLowerCase().includes("soa") ? "soamanager" : "apirest"
        });
      }
    }
    res.json({ count: urls.length, urls, fileName: req.file.originalname });
  } catch (e) {
    res.status(400).json({ error: "Error leyendo Excel: " + e.message });
  }
});

// ── Export Excel con estatus coloreado ────────────────────────
app.post("/api/export-excel", (req, res) => {
  try {
    const { results, fileName, xlsxBase64 } = req.body;
    if (!results || !results.length) return res.status(400).json({ error: "Sin resultados" });

    let buffer;
    if (tempExcelPath && fs.existsSync(tempExcelPath)) {
      buffer = fs.readFileSync(tempExcelPath);
    } else if (xlsxBase64) {
      buffer = Buffer.from(xlsxBase64, "base64");
    } else {
      return res.status(400).json({ error: "No hay Excel cargado. Sube el archivo Excel nuevamente." });
    }
    const workbook = XLSX.read(buffer, { type: "buffer", cellStyles: true });
    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1:D1");

    // Detectar columnas por header
    let urlCol = 1, nombreCol = 0, estatusCol = 3;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
      if (!cell) continue;
      const v = String(cell.v).toLowerCase().trim();
      if (v === "url") urlCol = c;
      if (v === "nombre" || v === "name") nombreCol = c;
      if (v.includes("estatus") || v.includes("status")) estatusCol = c;
    }

    // Mapa por URL y por nombre
    const byUrl = {}, byName = {};
    for (const r of results) {
      if (r.url) byUrl[r.url.trim()] = r;
      if (r.name) byName[r.name.trim()] = r;
    }

    let updated = 0;
    for (let row = 1; row <= range.e.r; row++) {
      const urlCell    = ws[XLSX.utils.encode_cell({ r: row, c: urlCol })];
      const nombreCell = ws[XLSX.utils.encode_cell({ r: row, c: nombreCol })];
      if (!urlCell && !nombreCell) continue;

      const urlVal    = urlCell    ? String(urlCell.v).trim()    : "";
      const nombreVal = nombreCell ? String(nombreCell.v).trim() : "";

      // Match por URL exacta primero, luego por nombre
      const result = byUrl[urlVal] || byName[nombreVal];
      if (!result) continue;

      const ok = result.ok;
      let statusText = "";
      if (ok) {
        statusText = `OK - HTTP ${result.status}`;
      } else if (!result.reachable) {
        statusText = result.error_type || "ERROR - Sin conexión";
      } else {
        statusText = `ERROR - HTTP ${result.status || "N/A"}`;
      }

      const addr = XLSX.utils.encode_cell({ r: row, c: estatusCol });
      ws[addr] = {
        v: statusText, t: "s",
        s: {
          fill: { patternType: "solid", fgColor: { rgb: ok ? "C6EFCE" : "FFC7CE" } },
          font: { color: { rgb: ok ? "276221" : "9C0006" }, bold: true, sz: 11 },
          alignment: { horizontal: "center", vertical: "center" }
        }
      };
      updated++;
    }

    // Asegurar que la columna Estatus esté dentro del rango
    if (estatusCol > range.e.c) range.e.c = estatusCol;
    ws["!ref"] = XLSX.utils.encode_range(range);

    // Ajustar ancho columna Estatus
    if (!ws["!cols"]) ws["!cols"] = [];
    ws["!cols"][estatusCol] = { wch: 22 };

    const outBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx", cellStyles: true });
    const outName = (tempExcelName || "resultado").replace(/\.xlsx$/i, "") + "_estatus.xlsx";

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);

    res.send(outBuffer);
  } catch (e) {
    res.status(500).json({ error: "Error generando Excel: " + e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SAP URL Tester en puerto ${PORT}`));
