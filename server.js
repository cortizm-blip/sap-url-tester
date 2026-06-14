const express = require("express");
const fetch = require("node-fetch");
const https = require("https");
const path = require("path");
const multer = require("multer");
const XLSX = require("xlsx");

const app = express();
app.use(express.json({ limit: "10mb" }));
app.use(express.static(path.join(__dirname, "public")));

const upload = multer({ storage: multer.memoryStorage() });
const httpsAgent = new https.Agent({ rejectUnauthorized: false });

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

// ── Endpoints ─────────────────────────────────────────────────
app.post("/api/test-url", async (req, res) => {
  const { url, username, password, type } = req.body;
  if (!url) return res.status(400).json({ error: "URL requerida" });
  res.json(await testUrl({ url, username, password, type }));
});

// ── Parse Excel ───────────────────────────────────────────────
app.post("/api/parse-excel", upload.single("file"), (req, res) => {
  try {
    const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
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
          rowIndex: i + (hasHeader ? 2 : 1), // 1-based Excel row
          name: name || url,
          url,
          type: type.toLowerCase().includes("soa") ? "soamanager" : "apirest"
        });
      }
    }
    // Guardar el buffer original en memoria para usarlo en export
    // Lo devolvemos como base64 para que el frontend lo guarde
    const xlsxBase64 = req.file.buffer.toString("base64");
    res.json({ count: urls.length, urls, xlsxBase64, fileName: req.file.originalname });
  } catch (e) {
    res.status(400).json({ error: "Error leyendo Excel: " + e.message });
  }
});

// ── Export Excel con estatus coloreado ────────────────────────
app.post("/api/export-excel", express.json({ limit: "20mb" }), (req, res) => {
  try {
    const { xlsxBase64, fileName, results } = req.body;
    // results: [{ url, ok, status, error_type }]

    const buffer = Buffer.from(xlsxBase64, "base64");
    const workbook = XLSX.read(buffer, { type: "buffer", cellStyles: true });
    const sheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];

    // Obtener rango
    const range = XLSX.utils.decode_range(ws["!ref"] || "A1:D1");

    // Encontrar columna "Estatus" en header (fila 1)
    let estatusCol = 3; // default columna D (index 3)
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
      if (cell && String(cell.v).toLowerCase().includes("estatus")) {
        estatusCol = c;
        break;
      }
    }

    // Encontrar columna URL
    let urlCol = 1; // default columna B
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
      if (cell && String(cell.v).toLowerCase() === "url") {
        urlCol = c;
        break;
      }
    }

    // Crear mapa url -> resultado
    const resultMap = {};
    for (const r of results) resultMap[r.url] = r;

    // Recorrer filas de datos (desde fila 2, index 1)
    for (let row = 1; row <= range.e.r; row++) {
      const urlCell = ws[XLSX.utils.encode_cell({ r: row, c: urlCol })];
      if (!urlCell) continue;
      const url = String(urlCell.v).trim();
      const result = resultMap[url];
      if (!result) continue;

      const addr = XLSX.utils.encode_cell({ r: row, c: estatusCol });
      const ok = result.ok;
      const statusText = ok
        ? `OK - HTTP ${result.status}`
        : result.error_type || `ERROR - HTTP ${result.status || "N/A"}`;

      ws[addr] = {
        v: statusText,
        t: "s",
        s: {
          fill: { fgColor: { rgb: ok ? "C6EFCE" : "FFC7CE" } },
          font: { color: { rgb: ok ? "276221" : "9C0006" }, bold: true },
          alignment: { horizontal: "center" }
        }
      };
    }

    // Actualizar rango si estatus col está fuera
    if (estatusCol > range.e.c) range.e.c = estatusCol;
    ws["!ref"] = XLSX.utils.encode_range(range);

    const outBuffer = XLSX.write(workbook, { type: "buffer", bookType: "xlsx", cellStyles: true });
    const outName = (fileName || "resultado").replace(".xlsx", "") + "_estatus.xlsx";

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    res.send(outBuffer);
  } catch (e) {
    res.status(500).json({ error: "Error generando Excel: " + e.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SAP URL Tester en puerto ${PORT}`));
