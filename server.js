const express = require("express");
const fetch = require("node-fetch");
const https = require("https");
const path = require("path");
const multer = require("multer");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");
const fs = require("fs");
const os = require("os");

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
    // SOA Manager: GET para obtener WSDL
    // API REST: POST con body vacío (SAP ICM ignora GET/HEAD en servicios REST)
    const method = type === "soamanager" ? "GET" : "POST";
    if (type === "apirest") {
      headers["Content-Type"] = "application/json";
    }
    const response = await fetch(url, {
      method,
      headers,
      body: type === "apirest" ? "{}" : undefined,
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
    if (type === "apirest") {
      // Para REST: cualquier respuesta HTTP = conectividad OK (el servicio puede requerir POST)
      result.ok = true;
      if (response.status === 200) result.auth_status = "OK - HTTP 200";
      else if (response.status === 401) result.auth_status = "Accesible - Requiere Auth";
      else if (response.status === 403) result.auth_status = "Accesible - Sin permisos";
      else if (response.status === 404) result.auth_status = "Accesible - Servicio GET no soportado";
      else if (response.status === 405) result.auth_status = "Accesible - Método no permitido (OK)";
      else result.auth_status = `Accesible - HTTP ${response.status}`;
    } else {
      if (response.status === 401) { result.auth_status = "FAIL - Credenciales inválidas"; result.ok = false; }
      else if (response.status === 403) { result.auth_status = "FAIL - Sin autorización"; result.ok = false; }
      else if (response.status >= 200 && response.status < 400) { result.auth_status = username ? "OK - Autenticado" : "OK - Sin auth"; result.ok = true; }
      else { result.auth_status = `HTTP ${response.status}`; result.ok = false; }
    }
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

    // Detectar header
    const firstRow = rows[0] || [];
    const hasHeader = firstRow.some(c =>
      ["nombre","name","url","tipo","type","estatus","status"].includes(String(c).toLowerCase().trim())
    );

    // Detectar columnas
    let urlColIdx = 1, nombreColIdx = 0;
    if (hasHeader) {
      firstRow.forEach((c, i) => {
        const v = String(c).toLowerCase().trim();
        if (v === "url") urlColIdx = i;
        if (v === "nombre" || v === "name") nombreColIdx = i;
      });
    }

    const dataRows = hasHeader ? rows.slice(1) : rows;
    const urls = [];

    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      if (!row || row.length === 0) continue;
      const urlVal    = String(row[urlColIdx] || "").trim();
      const nombreVal = String(row[nombreColIdx] || "").trim();
      const tipoVal   = String(row[2] || "").trim();
      if (!urlVal || !urlVal.startsWith("http")) continue;
      urls.push({
        rowIndex: i + (hasHeader ? 2 : 1), // 1-based Excel row number
        name: nombreVal || urlVal,
        url: urlVal,
        type: tipoVal.toLowerCase().includes("soa") ? "soamanager" : "apirest"
      });
    }
    res.json({ count: urls.length, urls, fileName: req.file.originalname });
  } catch (e) {
    res.status(400).json({ error: "Error leyendo Excel: " + e.message });
  }
});

// ── Export Excel con colores ──────────────────────────────────
app.post("/api/export-excel", async (req, res) => {
  try {
    const { results, xlsxBase64 } = req.body;
    if (!results || !results.length) return res.status(400).json({ error: "Sin resultados" });

    // Restaurar Excel desde base64 si el temp fue borrado
    if (xlsxBase64 && (!tempExcelPath || !fs.existsSync(tempExcelPath))) {
      tempExcelPath = path.join(os.tmpdir(), "sap_tester_upload.xlsx");
      fs.writeFileSync(tempExcelPath, Buffer.from(xlsxBase64, "base64"));
    }
    if (!tempExcelPath || !fs.existsSync(tempExcelPath))
      return res.status(400).json({ error: "Sube el Excel nuevamente e intenta otra vez." });

    // Leer Excel con ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(tempExcelPath);
    const ws = workbook.worksheets[0];

    // Detectar fila de header y columnas (puede haber filas vacías al inicio)
    let headerRowNum = 1, urlColNum = 3, nombreColNum = 2, estatusColNum = 5;
    ws.eachRow((row, rowNumber) => {
      if (headerRowNum > 1) return; // ya encontrado
      let isHeader = false;
      row.eachCell((cell) => {
        const v = String(cell.value || "").toLowerCase().trim();
        if (v === "url" || v === "nombre" || v === "name") isHeader = true;
      });
      if (isHeader) {
        headerRowNum = rowNumber;
        row.eachCell((cell, colNumber) => {
          const v = String(cell.value || "").toLowerCase().trim();
          if (v === "url") urlColNum = colNumber;
          if (v === "nombre" || v === "name") nombreColNum = colNumber;
          if (v.includes("estatus") || v === "status") estatusColNum = colNumber;
        });
      }
    });
    console.log(`Header en fila ${headerRowNum}, url=col${urlColNum}, nombre=col${nombreColNum}, estatus=col${estatusColNum}`);

    // Construir mapa URL -> resultado
    const byUrl  = {};
    const byName = {};
    for (const r of results) {
      if (r.url)  byUrl[r.url.trim()]   = r;
      if (r.name) byName[r.name.trim()] = r;
    }

    let updated = 0;
    ws.eachRow((row, rowNumber) => {
      if (rowNumber <= headerRowNum) return; // skip header y filas vacías
      const urlVal    = String(row.getCell(urlColNum).value    || "").trim();
      const nombreVal = String(row.getCell(nombreColNum).value || "").trim();
      if (!urlVal && !nombreVal) return;

      // Match por URL completa, luego por nombre
      const result = byUrl[urlVal] || byName[nombreVal];
      if (!result) return;

      const ok = result.ok;
      let statusText;
      if (ok) {
        statusText = `OK - HTTP ${result.status}`;
      } else if (!result.reachable) {
        statusText = result.error_type || "ERROR - Sin conexión";
      } else {
        statusText = `ERROR - HTTP ${result.status || "N/A"}`;
      }

      const cell = row.getCell(estatusColNum);
      cell.value = statusText;
      cell.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: ok ? "FFC6EFCE" : "FFFFC7CE" } };
      cell.font  = { color: { argb: ok ? "FF276221" : "FF9C0006" }, bold: true, size: 11 };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      updated++;
    });

    console.log(`Export: ${updated} filas actualizadas de ${results.length} resultados`);

    const outName = tempExcelName.replace(/\.xlsx$/i, "") + "_estatus.xlsx";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error("Export error:", e);
    res.status(500).json({ error: "Error generando Excel: " + e.message });
  }
});

// ── Custom request (mini tester) ─────────────────────────────
app.post("/api/custom-request", async (req, res) => {
  const { url, method, body, username, password } = req.body;
  if (!url) return res.status(400).json({ error: "URL requerida" });

  const startTime = Date.now();
  try {
    const headers = {};
    if (username && password)
      headers["Authorization"] = "Basic " + Buffer.from(`${username}:${password}`).toString("base64");
    if (body) headers["Content-Type"] = "application/json";

    const fetchOpts = {
      method: method || "GET",
      headers,
      agent: url.startsWith("https") ? httpsAgent : undefined,
      timeout: 15000,
    };
    if (body && (method === "POST" || method === "PUT" || method === "PATCH")) {
      fetchOpts.body = body;
    }

    const response = await fetch(url, fetchOpts);
    const elapsed = Date.now() - startTime;
    const bodyText = await response.text();

    res.json({
      status: response.status,
      statusText: response.statusText,
      elapsed_ms: elapsed,
      content_type: response.headers.get("content-type") || "",
      body: bodyText.substring(0, 5000),
    });
  } catch (err) {
    res.json({
      status: 0,
      statusText: "Error",
      elapsed_ms: Date.now() - startTime,
      error: err.message,
      body: "",
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`SAP URL Tester en puerto ${PORT}`));
