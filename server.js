const express = require("express");
const fetch = require("node-fetch");
const https = require("https");
const path = require("path");

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// Agente HTTPS que ignora certificados auto-firmados (común en SAP on-premise)
const httpsAgent = new https.Agent({ rejectUnauthorized: false });

/**
 * POST /api/test-url
 * Body: { url, username, password, type }
 * Hace el request desde el servidor (evita CORS del browser)
 */
app.post("/api/test-url", async (req, res) => {
  const { url, username, password, type } = req.body;

  if (!url) return res.status(400).json({ error: "URL requerida" });

  const startTime = Date.now();
  const result = {
    url,
    type: type || "unknown",
    timestamp: new Date().toISOString(),
  };

  try {
    const headers = {};

    // Basic Auth si se proporcionan credenciales
    if (username && password) {
      const b64 = Buffer.from(`${username}:${password}`).toString("base64");
      headers["Authorization"] = `Basic ${b64}`;
    }

    // Para SOA Manager WSDL, pedir el WSDL explícitamente
    if (type === "soamanager") {
      headers["Accept"] =
        "text/xml,application/xml,application/xhtml+xml,text/html";
    }

    const response = await fetch(url, {
      method: "GET",
      headers,
      agent: url.startsWith("https") ? httpsAgent : undefined,
      timeout: 15000, // 15 segundos timeout
    });

    const elapsed = Date.now() - startTime;
    const bodyText = await response.text();

    result.status = response.status;
    result.statusText = response.statusText;
    result.elapsed_ms = elapsed;
    result.reachable = true;
    result.headers = Object.fromEntries(response.headers.entries());

    // Detectar tipo de contenido
    const ct = response.headers.get("content-type") || "";
    result.content_type = ct;

    // Para WSDL: verificar que el body sea XML válido con definiciones WSDL
    if (type === "soamanager") {
      result.has_wsdl =
        bodyText.includes("wsdl:definitions") ||
        bodyText.includes("definitions xmlns");
      result.body_preview = bodyText.substring(0, 300);
    } else {
      // Para API REST: mostrar preview del body
      result.body_preview = bodyText.substring(0, 300);
    }

    // Auth result
    if (response.status === 401) {
      result.auth_status = "FAIL - Credenciales inválidas";
      result.ok = false;
    } else if (response.status === 403) {
      result.auth_status = "FAIL - Sin autorización";
      result.ok = false;
    } else if (response.status >= 200 && response.status < 400) {
      result.auth_status = username ? "OK - Autenticado" : "OK - Sin auth";
      result.ok = true;
    } else {
      result.auth_status = `HTTP ${response.status}`;
      result.ok = false;
    }
  } catch (err) {
    const elapsed = Date.now() - startTime;
    result.elapsed_ms = elapsed;
    result.reachable = false;
    result.ok = false;
    result.error = err.message;

    // Clasificar tipo de error
    if (err.message.includes("ECONNREFUSED")) {
      result.error_type = "CONEXIÓN RECHAZADA - Puerto cerrado o servicio caído";
    } else if (
      err.message.includes("ENOTFOUND") ||
      err.message.includes("getaddrinfo")
    ) {
      result.error_type = "DNS - Host no encontrado (no accesible desde internet)";
    } else if (err.message.includes("ETIMEDOUT") || err.message.includes("timeout")) {
      result.error_type = "TIMEOUT - No responde en 15s (posible firewall/VPN requerida)";
    } else if (err.message.includes("CERT") || err.message.includes("SSL")) {
      result.error_type = "SSL/TLS - Certificado inválido";
    } else {
      result.error_type = "ERROR DE RED";
    }
  }

  res.json(result);
});

/**
 * POST /api/test-batch
 * Body: { urls: [{url, username, password, type, name}], globalUser, globalPass }
 * Prueba múltiples URLs con concurrencia controlada
 */
app.post("/api/test-batch", async (req, res) => {
  const { urls, globalUser, globalPass } = req.body;

  if (!urls || !Array.isArray(urls)) {
    return res.status(400).json({ error: "Array de URLs requerido" });
  }

  // Limitar a 60 URLs máximo
  const urlList = urls.slice(0, 60);

  // Concurrencia de 5 requests paralelos para no saturar
  const CONCURRENCY = 5;
  const results = [];

  for (let i = 0; i < urlList.length; i += CONCURRENCY) {
    const batch = urlList.slice(i, i + CONCURRENCY);
    const batchResults = await Promise.all(
      batch.map(async (item) => {
        const testReq = {
          body: {
            url: item.url,
            username: item.username || globalUser || "",
            password: item.password || globalPass || "",
            type: item.type || "apirest",
          },
        };

        return new Promise((resolve) => {
          const mockRes = {
            json: (data) => resolve({ ...data, name: item.name || item.url }),
            status: () => mockRes,
          };
          // Reusar el mismo handler
          app._router.handle(
            { ...testReq, method: "POST", url: "/api/test-url" },
            mockRes,
            () => {}
          );
        });
      })
    );

    // Llamar directamente a la lógica sin pasar por el router
    const directResults = await Promise.all(
      batch.map(async (item) => {
        const username = item.username || globalUser || "";
        const password = item.password || globalPass || "";
        const { url, type } = item;

        const startTime = Date.now();
        const result = {
          name: item.name || url,
          url,
          type: type || "apirest",
          timestamp: new Date().toISOString(),
        };

        try {
          const headers = {};
          if (username && password) {
            const b64 = Buffer.from(`${username}:${password}`).toString("base64");
            headers["Authorization"] = `Basic ${b64}`;
          }
          if (type === "soamanager") {
            headers["Accept"] = "text/xml,application/xml";
          }

          const response = await fetch(url, {
            method: "GET",
            headers,
            agent: url.startsWith("https") ? httpsAgent : undefined,
            timeout: 15000,
          });

          const elapsed = Date.now() - startTime;
          const bodyText = await response.text();

          result.status = response.status;
          result.statusText = response.statusText;
          result.elapsed_ms = elapsed;
          result.reachable = true;
          result.content_type = response.headers.get("content-type") || "";

          if (type === "soamanager") {
            result.has_wsdl =
              bodyText.includes("wsdl:definitions") ||
              bodyText.includes("definitions xmlns");
          }

          if (response.status === 401) {
            result.auth_status = "FAIL - Credenciales inválidas";
            result.ok = false;
          } else if (response.status === 403) {
            result.auth_status = "FAIL - Sin autorización";
            result.ok = false;
          } else if (response.status >= 200 && response.status < 400) {
            result.auth_status = username ? "OK - Autenticado" : "OK - Sin auth";
            result.ok = true;
          } else {
            result.auth_status = `HTTP ${response.status}`;
            result.ok = false;
          }
        } catch (err) {
          const elapsed = Date.now() - startTime;
          result.elapsed_ms = elapsed;
          result.reachable = false;
          result.ok = false;
          result.error = err.message;

          if (err.message.includes("ECONNREFUSED")) {
            result.error_type = "CONEXIÓN RECHAZADA";
          } else if (err.message.includes("ENOTFOUND") || err.message.includes("getaddrinfo")) {
            result.error_type = "DNS - Host no encontrado";
          } else if (err.message.includes("ETIMEDOUT") || err.message.includes("timeout")) {
            result.error_type = "TIMEOUT - Posible firewall/VPN";
          } else if (err.message.includes("CERT") || err.message.includes("SSL")) {
            result.error_type = "SSL/TLS";
          } else {
            result.error_type = "ERROR DE RED";
          }
        }

        return result;
      })
    );

    results.push(...directResults);
  }

  res.json({ total: results.length, results });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`SAP URL Tester corriendo en http://localhost:${PORT}`);
});
