import http from "node:http";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { readFile } from "node:fs/promises";

const PROJECT_DIR = dirname(fileURLToPath(import.meta.url));
const PORT = Number.parseInt(process.env.PORT || "4173", 10);
const MAX_REQUEST_BYTES = 64 * 1024;
const MAX_RESPONSE_BYTES = 1024 * 1024;
const MAX_MESSAGES = 24;
const MAX_MESSAGE_CHARS = 4000;
const MAX_TOTAL_CHARS = 48_000;
const CHAT_TIMEOUT_MS = 120_000;

const STATIC_FILES = new Map([
  ["/", ["index.html", "text/html; charset=utf-8"]],
  ["/index.html", ["index.html", "text/html; charset=utf-8"]],
  ["/styles.css", ["styles.css", "text/css; charset=utf-8"]],
  ["/app.js", ["app.js", "text/javascript; charset=utf-8"]],
]);

function getHermesConfig(urlValue, apiKey) {
  if (!urlValue || !apiKey) {
    return null;
  }
  try {
    const url = new URL(urlValue);
    if (
      !["http:", "https:"].includes(url.protocol) ||
      url.username ||
      url.password ||
      url.search ||
      url.hash ||
      !url.pathname.replace(/\/+$/, "").endsWith("/v1")
    ) {
      return null;
    }
    url.pathname = url.pathname.replace(/\/+$/, "");
    return { baseUrl: url.toString().replace(/\/$/, ""), apiKey };
  } catch {
    return null;
  }
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "X-Content-Type-Options": "nosniff",
  });
  response.end(JSON.stringify(payload));
}

function isAllowedLocalRequest(request) {
  const host = request.headers.host;
  if (!host) {
    return false;
  }
  let hostname;
  try {
    hostname = new URL(`http://${host}`).hostname;
  } catch {
    return false;
  }
  if (!["localhost", "127.0.0.1"].includes(hostname)) {
    return false;
  }
  if (request.headers.origin && request.headers.origin !== `http://${host}`) {
    return false;
  }
  const fetchSite = request.headers["sec-fetch-site"];
  return !fetchSite || fetchSite === "same-origin" || fetchSite === "none";
}

async function readRequestJson(request) {
  let byteCount = 0;
  const chunks = [];
  for await (const chunk of request) {
    byteCount += chunk.length;
    if (byteCount > MAX_REQUEST_BYTES) {
      const error = new Error("La richiesta supera il limite consentito.");
      error.statusCode = 413;
      throw error;
    }
    chunks.push(chunk);
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    const error = new Error("Il corpo della richiesta deve essere JSON valido.");
    error.statusCode = 400;
    throw error;
  }
}

function validateMessages(payload) {
  if (!payload || !Array.isArray(payload.messages) || payload.messages.length === 0) {
    return "Inserisci almeno un messaggio nella conversazione.";
  }
  if (payload.messages.length > MAX_MESSAGES) {
    return "La conversazione è troppo lunga. Avvia una nuova chat.";
  }
  let totalChars = 0;
  for (const message of payload.messages) {
    if (
      !message ||
      !["user", "assistant"].includes(message.role) ||
      typeof message.content !== "string" ||
      !message.content.trim() ||
      message.content.length > MAX_MESSAGE_CHARS
    ) {
      return "La conversazione contiene un messaggio non valido.";
    }
    totalChars += message.content.length;
  }
  if (totalChars > MAX_TOTAL_CHARS) {
    return "La conversazione supera il limite di testo consentito.";
  }
  if (payload.messages[payload.messages.length - 1].role !== "user") {
    return "L’ultimo messaggio deve essere una domanda dell’utente.";
  }
  return "";
}

async function readLimitedJson(response) {
  if (!response.body) {
    throw new Error("Il servizio Hermes ha restituito una risposta vuota.");
  }
  const reader = response.body.getReader();
  const chunks = [];
  let byteCount = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    byteCount += value.byteLength;
    if (byteCount > MAX_RESPONSE_BYTES) {
      await reader.cancel();
      throw new Error("La risposta del servizio Hermes supera il limite consentito.");
    }
    chunks.push(Buffer.from(value));
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw new Error("Il servizio Hermes ha restituito una risposta non valida.");
  }
}

async function checkHermesConnection(config, fetchImpl) {
  try {
    const response = await fetchImpl(`${config.baseUrl}/models`, {
      headers: { Authorization: `Bearer ${config.apiKey}`, Accept: "application/json" },
      signal: AbortSignal.timeout(8_000),
      redirect: "error",
    });
    return response.ok;
  } catch {
    return false;
  }
}

export function createAppServer({
  hermesUrl = process.env.HERMES_API_URL || "",
  hermesApiKey = process.env.HERMES_API_KEY || "",
  fetchImpl = fetch,
} = {}) {
  const config = getHermesConfig(hermesUrl, hermesApiKey);

  return http.createServer(async (request, response) => {
    response.setHeader("X-Content-Type-Options", "nosniff");
    response.setHeader("Referrer-Policy", "no-referrer");
    response.setHeader(
      "Content-Security-Policy",
      "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' blob: data:; connect-src 'self'; object-src 'none'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'",
    );

    let requestUrl;
    try {
      requestUrl = new URL(request.url || "/", "http://localhost");
    } catch {
      sendJson(response, 400, { error: "URL della richiesta non valida." });
      return;
    }

    if (requestUrl.pathname.startsWith("/api/")) {
      if (!isAllowedLocalRequest(request)) {
        sendJson(response, 403, { error: "Richiesta non autorizzata." });
        return;
      }
      if (requestUrl.pathname === "/api/status" && request.method === "GET") {
        sendJson(response, 200, { configured: Boolean(config) });
        return;
      }
      if (requestUrl.pathname === "/api/check" && request.method === "POST") {
        if (!config) {
          sendJson(response, 503, { error: "Configura HERMES_API_URL e HERMES_API_KEY nel file .env." });
          return;
        }
        try {
          await readRequestJson(request);
          const connected = await checkHermesConnection(config, fetchImpl);
          if (!connected) {
            sendJson(response, 502, { error: "Hermes non è raggiungibile: controlla indirizzo, porta, rete privata e chiave API." });
            return;
          }
          sendJson(response, 200, { connected: true });
        } catch (error) {
          sendJson(response, error.statusCode || 400, { error: error.message });
        }
        return;
      }
      if (requestUrl.pathname === "/api/chat" && request.method === "POST") {
        if (!config) {
          sendJson(response, 503, { error: "Configura HERMES_API_URL e HERMES_API_KEY nel file .env." });
          return;
        }
        try {
          const payload = await readRequestJson(request);
          const validationError = validateMessages(payload);
          if (validationError) {
            sendJson(response, 400, { error: validationError });
            return;
          }
          const upstream = await fetchImpl(`${config.baseUrl}/chat/completions`, {
            method: "POST",
            headers: {
              Authorization: `Bearer ${config.apiKey}`,
              "Content-Type": "application/json",
              Accept: "application/json",
            },
            body: JSON.stringify({
              model: "hermes-agent",
              messages: payload.messages,
              stream: false,
            }),
            signal: AbortSignal.timeout(CHAT_TIMEOUT_MS),
            redirect: "error",
          });
          if (!upstream.ok) {
            sendJson(response, 502, { error: `Hermes ha risposto con errore HTTP ${upstream.status}. Controlla la configurazione dell’agente.` });
            return;
          }
          const result = await readLimitedJson(upstream);
          sendJson(response, 200, result);
        } catch (error) {
          const statusCode = error.statusCode || (error.name === "TimeoutError" ? 504 : 502);
          sendJson(response, statusCode, {
            error: error.statusCode ? error.message : "Errore di connessione al relay o al servizio Hermes.",
          });
        }
        return;
      }
      sendJson(response, 404, { error: "Endpoint non trovato." });
      return;
    }

    if (request.method !== "GET" || !STATIC_FILES.has(requestUrl.pathname)) {
      response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Not found");
      return;
    }
    const [filename, contentType] = STATIC_FILES.get(requestUrl.pathname);
    try {
      const content = await readFile(resolve(PROJECT_DIR, filename));
      response.writeHead(200, {
        "Content-Type": contentType,
        "Cache-Control": "no-store",
      });
      response.end(content);
    } catch {
      response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Unable to load application files.");
    }
  });
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  if (!Number.isInteger(PORT) || PORT < 1 || PORT > 65535) {
    throw new Error("PORT deve essere un numero compreso tra 1 e 65535.");
  }
  const server = createAppServer();
  server.listen(PORT, "127.0.0.1", () => {
    console.log(`Apprendistato Studio disponibile su http://127.0.0.1:${PORT}`);
    if (!process.env.HERMES_API_URL || !process.env.HERMES_API_KEY) {
      console.log("Chat Hermes non configurata: copia .env.example in .env e inserisci endpoint e chiave.");
    }
  });
}
