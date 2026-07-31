/* global fetch, setTimeout, clearTimeout */
// Shared OpenAI-compatible API client for Word and Excel add-ins.
// Provider config is served as /config.json alongside each production bundle.

const DEFAULT_TIMEOUT_MS = 60000;
let cachedConfig;

function isTimeoutOrNetworkError(err) {
  return (
    err instanceof TypeError ||
    (err && (err.name === "AbortError" || err.name === "TimeoutError"))
  );
}

// Office for Mac's WKWebView rejects an AbortSignal passed to fetch() with
// "The string did not match the expected pattern" (both AbortSignal.timeout()
// and AbortController.signal). Race the fetch against a timer instead — no
// signal is passed to fetch at all.
function fetchWithTimeout(url, init, timeoutMs) {
  let id;
  const timer = new Promise((_, reject) => {
    id = setTimeout(() => {
      const e = new Error("timeout");
      e.name = "TimeoutError";
      reject(e);
    }, timeoutMs);
  });
  return Promise.race([fetch(url, init), timer]).finally(() => clearTimeout(id));
}

async function loadConfig() {
  if (cachedConfig) return cachedConfig;
  let data;
  let configUrl;
  try {
    // Word for Mac's WKWebView rejects fetch("/config.json") with
    // "The string did not match the expected pattern". Resolve it explicitly
    // against the add-in's current HTTPS origin instead.
    configUrl =
      typeof window === "undefined"
        ? "/config.json"
        : new URL("/config.json", window.location.href).href;
  } catch (err) {
    throw new Error(`[config-url] ${err.message || err}`);
  }
  try {
    const res = await fetch(configUrl);
    if (!res.ok) throw new Error(`config ${res.status}`);
    data = await res.json();
  } catch (err) {
    throw new Error(`[config-fetch:${configUrl}] ${err.message || err}`);
  }
  const baseUrl = String(data.baseUrl || "").replace(/\/+$/, "");
  const apiKey = String(data.apiKey || "");
  const model = String(data.model || "");
  if (!/^https:\/\//i.test(baseUrl) || !apiKey || !model) {
    throw new Error("config.json cần có baseUrl HTTPS, apiKey, model.");
  }
  cachedConfig = { baseUrl, apiKey, model };
  return cachedConfig;
}

export async function callApi(
  messages,
  { idempotencyKey, timeoutMs = DEFAULT_TIMEOUT_MS } = {}
) {
  let phase = "load-config";
  try {
    const config = await loadConfig();
    phase = "build-headers";
    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${config.apiKey}`,
    };
    if (idempotencyKey) headers["Idempotency-Key"] = idempotencyKey;
    phase = "build-url";
    const url = `${config.baseUrl}/chat/completions`;
    phase = "fetch";
    const res = await fetchWithTimeout(
      url,
      {
        method: "POST",
        headers,
        body: JSON.stringify({ model: config.model, messages }),
      },
      timeoutMs,
    );
    phase = "read-response";
    const body = await res.text();
    if (!res.ok) throw new Error(`Provider ${res.status}: ${body}`);
    // Some OpenAI-compatible routers append an SSE terminator (`data: [DONE]`)
    // after an otherwise normal JSON completion. Response.json() rejects that
    // body; parse its leading JSON object instead. Standard JSON still works.
    const marker = body.indexOf("data: [DONE]");
    const data = JSON.parse((marker === -1 ? body : body.slice(0, marker)).trim());
    const choice = data && data.choices && data.choices[0];
    if (!choice || !choice.message) {
      throw new Error(`Provider bad payload: ${JSON.stringify(data).slice(0, 200)}`);
    }
    return choice.message.content;
  } catch (err) {
    const tagged = new Error(`[call-api:${phase}] ${err.message || err}`);
    tagged.name = err && err.name ? err.name : "Error";
    throw tagged;
  }
}

// Kept name for minimal caller diff and backward compatibility.
export async function askHermes(
  messages,
  { idempotencyKey, timeoutMs = DEFAULT_TIMEOUT_MS } = {}
) {
  try {
    return await callApi(messages, { idempotencyKey, timeoutMs });
  } catch (err) {
    if (!isTimeoutOrNetworkError(err)) {
      throw new Error(`[api-first] ${err.message || err}`);
    }
    try {
      return await callApi(messages, { idempotencyKey, timeoutMs });
    } catch (err2) {
      if (isTimeoutOrNetworkError(err2)) {
        throw new Error("Provider không phản hồi kịp thời, vui lòng thử lại.");
      }
      throw new Error(`[api-retry] ${err2.message || err2}`);
    }
  }
}
