/* global fetch, setTimeout, clearTimeout */
// Shared chat client for Word and Excel add-ins.
//
// Every call goes to the Local Gateway (the Caddy instance serving this bundle)
// at same-origin /v1/chat/completions — never to the Provider directly. The
// Gateway holds the API key and forwards upstream, so nothing here authenticates
// and no Authorization header is sent. See docs/adr/0001 and 0002.

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

// Word for Mac's WKWebView rejects a root-relative fetch("/…") with "The string
// did not match the expected pattern". Resolve against the add-in's own origin
// instead — which is the Local Gateway, since it served this bundle.
function gatewayUrl(path) {
  return typeof window === "undefined"
    ? path
    : new URL(path, window.location.href).href;
}

async function loadConfig() {
  if (cachedConfig) return cachedConfig;
  let data;
  let configUrl;
  try {
    configUrl = gatewayUrl("/config.json");
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
  const model = String(data.model || "");
  if (!model) {
    throw new Error("config.json cần có model. Chạy: npm run setup");
  }
  cachedConfig = { model };
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
    // No Authorization: the Local Gateway attaches the key on the way upstream.
    const headers = { "Content-Type": "application/json" };
    if (idempotencyKey) headers["Idempotency-Key"] = idempotencyKey;
    phase = "build-url";
    const url = gatewayUrl("/v1/chat/completions");
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
