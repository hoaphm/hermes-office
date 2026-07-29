/* global fetch, AbortSignal, AbortController, setTimeout, clearTimeout */
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

function makeTimeout(timeoutMs) {
  if (typeof AbortSignal !== "undefined" && typeof AbortSignal.timeout === "function") {
    return { signal: AbortSignal.timeout(timeoutMs), cancel() {} };
  }
  if (typeof AbortController === "undefined") {
    return { signal: undefined, cancel() {} };
  }
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);
  return { signal: controller.signal, cancel: () => clearTimeout(id) };
}

async function loadConfig() {
  if (cachedConfig) return cachedConfig;
  let data;
  try {
    const res = await fetch("/config.json");
    if (!res.ok) throw new Error(`config ${res.status}`);
    data = await res.json();
  } catch {
    throw new Error("Không tải được config.json. Chạy scripts/setup để thiết lập provider.");
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
  const config = await loadConfig();
  const timeout = makeTimeout(timeoutMs);
  try {
    const headers = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${config.apiKey}`,
    };
    if (idempotencyKey) headers["Idempotency-Key"] = idempotencyKey;
    const res = await fetch(`${config.baseUrl}/chat/completions`, {
      method: "POST",
      headers,
      body: JSON.stringify({ model: config.model, messages }),
      signal: timeout.signal,
    });
    if (!res.ok) throw new Error(`Provider ${res.status}: ${await res.text()}`);
    const data = await res.json();
    const choice = data && data.choices && data.choices[0];
    if (!choice || !choice.message) {
      throw new Error(`Provider bad payload: ${JSON.stringify(data).slice(0, 200)}`);
    }
    return choice.message.content;
  } finally {
    timeout.cancel();
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
    if (!isTimeoutOrNetworkError(err)) throw err;
    try {
      return await callApi(messages, { idempotencyKey, timeoutMs });
    } catch (err2) {
      if (isTimeoutOrNetworkError(err2)) {
        throw new Error("Provider không phản hồi kịp thời, vui lòng thử lại.");
      }
      throw err2;
    }
  }
}
