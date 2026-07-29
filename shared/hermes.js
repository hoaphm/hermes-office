/* global fetch, AbortSignal, AbortController, setTimeout, clearTimeout */
// Shared OpenAI-compatible API client for Word and Excel add-ins.
// Provider config is served as /config.json alongside each production bundle.

const DEFAULT_TIMEOUT_MS = 60000;
// config.json is fetched once per task-pane session and memoised. Reloading
// the pane picks up an edited config; resetConfigCache() forces a re-fetch
// without one, and is what keeps this module's tests independent of each
// other (without it, whichever test ran first decided the config for all the
// rest, and their fetch-call counts silently drifted).
let cachedConfig;

export function resetConfigCache() {
  cachedConfig = undefined;
}

// fetch() rejects with a TypeError when the request never completed at the
// network layer — DNS failure, connection refused, TLS rejected, or (by far
// the most common here) the browser blocking a cross-origin response because
// the provider sent no CORS headers. It carries no detail by design.
function isNetworkError(err) {
  return err instanceof TypeError;
}

function isTimeoutError(err) {
  return !!err && (err.name === "AbortError" || err.name === "TimeoutError");
}

function isRetriable(err) {
  return isNetworkError(err) || isTimeoutError(err);
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

// Turn a chat-completions response BODY into the assistant's text.
//
// res.json() is not enough. OpenAI-compatible routers in the wild answer a
// non-streaming request with Content-Type: text/event-stream, whitespace
// padding to defeat proxy buffering, and a trailing "data: [DONE]" sentinel
// after the completion object — all of which make JSON.parse throw on the
// whole body. Handled shapes:
//
//   {...}                       plain JSON
//   "   \n {...}"               padded
//   "{...}data: [DONE]"         object plus sentinel, no framing
//   "data: {...}\n\ndata: [DONE]"   SSE-framed single object
//   "data: {delta}\n\ndata: {delta}…"  true streamed chunks, concatenated
export function parseCompletionBody(text) {
  const body = String(text ?? "").trim();
  if (!body) throw new Error("Provider trả về body rỗng.");

  const contentOf = (obj) => {
    const choice = obj && obj.choices && obj.choices[0];
    if (!choice) return undefined;
    if (choice.message && typeof choice.message.content === "string") return choice.message.content;
    if (choice.delta && typeof choice.delta.content === "string") return choice.delta.content;
    return undefined;
  };

  // SSE framing: collect every `data:` payload and concatenate their content,
  // which handles both a single framed object and real streamed deltas.
  if (/^data:\s/m.test(body)) {
    const parts = [];
    let sawPayload = false;
    for (const line of body.split(/\r?\n/)) {
      const m = line.match(/^data:\s*(.*)$/);
      if (!m) continue;
      const payload = m[1].trim();
      if (!payload || payload === "[DONE]") continue;
      try {
        const c = contentOf(JSON.parse(payload));
        sawPayload = true;
        if (typeof c === "string") parts.push(c);
      } catch {
        /* skip a frame we cannot parse */
      }
    }
    if (sawPayload) return parts.join("");
  }

  // Otherwise: parse the FIRST complete JSON object and ignore any trailing
  // sentinel. JSON.parse on the whole string would throw on "data: [DONE]".
  let obj;
  try {
    obj = JSON.parse(body);
  } catch {
    const start = body.indexOf("{");
    const end = start === -1 ? -1 : matchingBrace(body, start);
    if (end === -1) {
      throw new Error(`Provider trả về body không đọc được: ${body.slice(0, 200)}`);
    }
    try {
      obj = JSON.parse(body.slice(start, end + 1));
    } catch {
      throw new Error(`Provider trả về body không đọc được: ${body.slice(0, 200)}`);
    }
  }

  // A structured error can arrive with HTTP 200 (the router reports upstream
  // failures in the body). Surface the provider's own words — "bad payload"
  // tells the user nothing about an upstream 504.
  if (obj && obj.error) {
    const e = obj.error;
    const msg = typeof e === "string" ? e : e.message || JSON.stringify(e);
    const code = e && e.code !== undefined ? ` (code ${e.code})` : "";
    throw new Error(`Provider báo lỗi${code}: ${msg}`);
  }

  const content = contentOf(obj);
  if (typeof content !== "string") {
    throw new Error(`Provider bad payload: ${JSON.stringify(obj).slice(0, 200)}`);
  }
  return content;
}

// Index of the `}` closing the `{` at `start`, string-aware so braces inside
// JSON string values do not throw off the depth count.
function matchingBrace(text, start) {
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === "{") depth++;
    else if (ch === "}" && --depth === 0) return i;
  }
  return -1;
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
    // Read as text, not res.json(): see parseCompletionBody for the padded /
    // SSE-framed shapes real routers return for a non-streaming request.
    return parseCompletionBody(await res.text());
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
    if (!isRetriable(err)) throw err;
    try {
      return await callApi(messages, { idempotencyKey, timeoutMs });
    } catch (err2) {
      // Distinguish the two. They used to share one "didn't respond in time"
      // message, which is actively misleading for a CORS block: that fails
      // instantly, and the message sends you hunting a timeout that never
      // happened.
      if (isTimeoutError(err2)) {
        throw new Error(
          `Provider không phản hồi kịp thời (quá ${Math.round(timeoutMs / 1000)}s). Thử lại, hoặc chọn model nhanh hơn.`
        );
      }
      if (isNetworkError(err2)) {
        const host = cachedConfig ? cachedConfig.baseUrl : "provider";
        throw new Error(
          `Không gọi được provider (${host}). Request bị chặn ở tầng mạng — thường là provider thiếu CORS headers ` +
            `(phải trả Access-Control-Allow-Origin cho https://localhost:8643 và cho OPTIONS qua mà không cần API key), ` +
            `hoặc mất mạng / sai baseUrl. Xem mục "Provider CORS" trong README.`
        );
      }
      throw err2;
    }
  }
}
