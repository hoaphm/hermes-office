/* global fetch, setTimeout, clearTimeout */
// Shared chat client for Word and Excel add-ins.
//
// Every call goes to the Local Gateway (the Caddy instance serving this bundle)
// at same-origin /v1/chat/completions — never to the Provider directly. The
// Gateway holds the API key and forwards upstream, so nothing here authenticates
// and no Authorization header is sent. See docs/adr/0001 and 0002.

const DEFAULT_TIMEOUT_MS = 60000;
let cachedConfig;

// Set by the Local Gateway's own `handle_errors` block on a response it
// generated itself, rather than one it proxied back from the Provider. It is
// the only way to tell "the Gateway could not reach the Provider" (an Upstream
// Hop failure — what a change of network looks like) from "the Provider itself
// answered 502", which are otherwise the same status code. See ADR-0005.
export const GATEWAY_ERROR_HEADER = "X-Hermes-Gateway-Error";

// Facts about a failure, recorded where they are known and read by
// shared/failures.js, which alone decides which Disclosed Failure they add up
// to. Transport records; policy classifies. Never a message — the whole point
// is that no upstream text travels with these.
function withDetail(err, detail) {
  err.hermes = detail;
  return err;
}

// Re-wrapping an error for its stage tag would otherwise drop the facts the
// throw site recorded, leaving failures.js nothing to classify but a string.
function carry(next, previous) {
  if (previous && previous.hermes) next.hermes = previous.hermes;
  return next;
}

// A stubbed fetch in tests need not supply Headers, and neither does every
// WebView error path.
function gatewayGenerated(res) {
  const headers = res && res.headers;
  if (!headers || typeof headers.get !== "function") return false;
  return headers.get(GATEWAY_ERROR_HEADER) != null;
}

// Match on `name`, not `instanceof`: callApi re-wraps every failure in a plain
// Error carrying only the original's name, so an `err instanceof TypeError`
// test here could never be true and genuine network failures went un-retried.
// A network failure is exactly what a retry is for.
function isTimeoutOrNetworkError(err) {
  const name = err && err.name;
  return name === "TypeError" || name === "AbortError" || name === "TimeoutError";
}

// A unique-per-turn key so a Provider that honours Idempotency-Key can dedupe
// a timeout retry instead of running (and billing) the completion twice. Fresh
// per call — NOT content-derived, so a user who re-asks the same question gets
// a genuinely new answer rather than the cached first one. Cheap, collision is
// not a correctness problem (worst case: a dedup that never fires).
let turnSeq = 0;
function nextTurnKey() {
  turnSeq += 1;
  return `${Date.now().toString(36)}-${turnSeq.toString(36)}`;
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
    throw withDetail(new Error(`[config-url] ${err.message || err}`), {
      stage: "config-url",
    });
  }
  let status;
  try {
    const res = await fetch(configUrl);
    status = res.status;
    if (!res.ok) throw new Error(`config ${res.status}`);
    data = await res.json();
  } catch (err) {
    // No status means the request never got an answer: on the Local Hop that
    // is the Gateway not listening at all.
    throw withDetail(new Error(`[config-fetch:${configUrl}] ${err.message || err}`), {
      stage: "config-fetch",
      status,
    });
  }
  const model = String(data.model || "");
  if (!model) {
    throw withDetail(new Error("config.json cần có model. Chạy: npm run setup"), {
      stage: "config-model",
    });
  }
  // Custom Functions reach the Provider from a worksheet cell, outside the
  // Apply boundary — a workbook that merely CONTAINS `=HERMES.*` fires them on
  // open. Off unless config.json opts in; config.json is a local file, so a
  // document can never flip this on. See CONTEXT.md and the README.
  cachedConfig = { model, customFunctions: data.customFunctions === true };
  return cachedConfig;
}

// Whether `=HERMES.*` worksheet functions are permitted to call the Provider.
export async function customFunctionsEnabled() {
  const config = await loadConfig();
  return config.customFunctions;
}

export async function callApi(
  messages,
  { idempotencyKey, timeoutMs = DEFAULT_TIMEOUT_MS } = {}
) {
  let phase = "load-config";
  const detail = { stage: phase };
  try {
    const config = await loadConfig();
    phase = "build-headers";
    // No Authorization: the Local Gateway attaches the key on the way upstream.
    const headers = { "Content-Type": "application/json" };
    if (idempotencyKey) headers["Idempotency-Key"] = idempotencyKey;
    phase = "build-url";
    const url = gatewayUrl("/v1/chat/completions");
    phase = "fetch";
    detail.stage = phase;
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
    detail.stage = phase;
    detail.status = res.status;
    detail.gateway = gatewayGenerated(res);
    const body = await res.text();
    if (!res.ok) throw new Error(`Provider ${res.status}: ${body}`);
    // Some OpenAI-compatible routers append an SSE terminator (`data: [DONE]`)
    // after an otherwise normal JSON completion. Response.json() rejects that
    // body; parse its leading JSON object instead. Standard JSON still works.
    const marker = body.indexOf("data: [DONE]");
    let data;
    try {
      data = JSON.parse((marker === -1 ? body : body.slice(0, marker)).trim());
    } catch {
      detail.badPayload = true;
      throw new Error(`Provider bad payload: ${body.slice(0, 200)}`);
    }
    const choice = data && data.choices && data.choices[0];
    if (!choice || !choice.message) {
      detail.badPayload = true;
      throw new Error(`Provider bad payload: ${JSON.stringify(data).slice(0, 200)}`);
    }
    return choice.message.content;
  } catch (err) {
    const tagged = new Error(`[call-api:${phase}] ${err.message || err}`);
    tagged.name = err && err.name ? err.name : "Error";
    // loadConfig already recorded the more specific stage; do not overwrite it.
    withDetail(tagged, (err && err.hermes) || { ...detail, name: tagged.name });
    throw tagged;
  }
}

// ---- Hop check --------------------------------------------------------------
//
// Probe the two hops separately so a user can find out which one is broken
// without having to provoke a failure with a real question. Costs no tokens:
// the Local Hop is the config fetch the task pane already makes, and the
// Upstream Hop is a GET the Provider need not even like — *any* answer that
// the Gateway did not generate itself proves the Gateway reached it, which is
// the only thing being asked. See CONTEXT.md for the two hops.

async function probeLocalHop(timeoutMs) {
  let status;
  try {
    const res = await fetchWithTimeout(gatewayUrl("/config.json"), {}, timeoutMs);
    status = res.status;
    if (!res.ok) throw new Error(`config ${res.status}`);
    const data = await res.json();
    if (!String(data.model || "")) {
      throw withDetail(new Error("[config-model] no model"), { stage: "config-model" });
    }
    return { ok: true, error: null };
  } catch (err) {
    if (err && err.hermes) return { ok: false, error: err };
    return {
      ok: false,
      error: withDetail(new Error(`[config-fetch:probe] ${err.message || err}`), {
        stage: "config-fetch",
        status,
        name: err && err.name,
      }),
    };
  }
}

async function probeUpstreamHop(timeoutMs) {
  try {
    const res = await fetchWithTimeout(gatewayUrl("/v1/models"), { method: "GET" }, timeoutMs);
    if (!gatewayGenerated(res)) return { ok: true, error: null };
    return {
      ok: false,
      error: withDetail(new Error("[call-api:probe] gateway did not reach the Provider"), {
        stage: "read-response",
        status: res.status,
        gateway: true,
      }),
    };
  } catch (err) {
    return {
      ok: false,
      error: withDetail(new Error(`[call-api:probe] ${err.message || err}`), {
        stage: "fetch",
        name: err && err.name,
      }),
    };
  }
}

/**
 * Test the Local Hop and the Upstream Hop, in that order. The Upstream Hop is
 * not probed when the Local Hop is down: every upstream request travels
 * through the Gateway, so the answer would say nothing about the Provider.
 * @param {{timeoutMs?: number}} [opts]
 */
export async function checkHops({ timeoutMs = 15000 } = {}) {
  const local = await probeLocalHop(timeoutMs);
  if (!local.ok) return { local, upstream: { ok: false, error: null, skipped: true } };
  return { local, upstream: await probeUpstreamHop(timeoutMs) };
}

// Kept name for minimal caller diff and backward compatibility.
export async function askHermes(
  messages,
  { idempotencyKey, timeoutMs = DEFAULT_TIMEOUT_MS } = {}
) {
  // BUG-03: chat turns (Word/Excel task panes) used to send no Idempotency-Key,
  // so a timeout retry below ran a SECOND full completion with nothing for the
  // Provider to dedupe against. Generate one per logical turn — the retry must
  // carry the SAME key as the first attempt, while a fresh user turn gets a
  // fresh key (a re-ask must produce a new answer, not a cached one). Callers
  // that want content-derived keys (Custom Functions, for recalc-storm dedup)
  // pass their own.
  const key = idempotencyKey || nextTurnKey();
  try {
    return await callApi(messages, { idempotencyKey: key, timeoutMs });
  } catch (err) {
    if (!isTimeoutOrNetworkError(err)) {
      throw carry(new Error(`[api-first] ${err.message || err}`), err);
    }
    try {
      return await callApi(messages, { idempotencyKey: key, timeoutMs });
    } catch (err2) {
      if (isTimeoutOrNetworkError(err2)) {
        throw withDetail(
          new Error("Provider không phản hồi kịp thời, vui lòng thử lại."),
          { stage: "fetch", name: "TimeoutError" },
        );
      }
      throw carry(new Error(`[api-retry] ${err2.message || err2}`), err2);
    }
  }
}
