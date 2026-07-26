/* global fetch, AbortSignal, AbortController, setTimeout, clearTimeout */
// Canonical Hermes client. word/src/shared/hermes.js and
// excel/src/shared/hermes.js both re-export this module — there is no npm
// workspace between the two add-ins (no registry publish), so this
// relative-path shared/ folder at the repo root is how the implementation is
// deduped instead.
// Caddy (https://localhost:8643) terminates TLS and injects the Authorization
// header, so we never send the API key from here.

const ENDPOINT = "/v1/chat/completions";
const DEFAULT_TIMEOUT_MS = 60000;

function isTimeoutOrNetworkError(err) {
  return (
    err instanceof TypeError || // fetch network failure (offline, DNS, blocked CORS preflight, etc.)
    (err && (err.name === "AbortError" || err.name === "TimeoutError"))
  );
}

// AbortSignal.timeout() is missing on the older WebViews some Office builds
// still embed (and core-js does not polyfill it), where calling it throws a
// TypeError that isTimeoutOrNetworkError misreads as a network blip — so the
// user got "Hermes did not respond in time" for what was really an unsupported
// API. Fall back to AbortController + setTimeout, whose abort surfaces as the
// AbortError we already handle. Returns a canceller so a completed request
// doesn't leave a timer armed.
function makeTimeout(timeoutMs) {
  if (typeof AbortSignal !== "undefined" && typeof AbortSignal.timeout === "function") {
    return { signal: AbortSignal.timeout(timeoutMs), cancel() {} };
  }
  if (typeof AbortController === "undefined") {
    return { signal: undefined, cancel() {} }; // no abort support — rely on the server
  }
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);
  return { signal: controller.signal, cancel: () => clearTimeout(id) };
}

async function callHermes(messages, idempotencyKey, timeoutMs) {
  // Note: only Authorization, Content-Type, and Idempotency-Key are allowed by
  // the API server's CORS policy. Don't add other custom headers or the browser
  // preflight will fail ("Failed to fetch"). Conversation continuity comes from
  // sending the full message history each call.
  const headers = { "Content-Type": "application/json" };
  if (idempotencyKey) headers["Idempotency-Key"] = idempotencyKey;

  const timeout = makeTimeout(timeoutMs);
  try {
    const res = await fetch(ENDPOINT, {
      method: "POST",
      headers,
      body: JSON.stringify({ model: "hermes-agent", messages }),
      signal: timeout.signal,
    });
    if (!res.ok) {
      throw new Error(`Hermes ${res.status}: ${await res.text()}`);
    }
    const data = await res.json();
    const choice = data && data.choices && data.choices[0];
    if (!choice || !choice.message)
      throw new Error(`Hermes bad payload: ${JSON.stringify(data).slice(0, 200)}`);
    return choice.message.content;
  } finally {
    // Cancelled only once the body has been fully read — aborting earlier
    // would tear down the stream mid-parse.
    timeout.cancel();
  }
}

export async function askHermes(
  messages,
  { idempotencyKey, timeoutMs = DEFAULT_TIMEOUT_MS } = {}
) {
  try {
    return await callHermes(messages, idempotencyKey, timeoutMs);
  } catch (err) {
    // Retry exactly once, and only for a transient network failure or a
    // request timeout. An HTTP error response (res.ok === false) is a real
    // server-side answer, not a blip, so it is NOT retried.
    if (!isTimeoutOrNetworkError(err)) throw err;
    try {
      return await callHermes(messages, idempotencyKey, timeoutMs);
    } catch (err2) {
      if (isTimeoutOrNetworkError(err2)) {
        throw new Error(
          "Hermes không phản hồi kịp thời, vui lòng thử lại. / Hermes did not respond in time — please try again."
        );
      }
      throw err2;
    }
  }
}
