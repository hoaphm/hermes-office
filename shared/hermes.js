/* global fetch, AbortSignal */
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

async function callHermes(messages, idempotencyKey, timeoutMs) {
  // Note: only Authorization, Content-Type, and Idempotency-Key are allowed by
  // the API server's CORS policy. Don't add other custom headers or the browser
  // preflight will fail ("Failed to fetch"). Conversation continuity comes from
  // sending the full message history each call.
  const headers = { "Content-Type": "application/json" };
  if (idempotencyKey) headers["Idempotency-Key"] = idempotencyKey;

  const res = await fetch(ENDPOINT, {
    method: "POST",
    headers,
    body: JSON.stringify({ model: "hermes-agent", messages }),
    signal: AbortSignal.timeout(timeoutMs),
  });
  if (!res.ok) {
    throw new Error(`Hermes ${res.status}: ${await res.text()}`);
  }
  const data = await res.json();
  const choice = data && data.choices && data.choices[0];
  if (!choice || !choice.message)
    throw new Error(`Hermes bad payload: ${JSON.stringify(data).slice(0, 200)}`);
  return choice.message.content;
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
