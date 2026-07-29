import test from "node:test";
import assert from "node:assert/strict";
import { askHermes, resetConfigCache, parseCompletionBody } from "./hermes.js";

const CONFIG = { baseUrl: "https://api.example.test/v1", apiKey: "secret", model: "test-model" };

// loadConfig() memoises config.json for the lifetime of the module. Every test
// here counts fetch calls, so each must start from a cold cache — otherwise
// whichever test ran first supplies the config for the others and their counts
// come up one short.
test.beforeEach(() => {
  resetConfigCache();
});

// Provider responses are read with res.text() (see parseCompletionBody), so
// mocks return a body string rather than a parsed object.
function completion(content) {
  return {
    ok: true,
    status: 200,
    text: async () => JSON.stringify({ choices: [{ message: { content } }] }),
  };
}

function configFetch(response, seen, config = CONFIG) {
  return async (url, init) => {
    seen.push({ url, init });
    if (url === "/config.json") {
      return { ok: true, json: async () => config };
    }
    return response;
  };
}

test("askHermes calls configured OpenAI-compatible endpoint", async () => {
  const seen = [];
  global.fetch = configFetch(completion("ok"), seen);
  assert.equal(await askHermes([{ role: "user", content: "hi" }]), "ok");
  assert.equal(seen[0].url, "/config.json");
  const request = seen[1];
  assert.equal(request.url, "https://api.example.test/v1/chat/completions");
  assert.equal(request.init.headers.Authorization, "Bearer secret");
  assert.deepEqual(JSON.parse(request.init.body), {
    model: "test-model",
    messages: [{ role: "user", content: "hi" }],
  });
});

test("askHermes surfaces provider HTTP errors without retrying", async () => {
  const seen = [];
  global.fetch = configFetch({ ok: false, status: 500, text: async () => "boom" }, seen);
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /Provider 500/);
  // Exactly one config fetch + one provider call: an HTTP error is not retried.
  assert.equal(seen.length, 2);
  assert.equal(seen.filter((s) => s.url !== "/config.json").length, 1);
});

test("askHermes retries once on timeout", async () => {
  let providerCalls = 0;
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    providerCalls++;
    const err = new Error("timeout");
    err.name = "TimeoutError";
    throw err;
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /kịp thời/);
  assert.equal(providerCalls, 2);
});

test("askHermes caches config.json across calls", async () => {
  const seen = [];
  global.fetch = configFetch(completion("ok"), seen);
  await askHermes([{ role: "user", content: "one" }]);
  await askHermes([{ role: "user", content: "two" }]);
  assert.equal(seen.filter((s) => s.url === "/config.json").length, 1);
});

test("resetConfigCache forces config.json to be re-fetched", async () => {
  const seen = [];
  global.fetch = configFetch(completion("ok"), seen);
  await askHermes([{ role: "user", content: "one" }]);
  resetConfigCache();
  await askHermes([{ role: "user", content: "two" }]);
  assert.equal(seen.filter((s) => s.url === "/config.json").length, 2);
});

test("askHermes rejects a non-HTTPS baseUrl", async () => {
  const seen = [];
  global.fetch = configFetch(completion("ok"), seen, {
    ...CONFIG,
    baseUrl: "http://api.example.test/v1",
  });
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /baseUrl HTTPS/);
});

// ---------------------------------------------------------------------------
// Response-body parsing.
//
// These cases are taken verbatim from a real OpenAI-compatible router
// (an OpenRouter-backed nginx front end) that answers a NON-streaming request
// with Content-Type: text/event-stream, ~55 bytes of whitespace padding, the
// completion object, and a trailing "data: [DONE]" sentinel. res.json() throws
// a SyntaxError on that trailing data, so the reply never reached the pane.

test("parseCompletionBody: plain JSON body", () => {
  const body = JSON.stringify({ choices: [{ message: { content: "hi" } }] });
  assert.equal(parseCompletionBody(body), "hi");
});

test("parseCompletionBody: leading whitespace padding", () => {
  const body = "\n   \n  " + JSON.stringify({ choices: [{ message: { content: "hi" } }] });
  assert.equal(parseCompletionBody(body), "hi");
});

test("parseCompletionBody: trailing data: [DONE] after the object", () => {
  const body =
    "     \n\n" + JSON.stringify({ choices: [{ message: { content: "OK." } }] }) + "data: [DONE]";
  assert.equal(parseCompletionBody(body), "OK.");
});

test("parseCompletionBody: SSE-framed single object", () => {
  const body =
    "data: " + JSON.stringify({ choices: [{ message: { content: "framed" } }] }) + "\n\ndata: [DONE]\n";
  assert.equal(parseCompletionBody(body), "framed");
});

test("parseCompletionBody: streamed delta chunks are concatenated", () => {
  const chunk = (c) => "data: " + JSON.stringify({ choices: [{ delta: { content: c } }] }) + "\n\n";
  const body = chunk("Xin ") + chunk("chào") + chunk("!") + "data: [DONE]\n";
  assert.equal(parseCompletionBody(body), "Xin chào!");
});

test("parseCompletionBody: empty content is preserved, not treated as failure", () => {
  const body = JSON.stringify({ choices: [{ message: { content: "" } }] });
  assert.equal(parseCompletionBody(body), "");
});

test("parseCompletionBody: unparseable body throws a body-shaped error", () => {
  assert.throws(() => parseCompletionBody("<html>502 Bad Gateway</html>"), /không đọc được/i);
});

test("parseCompletionBody: valid JSON with neither choices nor error throws", () => {
  assert.throws(() => parseCompletionBody(JSON.stringify({ id: "x", object: "chat.completion" })), /payload/i);
});

test("parseCompletionBody: a bare string error field is surfaced too", () => {
  assert.throws(() => parseCompletionBody(JSON.stringify({ error: "nope" })), /Provider báo lỗi: nope/);
});

// ---------------------------------------------------------------------------
// Error classification. A CORS block rejects fetch with a TypeError instantly;
// reporting that as "provider didn't answer in time" sends everyone hunting a
// timeout that never happened.

test("askHermes reports a network/CORS failure as such, not as a timeout", async () => {
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    throw new TypeError("Load failed");
  };
  await assert.rejects(
    () => askHermes([{ role: "user", content: "hi" }]),
    (err) => {
      assert.match(err.message, /CORS|mạng/i);
      assert.doesNotMatch(err.message, /kịp thời/);
      return true;
    }
  );
});

test("askHermes still reports an actual timeout as a timeout", async () => {
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    const err = new Error("timeout");
    err.name = "TimeoutError";
    throw err;
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /kịp thời/);
});

test("parseCompletionBody: a structured provider error is surfaced verbatim", () => {
  // Real shape observed from the router: padding + error object + sentinel.
  const body =
    "   \n" +
    JSON.stringify({ error: { message: "The operation was aborted", code: 504 } }) +
    "data: [DONE]";
  assert.throws(() => parseCompletionBody(body), (err) => {
    assert.match(err.message, /The operation was aborted/);
    assert.match(err.message, /504/);
    // Must read as the provider's own error, not as a parser complaint.
    assert.doesNotMatch(err.message, /bad payload/i);
    return true;
  });
});
