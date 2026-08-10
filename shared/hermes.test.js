import test from "node:test";
import assert from "node:assert/strict";
import { askHermes, callApi, GATEWAY_ERROR_HEADER } from "./hermes.js";

const CONFIG = { name: "Custom", model: "test-model" };

function configFetch(response, seen) {
  return async (url, init) => {
    seen.push({ url, init });
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    return response;
  };
}

// Must run first: loadConfig caches on success, and every later test relies on
// that cache being populated.
test("callApi rejects a config.json without a model", async () => {
  global.fetch = async () => ({ ok: true, json: async () => ({ name: "Custom" }) });
  await assert.rejects(() => callApi([{ role: "user", content: "hi" }]), /cần có model/);
});

test("askHermes posts to the Local Gateway, not to a provider URL", async () => {
  const seen = [];
  global.fetch = configFetch({
    ok: true,
    status: 200,
    text: async () => JSON.stringify({ choices: [{ message: { content: "ok" } }] }),
  }, seen);
  assert.equal(await askHermes([{ role: "user", content: "hi" }]), "ok");
  const request = seen[1];
  assert.equal(request.url, "/v1/chat/completions");
  assert.deepEqual(JSON.parse(request.init.body), {
    model: "test-model",
    messages: [{ role: "user", content: "hi" }],
  });
});

// ADR-0002: the key lives in the Caddyfile. If this ever fails, a credential is
// being served to the WebView again.
test("callApi sends no Authorization header", async () => {
  let seenInit;
  global.fetch = async (url, init) => {
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    seenInit = init;
    return { ok: true, text: async () => JSON.stringify({ choices: [{ message: { content: "ok" } }] }) };
  };
  assert.equal(await callApi([{ role: "user", content: "hi" }]), "ok");
  assert.equal(seenInit.headers.Authorization, undefined);
  assert.equal(JSON.stringify(seenInit).toLowerCase().includes("bearer"), false);
});

test("callApi accepts a JSON completion followed by SSE done marker", async () => {
  global.fetch = async () => ({
    ok: true,
    text: async () => `${JSON.stringify({ choices: [{ message: { content: "ok" } }] })}data: [DONE]`,
  });
  assert.equal(await callApi([{ role: "user", content: "hi" }]), "ok");
});

test("askHermes surfaces provider HTTP errors without retrying", async () => {
  let callCount = 0;
  global.fetch = async (url) => {
    callCount++;
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    return { ok: false, status: 500, text: async () => "boom" };
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /Provider 500/);
  // callCount is 1 (not 2): config is cached from the previous test, so only
  // the API call is counted. The point is NO retry on HTTP error.
  assert.equal(callCount, 1);
});

test("callApi passes no AbortSignal to fetch (Office WebView compatibility)", async () => {
  let seenInit;
  global.fetch = async (url, init) => {
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    seenInit = init;
    return { ok: true, text: async () => JSON.stringify({ choices: [{ message: { content: "ok" } }] }) };
  };
  assert.equal(await callApi([{ role: "user", content: "hi" }]), "ok");
  assert.equal(seenInit.signal, undefined, "fetch must receive no signal");
});

// A network failure rejects fetch with a TypeError. callApi re-wraps it, so the
// retry check must match on `name` — an `instanceof TypeError` test there was
// always false and no network failure was ever retried.
test("askHermes retries once on a network error", async () => {
  let apiCalls = 0;
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    apiCalls++;
    throw new TypeError("Failed to fetch");
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]));
  assert.equal(apiCalls, 2);
});

test("askHermes retries once on timeout", async () => {
  let callCount = 0;
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, json: async () => CONFIG };
    callCount++;
    const err = new Error("timeout");
    err.name = "TimeoutError";
    throw err;
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /kịp thời/);
  assert.equal(callCount, 2);
});

// ---- Hop check --------------------------------------------------------------
// The point of the check is that the two hops are probed separately, and that
// the upstream probe costs nothing — see ADR-0005.

test("checkHops probes the Upstream Hop with a GET that spends no tokens", async () => {
  const seen = [];
  global.fetch = async (url, init) => {
    seen.push({ url, init });
    if (url === "/config.json") return { ok: true, status: 200, json: async () => CONFIG };
    return { ok: true, status: 200, headers: new Map() };
  };
  const mod = await import("./hermes.js?hop=ok");
  const result = await mod.checkHops();
  assert.deepEqual(
    { local: result.local.ok, upstream: result.upstream.ok },
    { local: true, upstream: true },
  );
  const probe = seen[seen.length - 1];
  assert.equal(probe.url, "/v1/models");
  assert.equal(probe.init.method, "GET");
  assert.equal(probe.init.body, undefined, "a probe must not post a completion");
});

// Any answer the Gateway did not write itself proves it reached the Provider,
// whatever the Provider thought of the request.
test("checkHops counts an unmarked Provider rejection as a reachable Upstream Hop", async () => {
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, status: 200, json: async () => CONFIG };
    return { ok: false, status: 404, headers: new Map() };
  };
  const mod = await import("./hermes.js?hop=404");
  assert.equal((await mod.checkHops()).upstream.ok, true);
});

test("checkHops reports an unreachable Upstream Hop only on the Gateway's own mark", async () => {
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, status: 200, json: async () => CONFIG };
    return { ok: false, status: 502, headers: new Map([[GATEWAY_ERROR_HEADER, "502"]]) };
  };
  const mod = await import("./hermes.js?hop=marked");
  const { upstream } = await mod.checkHops();
  assert.equal(upstream.ok, false);
  assert.equal(upstream.error.hermes.gateway, true);
});

test("checkHops skips the Upstream Hop when the Local Hop is down", async () => {
  let calls = 0;
  global.fetch = async () => {
    calls++;
    throw new TypeError("Failed to fetch");
  };
  const mod = await import("./hermes.js?hop=down");
  const result = await mod.checkHops();
  assert.equal(result.local.ok, false);
  assert.equal(result.upstream.skipped, true);
  assert.equal(calls, 1, "no upstream probe travels through a Gateway that is not there");
});

// ---- Custom Function gate ---------------------------------------------------
// loadConfig caches per module instance, so each case imports a fresh copy via
// a distinct query string rather than fighting the cache.

test("customFunctionsEnabled: off when config.json says nothing", async () => {
  global.fetch = async () => ({ ok: true, json: async () => ({ model: "m" }) });
  const mod = await import("./hermes.js?cf=absent");
  assert.equal(await mod.customFunctionsEnabled(), false);
});

test("customFunctionsEnabled: on only for a literal true", async () => {
  global.fetch = async () => ({
    ok: true,
    json: async () => ({ model: "m", customFunctions: true }),
  });
  const mod = await import("./hermes.js?cf=true");
  assert.equal(await mod.customFunctionsEnabled(), true);
});

test("customFunctionsEnabled: a truthy string does not open the gate", async () => {
  global.fetch = async () => ({
    ok: true,
    json: async () => ({ model: "m", customFunctions: "yes" }),
  });
  const mod = await import("./hermes.js?cf=string");
  assert.equal(await mod.customFunctionsEnabled(), false);
});
