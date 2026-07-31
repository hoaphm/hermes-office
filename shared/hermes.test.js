import test from "node:test";
import assert from "node:assert/strict";
import { askHermes, callApi } from "./hermes.js";

function configFetch(response, seen) {
  return async (url, init) => {
    seen.push({ url, init });
    if (url === "/config.json") {
      return { ok: true, json: async () => ({ baseUrl: "https://api.example.test/v1", apiKey: "secret", model: "test-model" }) };
    }
    return response;
  };
}

test("askHermes calls configured OpenAI-compatible endpoint", async () => {
  const seen = [];
  global.fetch = configFetch({
    ok: true,
    status: 200,
    text: async () => JSON.stringify({ choices: [{ message: { content: "ok" } }] }),
  }, seen);
  assert.equal(await askHermes([{ role: "user", content: "hi" }]), "ok");
  const request = seen[1];
  assert.equal(request.url, "https://api.example.test/v1/chat/completions");
  assert.equal(request.init.headers.Authorization, "Bearer secret");
  assert.deepEqual(JSON.parse(request.init.body), {
    model: "test-model",
    messages: [{ role: "user", content: "hi" }],
  });
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
    if (url === "/config.json") return { ok: true, json: async () => ({ baseUrl: "https://api.example.test/v1", apiKey: "secret", model: "m" }) };
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
    if (url === "/config.json") return { ok: true, json: async () => ({ baseUrl: "https://api.example.test/v1", apiKey: "secret", model: "m" }) };
    seenInit = init;
    return { ok: true, text: async () => JSON.stringify({ choices: [{ message: { content: "ok" } }] }) };
  };
  assert.equal(await callApi([{ role: "user", content: "hi" }]), "ok");
  assert.equal(seenInit.signal, undefined, "fetch must receive no signal");
});

test("askHermes retries once on timeout", async () => {
  let callCount = 0;
  global.fetch = async (url) => {
    if (url === "/config.json") return { ok: true, json: async () => ({ baseUrl: "https://api.example.test/v1", apiKey: "secret", model: "m" }) };
    callCount++;
    const err = new Error("timeout");
    err.name = "TimeoutError";
    throw err;
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /kịp thời/);
  assert.equal(callCount, 2);
});
