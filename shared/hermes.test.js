import test from "node:test";
import assert from "node:assert/strict";
import { askHermes } from "./hermes.js";

test("askHermes surfaces a clear error on a non-OK response, without retrying", async () => {
  let callCount = 0;
  global.fetch = async () => {
    callCount++;
    return { ok: false, status: 500, text: async () => "boom" };
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /Hermes 500/);
  assert.equal(callCount, 1, "an HTTP error response is a real answer and should not be retried");
});

test("askHermes retries once on timeout, then surfaces a clear bilingual error", async () => {
  let callCount = 0;
  global.fetch = async () => {
    callCount++;
    const err = new Error("The operation was aborted due to timeout");
    err.name = "TimeoutError"; // what AbortSignal.timeout() produces on abort
    throw err;
  };
  await assert.rejects(() => askHermes([{ role: "user", content: "hi" }]), /timeout|kịp thời/i);
  assert.equal(callCount, 2, "should retry exactly once on timeout");
});

test("askHermes still works where AbortSignal.timeout is unavailable", async () => {
  // Older Office WebViews lack AbortSignal.timeout; calling it there threw a
  // TypeError that the retry path misread as a network blip, so the user saw
  // a bogus "did not respond in time" instead of the real answer.
  const original = AbortSignal.timeout;
  delete AbortSignal.timeout;
  try {
    let seenSignal;
    global.fetch = async (_url, init) => {
      seenSignal = init.signal;
      return {
        ok: true,
        status: 200,
        json: async () => ({ choices: [{ message: { content: "ok" } }] }),
      };
    };
    assert.equal(await askHermes([{ role: "user", content: "hi" }]), "ok");
    assert.ok(seenSignal, "should still pass an abort signal via AbortController");
    assert.equal(seenSignal.aborted, false);
  } finally {
    AbortSignal.timeout = original;
  }
});

test("askHermes recovers from a single transient network failure", async () => {
  let callCount = 0;
  global.fetch = async () => {
    callCount++;
    if (callCount === 1) throw new TypeError("Failed to fetch");
    return {
      ok: true,
      status: 200,
      json: async () => ({ choices: [{ message: { content: "hello" } }] }),
    };
  };
  const reply = await askHermes([{ role: "user", content: "hi" }]);
  assert.equal(reply, "hello");
  assert.equal(callCount, 2);
});
