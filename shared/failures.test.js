import test from "node:test";
import assert from "node:assert/strict";
import {
  FAILURES,
  classifyFailure,
  userErrorMessage,
  isProviderError,
  hopReport,
} from "./failures.js";

// One test per Disclosed Failure, plus the two invariants ADR-0005 rests on:
// nothing from the Provider is ever disclosed, and a local task-pane error is
// not a Disclosed Failure at all.
//
// `hermes` is the fact bag shared/hermes.js records at the throw site; the
// helper mirrors what the real pipeline attaches.
const pipelineError = (hermes, message = "[call-api:read-response] Provider …") =>
  Object.assign(new Error(message), { hermes });

test("GATEWAY_DOWN: nothing answered on the Local Hop", () => {
  // Gateway not listening: the config fetch never got a status at all.
  assert.equal(
    classifyFailure(pipelineError({ stage: "config-fetch", status: undefined })),
    FAILURES.GATEWAY_DOWN,
  );
  // The chat POST itself failing at the network layer is the same thing.
  assert.equal(
    classifyFailure(pipelineError({ stage: "fetch", name: "TypeError" })),
    FAILURES.GATEWAY_DOWN,
  );
});

test("NOT_CONFIGURED: the Gateway answered but has no usable config", () => {
  assert.equal(
    classifyFailure(pipelineError({ stage: "config-model" })),
    FAILURES.NOT_CONFIGURED,
  );
  // config.json served as a 404 is the same problem with the same fix.
  assert.equal(
    classifyFailure(pipelineError({ stage: "config-fetch", status: 404 })),
    FAILURES.NOT_CONFIGURED,
  );
});

// The failure this whole design exists for.
test("UPSTREAM_UNREACHABLE: only when the Gateway marked the response as its own", () => {
  const marked = pipelineError({ stage: "read-response", status: 502, gateway: true });
  assert.equal(classifyFailure(marked), FAILURES.UPSTREAM_UNREACHABLE);
  assert.equal(classifyFailure(marked).hop, "upstream");
});

test("UPSTREAM_ERROR: the same status unmarked came from the Provider itself", () => {
  const fromProvider = pipelineError({
    stage: "read-response",
    status: 502,
    gateway: false,
  });
  assert.equal(classifyFailure(fromProvider), FAILURES.UPSTREAM_ERROR);
  assert.notEqual(
    FAILURES.UPSTREAM_ERROR.message,
    FAILURES.UPSTREAM_UNREACHABLE.message,
    "the two 502s must not be disclosed as the same thing",
  );
});

test("KEY_REJECTED: 401 and 403", () => {
  for (const status of [401, 403]) {
    assert.equal(
      classifyFailure(pipelineError({ stage: "read-response", status })),
      FAILURES.KEY_REJECTED,
    );
  }
});

test("MODEL_REJECTED: 400 and 404 are a config mistake, not a network one", () => {
  for (const status of [400, 404]) {
    assert.equal(
      classifyFailure(pipelineError({ stage: "read-response", status })),
      FAILURES.MODEL_REJECTED,
    );
  }
  assert.match(FAILURES.MODEL_REJECTED.message, /model/);
});

test("OVERLOADED: 429", () => {
  assert.equal(
    classifyFailure(pipelineError({ stage: "read-response", status: 429 })),
    FAILURES.OVERLOADED,
  );
});

test("UPSTREAM_SLOW: a timeout, whichever stage it was raised in", () => {
  assert.equal(
    classifyFailure(pipelineError({ stage: "fetch", name: "TimeoutError" })),
    FAILURES.UPSTREAM_SLOW,
  );
  assert.equal(
    classifyFailure(pipelineError({ stage: "fetch", name: "AbortError" })),
    FAILURES.UPSTREAM_SLOW,
  );
});

test("BAD_PAYLOAD: a 200 whose body is not a completion", () => {
  assert.equal(
    classifyFailure(
      pipelineError({ stage: "read-response", status: 200, badPayload: true }),
    ),
    FAILURES.BAD_PAYLOAD,
  );
});

test("UNKNOWN: a pipeline failure that matches nothing, named as such", () => {
  assert.equal(
    classifyFailure(pipelineError({ stage: "read-response", status: 302 })),
    FAILURES.UNKNOWN,
  );
  // An error that lost its fact bag crossing a boundary still classifies as a
  // pipeline failure on its tag alone, rather than leaking its message.
  assert.equal(
    classifyFailure(new Error("[api-retry] [call-api:fetch] Provider 500: boom")),
    FAILURES.UNKNOWN,
  );
});

test("every state names its hop and a distinct message", () => {
  const messages = new Set();
  for (const [key, failure] of Object.entries(FAILURES)) {
    assert.equal(failure.code, key, "code must match its key");
    assert.ok(["local", "upstream"].includes(failure.hop), `${key} names a hop`);
    assert.ok(!messages.has(failure.message), `${key} has its own message`);
    messages.add(failure.message);
  }
});

// ---- ADR-0005 invariants ---------------------------------------------------

test("no Disclosed Failure can carry Provider text, a stack, a path or a URL", () => {
  const raw =
    'Provider 500: {"error":{"message":"internal","trace":"/app/src/router.js line 42"}}';
  const err = pipelineError(
    { stage: "read-response", status: 500 },
    `[stage:call-api] [api-first] [call-api:read-response] ${raw}`,
  );
  const shown = userErrorMessage(err);
  assert.equal(shown, FAILURES.UPSTREAM_ERROR.message);
  assert.ok(!shown.includes("trace"), "no Provider diagnostic");
  assert.ok(!shown.includes("line 42"), "no Provider line number");
  assert.ok(!shown.includes("router.js"), "no Provider path");
  assert.ok(!shown.includes("\n"), "single line, no stack");

  // The resolved gateway URL that the config-fetch tag embeds must not escape.
  const configErr = pipelineError(
    { stage: "config-fetch" },
    "[config-fetch:https://localhost:8643/config.json] Failed to fetch",
  );
  assert.ok(!userErrorMessage(configErr).includes("localhost"));
});

test("the message set as a whole is free of interpolated text", () => {
  for (const failure of Object.values(FAILURES)) {
    assert.ok(!/https?:\/\//.test(failure.message), `${failure.code}: no URL`);
    assert.ok(!failure.message.includes("\n"), `${failure.code}: single line`);
    // CONTEXT.md lists "Hermes server" among the names to avoid for the
    // Provider; the message this set replaced said "dịch vụ Hermes", which
    // blurred the two hops into one unnameable thing.
    assert.ok(!/dịch vụ Hermes/i.test(failure.message), `${failure.code}: glossary`);
  }
});

test("local task-pane errors are not Disclosed Failures and pass through verbatim", () => {
  const local = "Đề xuất có quá nhiều thay đổi để xem xét an toàn";
  assert.equal(classifyFailure(new Error(local)), null);
  assert.equal(isProviderError(new Error(local)), false);
  assert.equal(userErrorMessage(new Error(local)), local);
  assert.equal(userErrorMessage(new Error("[stage:read-doc] boom")), "[stage:read-doc] boom");
});

// ---- Hop report ------------------------------------------------------------

test("hopReport names each hop and reuses the Disclosed Failure wording", () => {
  assert.deepEqual(
    hopReport({ local: { ok: true }, upstream: { ok: true } }),
    ["Local Hop: OK", "Upstream Hop: OK"],
  );

  const broken = hopReport({
    local: { ok: true },
    upstream: {
      ok: false,
      error: pipelineError({ stage: "read-response", status: 502, gateway: true }),
    },
  });
  assert.equal(broken[0], "Local Hop: OK");
  assert.ok(broken[1].startsWith("Upstream Hop: "));
  assert.ok(broken[1].includes(FAILURES.UPSTREAM_UNREACHABLE.message));
});

test("hopReport does not pretend to know about the Upstream Hop when the local one is down", () => {
  const lines = hopReport({
    local: { ok: false, error: pipelineError({ stage: "config-fetch" }) },
    upstream: { ok: false, error: null, skipped: true },
  });
  assert.ok(lines[0].includes(FAILURES.GATEWAY_DOWN.message));
  assert.match(lines[1], /chưa kiểm tra/);
  // A skipped probe must never read as a verdict on the Provider.
  assert.equal(lines[1].includes(FAILURES.UPSTREAM_UNREACHABLE.message), false);
});

test("userErrorMessage falls back to String(err) when there is no message", () => {
  assert.equal(userErrorMessage({}), "[object Object]");
  assert.equal(userErrorMessage("boom"), "boom");
  assert.equal(userErrorMessage(undefined), "undefined");
});
