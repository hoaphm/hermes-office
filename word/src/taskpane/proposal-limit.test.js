import test from "node:test";
import assert from "node:assert/strict";
import {
  MAX_ACTIONS,
  assertReviewableActions,
  userErrorMessage,
  isProviderError,
  PROVIDER_UNAVAILABLE_MESSAGE,
} from "../../../shared/proposal-card.js";

// The Word reviewability gate. taskpane.js itself cannot be unit-tested in
// plain node (Office.onReady, document, Word.run), so the pure guard it calls
// in each of the table / text / fulldoc branches is the testable seam. The
// branch wiring is verified by inspection: each flow calls
// assertReviewableActions(parsedActions) BEFORE creating the Proposal or
// rendering a card, so an oversized payload throws before lastProposal is set
// or any table `old` value is mapped.

test("MAX_ACTIONS is the canonical export, equal to 100", () => {
  assert.equal(MAX_ACTIONS, 100);
});

test("assertReviewableActions accepts exactly the reviewable limit", () => {
  const actions = Array(100).fill({ type: "setCell", cell: "A1", new: "x" });
  assert.doesNotThrow(() => assertReviewableActions(actions));
  // Returns the array unchanged so callsites can chain.
  assert.equal(assertReviewableActions(actions), actions);
});

test("assertReviewableActions rejects one past the limit with the UI-safe message", () => {
  assert.throws(
    () => assertReviewableActions(Array(101).fill({})),
    /Đề xuất có quá nhiều thay đổi để xem xét an toàn/,
  );
});

test("assertReviewableActions rejects only on length, not content", () => {
  // 101 empty actions are still refused; a single action is always fine.
  assert.throws(() => assertReviewableActions(Array(101).fill({})), Error);
  assert.doesNotThrow(() => assertReviewableActions(Array(1).fill({})));
  assert.doesNotThrow(() => assertReviewableActions([]));
});

// ---- Error disclosure ------------------------------------------------------
// An oversized-Proposal rejection (and any other task-pane error) must reach
// the user as the safe message only — never the raw stack trace. The full
// error is logged via console.error by the caller.

test("userErrorMessage surfaces only the message, never the stack", () => {
  const err = new Error("Đề xuất có quá nhiều thay đổi để xem xét an toàn");
  const shown = userErrorMessage(err);
  assert.equal(shown, "Đề xuất có quá nhiều thay đổi để xem xét an toàn");
  assert.ok(!shown.includes(err.stack), "stack must not leak to the user");
  assert.ok(!shown.includes("\n"), "no multi-line disclosure");
});

test("userErrorMessage falls back to String(err) when there is no message", () => {
  assert.equal(userErrorMessage({}), "[object Object]");
  assert.equal(userErrorMessage("boom"), "boom");
  assert.equal(userErrorMessage(undefined), "undefined");
});

// ---- Provider / call-API disclosure policy --------------------------------
// Known call-pipeline stage failures must collapse to the fixed Vietnamese
// message (no upstream response text, stack, path, or line number), while
// local task-pane errors pass through verbatim.

test("isProviderError classifies call-pipeline stage tags", () => {
  assert.ok(isProviderError(new Error("[call-api:read-response] Provider 500: boom")));
  assert.ok(isProviderError(new Error("[call-api:fetch] TypeError: Failed to fetch")));
  assert.ok(isProviderError(new Error("[api-first] [call-api:read-response] Provider 503: x")));
  assert.ok(isProviderError(new Error("[api-retry] [call-api:fetch] boom")));
  // taskpane.js wraps askHermes failures in [stage:call-api].
  assert.ok(
    isProviderError(
      new Error("[stage:call-api] [api-first] [call-api:read-response] Provider 500: x"),
    ),
  );
  // config fetch is a network/HTTP stage of the call pipeline.
  assert.ok(
    isProviderError(
      new Error("[config-fetch:https://localhost:8643/config.json] network error"),
    ),
  );
  // Local task-pane errors are NOT provider failures.
  assert.ok(!isProviderError(new Error("Đề xuất có quá nhiều thay đổi để xem xét an toàn")));
  assert.ok(!isProviderError(new Error("Không còn tìm thấy bảng của đề xuất này.")));
  assert.ok(!isProviderError(new Error("[stage:read-doc] boom")));
});

test("provider failures show the fixed message and never the raw body", () => {
  const raw =
    "Provider 500: {\"error\":{\"message\":\"internal\",\"trace\":\"/app/src/line 42\"}}";
  const err = new Error(`[stage:call-api] [api-first] [call-api:read-response] ${raw}`);
  const shown = userErrorMessage(err);
  assert.equal(shown, PROVIDER_UNAVAILABLE_MESSAGE);
  assert.ok(
    !shown.includes(raw),
    "raw provider response body must never reach the user",
  );
  assert.ok(!shown.includes("trace"), "no provider diagnostic in the bubble");
  assert.ok(!shown.includes("line 42"), "no provider line number in the bubble");
  assert.ok(!shown.includes("\n"), "single-line, no stack");
});

test("config-fetch failures collapse to the fixed message", () => {
  const err = new Error("[config-fetch:https://localhost:8643/config.json] Failed to fetch");
  const shown = userErrorMessage(err);
  assert.equal(shown, PROVIDER_UNAVAILABLE_MESSAGE);
  assert.ok(!shown.includes("localhost"), "no local gateway URL in the bubble");
});

test("local task-pane errors are preserved verbatim", () => {
  assert.equal(
    userErrorMessage(new Error("Đề xuất có quá nhiều thay đổi để xem xét an toàn")),
    "Đề xuất có quá nhiều thay đổi để xem xét an toàn",
  );
  assert.equal(
    userErrorMessage(new Error("Không còn tìm thấy bảng của đề xuất này (bảng đã bị sửa hoặc xoá). Hãy chọn lại bảng rồi hỏi lại.")),
    "Không còn tìm thấy bảng của đề xuất này (bảng đã bị sửa hoặc xoá). Hãy chọn lại bảng rồi hỏi lại.",
  );
  assert.equal(
    userErrorMessage(new Error("Bảng đã thay đổi sau khi tạo đề xuất. Hãy hỏi lại Hermes.")),
    "Bảng đã thay đổi sau khi tạo đề xuất. Hãy hỏi lại Hermes.",
  );
  // Local document-read stage is not a provider failure.
  assert.equal(userErrorMessage(new Error("[stage:read-doc] boom")), "[stage:read-doc] boom");
});

test("the already-safe askHermes timeout message passes through unchanged", () => {
  const safe = "Provider không phản hồi kịp thời, vui lòng thử lại.";
  assert.equal(userErrorMessage(new Error(safe)), safe);
});
