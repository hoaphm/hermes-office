import test from "node:test";
import assert from "node:assert/strict";
import {
  MAX_ACTIONS,
  assertReviewableActions,
  userErrorMessage,
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
