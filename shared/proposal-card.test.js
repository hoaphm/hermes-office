import test from "node:test";
import assert from "node:assert/strict";
import {
  describeAction,
  previewGrid,
  writesBlind,
  partialApplyNotice,
} from "./proposal-card.js";

// The card IS the review boundary: Apply only means something if what it will
// write is on screen. These tests pin that contract for the actions whose
// payload does not fit in a one-line summary.

test("setCells: the card shows the values that will be written", () => {
  const d = describeAction({
    type: "setCells",
    range: "A1:C2",
    values: [["x", "y", "z"], ["1", "2", "3"]],
  });
  assert.ok(d.detail, "setCells must carry a payload preview");
  assert.ok(d.detail.includes("x") && d.detail.includes("3"));
});

test("setCells: a huge grid is bounded, and says how much was elided", () => {
  const values = Array.from({ length: 500 }, (_, r) =>
    Array.from({ length: 40 }, (_, c) => `r${r}c${c}`),
  );
  const detail = describeAction({ type: "setCells", range: "A1:AN500", values }).detail;
  assert.ok(detail.split("\n").length <= 10, "preview must stay short enough to read");
  assert.match(detail, /\+492 dòng nữa/);
  assert.match(detail, /\+32 cột/);
});

test("setCells: a malformed values payload does not throw", () => {
  assert.doesNotThrow(() => describeAction({ type: "setCells", range: "A1", values: "nope" }));
  assert.doesNotThrow(() => describeAction({ type: "setCells", range: "A1" }));
});

test("previewGrid: renders null and undefined cells as empty, not as text", () => {
  assert.equal(previewGrid([[null, undefined, "a"]]), "  │    │  a");
});

test("format: the card shows the fill colour it will paint", () => {
  const d = describeAction({ type: "format", range: "A1:Z999", fill: "#000000" });
  assert.match(d.summary, /#000000/);
});

// Apply replaces EVERY match, case-insensitively. The count belongs on the card,
// not in a toast fired once the document has already changed.
test("replace: the card states how many places will change", () => {
  assert.match(describeAction({ type: "replace", find: "a", replace: "b", matchCount: 7 }).summary, /TẤT CẢ 7/);
  assert.match(describeAction({ type: "replace", find: "a", replace: "b", matchCount: 1 }).summary, /1 chỗ/);
  assert.match(describeAction({ type: "replace", find: "a", replace: "b", matchCount: 0 }).summary, /không còn tìm thấy/);
});

// ---- Partial Apply: marking a Remainder ------------------------------------
// After a press has written part of a Proposal, the rest of it stands on a
// review the document has already moved past. Which rows that applies to is
// per-action information, so the card carries it per action — the same reason
// the replace count above is not summarised away. See ADR-0004.

test("writesBlind marks exactly the Actions with no old-value guard", () => {
  for (const type of ["setCells", "format", "createTable", "createChart"]) {
    assert.equal(writesBlind({ type }), true, `${type} writes blind`);
  }
  // setCell compares against a.old before writing, so the original review still
  // covers it even after the sheet has moved.
  assert.equal(writesBlind({ type: "setCell" }), false);
  assert.equal(writesBlind({ type: "newSheet" }), false);
  assert.equal(writesBlind(null), false);
  assert.equal(writesBlind(undefined), false);
});

test("partialApplyNotice counts the Actions the earlier review no longer covers", () => {
  const notice = partialApplyNotice([
    { type: "setCell", cell: "B2", old: "1", new: "2" },
    { type: "setCells", range: "A5:C9", values: [[]] },
    { type: "format", range: "C2:C50", bold: true },
  ]);
  assert.match(notice, /Sheet đã đổi/);
  assert.match(notice, /2 hành động/);
});

test("partialApplyNotice does not cry wolf when every Action still guards itself", () => {
  const notice = partialApplyNotice([
    { type: "setCell", cell: "B2", old: "1", new: "2" },
  ]);
  assert.match(notice, /Sheet đã đổi/);
  assert.match(notice, /vẫn kiểm tra giá trị hiện tại/);
  assert.doesNotMatch(notice, /ghi đè/);
});

test("replace: an uncounted edit claims nothing rather than claiming zero", () => {
  const d = describeAction({ type: "replace", find: "a", replace: "b", matchCount: null });
  assert.equal(d.summary, "Replace");
  assert.equal(describeAction({ type: "replace", find: "a", replace: "b" }).summary, "Replace");
});
