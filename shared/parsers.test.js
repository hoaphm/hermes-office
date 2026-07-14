import test from "node:test";
import assert from "node:assert/strict";
import {
  columnIndexToLetters,
  columnLettersToIndex,
  parseEdits,
  parseTableChanges,
  signature,
  hash,
  resolveRange,
  chartType,
} from "./parsers.js";

test("columnIndexToLetters / columnLettersToIndex round-trip", () => {
  for (const letters of ["A", "Z", "AA", "AZ", "BA", "ZZ", "AAA"]) {
    const index = columnLettersToIndex(letters);
    assert.equal(columnIndexToLetters(index), letters, `round-trip failed for ${letters}`);
  }
});

test("columnLettersToIndex known base-26 values", () => {
  assert.equal(columnLettersToIndex("A"), 0);
  assert.equal(columnLettersToIndex("Z"), 25);
  assert.equal(columnLettersToIndex("AA"), 26);
  assert.equal(columnLettersToIndex("AZ"), 51);
  assert.equal(columnLettersToIndex("BA"), 52);
  assert.equal(columnLettersToIndex("ZZ"), 701);
  assert.equal(columnLettersToIndex("AAA"), 702);
});

test("parseEdits: fenced ```json block", () => {
  const reply = 'Fixed a typo.\n```json\n{"edits":[{"find":"teh","replace":"the"}]}\n```';
  assert.deepEqual(parseEdits(reply), [{ find: "teh", replace: "the" }]);
});

test("parseEdits: bare inline object embedded in prose", () => {
  const reply = 'Here are the fixes: {"edits":[{"find":"foo","replace":"bar"}]} — done.';
  assert.deepEqual(parseEdits(reply), [{ find: "foo", replace: "bar" }]);
});

test("parseEdits: missing replace defaults to empty string", () => {
  const reply = '```json\n{"edits":[{"find":"x"}]}\n```';
  assert.deepEqual(parseEdits(reply), [{ find: "x", replace: "" }]);
});

test("parseEdits: no edits present returns []", () => {
  assert.deepEqual(parseEdits("Just a plain answer, no changes needed."), []);
});

test("parseTableChanges: fenced json block", () => {
  const reply = 'Updated.\n```json\n{"cells":[{"cell":"A1","value":"x"}]}\n```';
  assert.deepEqual(parseTableChanges(reply), [{ cell: "A1", value: "x" }]);
});

test("parseTableChanges: bare object with no fence", () => {
  const reply = 'ok {"cells":[{"cell":"B2","value":"y"}]}';
  assert.deepEqual(parseTableChanges(reply), [{ cell: "B2", value: "y" }]);
});

test("parseTableChanges: no cells present returns []", () => {
  assert.deepEqual(parseTableChanges("nothing to change here"), []);
});

test("signature: stable for an identical snapshot", () => {
  const snap = { name: "Sheet1", address: "A1:B2", values: [["a", "b"], ["c", "d"]], selection: null };
  assert.equal(signature(snap), signature(snap));
});

test("signature: differs when the selection changes", () => {
  const base = { name: "Sheet1", address: "A1:B2", values: [["a", "b"], ["c", "d"]] };
  const sig1 = signature({ ...base, selection: { address: "A1", values: [["a"]] } });
  const sig2 = signature({ ...base, selection: { address: "B2", values: [["d"]] } });
  assert.notEqual(sig1, sig2);
});

test("resolveRange: unqualified address uses the fallback sheet", () => {
  const calls = [];
  const fallbackSheet = { getRange: (addr) => { calls.push(["fallback.getRange", addr]); return { addr }; } };
  const wb = { worksheets: { getItem: () => { throw new Error("should not be called"); } } };
  const r = resolveRange(wb, fallbackSheet, "A1");
  assert.deepEqual(calls, [["fallback.getRange", "A1"]]);
  assert.deepEqual(r, { addr: "A1" });
});

test("resolveRange: sheet-qualified address resolves via workbook.worksheets.getItem", () => {
  const calls = [];
  const fallbackSheet = { getRange: () => { throw new Error("should not be called"); } };
  const wb = {
    worksheets: {
      getItem: (name) => {
        calls.push(["getItem", name]);
        return { getRange: (addr) => { calls.push(["sheet.getRange", addr]); return { name, addr }; } };
      },
    },
  };
  const r = resolveRange(wb, fallbackSheet, "Sheet!A1");
  assert.deepEqual(calls, [["getItem", "Sheet"], ["sheet.getRange", "A1"]]);
  assert.deepEqual(r, { name: "Sheet", addr: "A1" });
});

test("chartType: maps known aliases and defaults to ColumnClustered", () => {
  assert.equal(chartType("Pie"), "Pie");
  assert.equal(chartType("bar"), "BarClustered");
  assert.equal(chartType("unknown-type"), "ColumnClustered");
  assert.equal(chartType(undefined), "ColumnClustered");
});

test("hash: deterministic and sensitive to input", () => {
  assert.equal(hash("abc"), hash("abc"));
  assert.notEqual(hash("abc"), hash("abd"));
});
