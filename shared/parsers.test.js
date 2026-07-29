import test from "node:test";
import assert from "node:assert/strict";
import {
  columnIndexToLetters,
  columnLettersToIndex,
  parseEdits,
  parseTableChanges,
  extractJsonObject,
  signature,
  hash,
  resolveRange,
  chartType,
  parseEditsReply,
  partitionEdits,
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

test("parseTableChanges: a non-array `cells` is rejected, not returned", () => {
  // Used to `return obj.cells || []`, handing a bare object to callers that
  // immediately did .length / .map on it.
  assert.deepEqual(parseTableChanges('```json\n{"cells":{"cell":"A1"}}\n```'), []);
});

test("parseTableChanges: drops malformed entries", () => {
  const reply = '```json\n{"cells":[{"cell":"A1","value":"x"},{"value":"no cell"},null]}\n```';
  assert.deepEqual(parseTableChanges(reply), [{ cell: "A1", value: "x" }]);
});

test("parseEdits: unfenced object containing a nested array still parses", () => {
  // The old lazy /\[[\s\S]*?\]/ regex stopped at the first inner `]`, yielding
  // truncated JSON that failed to parse and silently dropped every edit.
  const reply = 'Fixes: {"edits":[{"find":"a","replace":"b","spans":[1,2]}]} done.';
  assert.deepEqual(parseEdits(reply), [{ find: "a", replace: "b" }]);
});

test("extractJsonObject: picks the object holding the key, ignoring earlier braces", () => {
  const text = 'noise {"other":1} then {"actions":[{"type":"setCell"}]} tail';
  assert.equal(extractJsonObject(text, "actions"), '{"actions":[{"type":"setCell"}]}');
});

test("extractJsonObject: braces inside string values do not break balancing", () => {
  const text = '{"actions":[{"type":"setCell","new":"} not a brace {"}]}';
  assert.equal(extractJsonObject(text, "actions"), text);
  assert.equal(JSON.parse(extractJsonObject(text, "actions")).actions.length, 1);
});

test("extractJsonObject: returns null when the key is absent", () => {
  assert.equal(extractJsonObject('{"cells":[]}', "actions"), null);
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

// ---------------------------------------------------------------------------
// parseEditsReply — one pass over a full-document reply, giving the task pane
// everything it needs: the edits, whether the model actually complied with the
// JSON contract, and the prose to show in the chat bubble.
//
// Word used to print the raw reply into the chat, so a compliant model that
// answered `{"edits":[]}` showed the user literal JSON instead of "no errors".

test("parseEditsReply: empty edits is structured compliance, not prose", () => {
  const r = parseEditsReply('{"edits":[]}');
  assert.equal(r.structured, true);
  assert.deepEqual(r.edits, []);
  assert.equal(r.prose, "");
});

test("parseEditsReply: edits are returned and stripped from the prose", () => {
  const r = parseEditsReply('Đây là các lỗi:\n{"edits":[{"find":"hoc","replace":"học"}]}');
  assert.equal(r.structured, true);
  assert.deepEqual(r.edits, [{ find: "hoc", replace: "học" }]);
  assert.equal(r.prose, "Đây là các lỗi:");
});

test("parseEditsReply: a fenced block is stripped from the prose too", () => {
  const r = parseEditsReply('Xong.\n```json\n{"edits":[{"find":"a","replace":"b"}]}\n```');
  assert.equal(r.structured, true);
  assert.equal(r.edits.length, 1);
  assert.equal(r.prose, "Xong.");
});

test("parseEditsReply: a prose-only reply is NOT structured", () => {
  const r = parseEditsReply("Tài liệu của bạn có vài chỗ nên viết lại cho gọn hơn.");
  assert.equal(r.structured, false);
  assert.deepEqual(r.edits, []);
  assert.equal(r.prose, "Tài liệu của bạn có vài chỗ nên viết lại cho gọn hơn.");
});

test("parseEdits: identical find/replace pairs are deduplicated", () => {
  // Observed from a real model: the same word listed once per occurrence.
  // Each edit replaces ALL matches, so the duplicate found nothing and was
  // reported as "1 skipped" on a run where nothing actually went wrong.
  const edits = parseEdits(
    '{"edits":[{"find":"hoc","replace":"học"},{"find":"dậy","replace":"dạy"},{"find":"hoc","replace":"học"}]}'
  );
  assert.deepEqual(edits, [
    { find: "hoc", replace: "học" },
    { find: "dậy", replace: "dạy" },
  ]);
});

test("parseEdits: same find with a DIFFERENT replace is kept (not a duplicate)", () => {
  const edits = parseEdits('{"edits":[{"find":"a","replace":"b"},{"find":"a","replace":"c"}]}');
  assert.equal(edits.length, 2);
});

// ---------------------------------------------------------------------------
// partitionEdits — decide which find strings are safe to hand to
// Word's body.search().
//
// search() has no matchWholeWord here, so it matches SUBSTRINGS and Apply
// replaces every hit. A 2-character find like "va" (-> "và") therefore also
// rewrites the inside of "van", "vay", "vang", "vai" — all real Vietnamese
// syllables. Short finds are refused rather than applied blindly.
//
// The floor is 3, not higher: "hoc" -> "học" is a 3-character correction and
// one of the most common Vietnamese typos there is, so raising the bar would
// throw away more than it protects.

test("partitionEdits: keeps finds at or above the minimum length", () => {
  const { applicable, tooShort } = partitionEdits([
    { find: "hoc", replace: "học" },
    { find: "choi", replace: "chơi" },
  ]);
  assert.equal(applicable.length, 2);
  assert.equal(tooShort.length, 0);
});

test("partitionEdits: refuses 1- and 2-character finds", () => {
  const { applicable, tooShort } = partitionEdits([
    { find: "va", replace: "và" },
    { find: "o", replace: "ở" },
    { find: "hoc", replace: "học" },
  ]);
  assert.deepEqual(
    applicable.map((e) => e.find),
    ["hoc"]
  );
  assert.deepEqual(
    tooShort.map((e) => e.find),
    ["va", "o"]
  );
});

test("partitionEdits: still refuses finds over the search-length ceiling", () => {
  const { applicable, tooLong } = partitionEdits([
    { find: "x".repeat(256), replace: "y" },
    { find: "hoc", replace: "học" },
  ]);
  assert.equal(applicable.length, 1);
  assert.equal(tooLong.length, 1);
});

test("partitionEdits: length is counted in characters, not UTF-16 units", () => {
  // "đã" is 2 characters; a naive .length on some normalisations would
  // over-count and wrongly let it through.
  const { tooShort } = partitionEdits([{ find: "đã", replace: "đá" }]);
  assert.equal(tooShort.length, 1);
});

test("partitionEdits: leading/trailing spaces do not buy length", () => {
  // " va " is 4 characters but only 2 of signal; padding must not be a way
  // to sneak a dangerous find past the floor.
  const { tooShort } = partitionEdits([{ find: " va ", replace: " và " }]);
  assert.equal(tooShort.length, 1);
});

test("partitionEdits: empty input yields empty buckets", () => {
  const r = partitionEdits([]);
  assert.deepEqual(r.applicable, []);
  assert.deepEqual(r.tooShort, []);
  assert.deepEqual(r.tooLong, []);
});

test("parseEdits: dedupe key cannot collide across the find/replace boundary", () => {
  // With a space separator both of these key to "a b c", so the second edit
  // would be dropped as a false duplicate. They are different instructions.
  const edits = parseEdits(
    '{"edits":[{"find":"a b","replace":"c"},{"find":"a","replace":"b c"}]}'
  );
  assert.equal(edits.length, 2);
});
