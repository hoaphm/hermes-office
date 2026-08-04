import test from "node:test";
import assert from "node:assert/strict";
import {
  createProposalMgr,
  hasChainedEdits,
  cellRefToPosition,
} from "./proposal-mgr.js";

function makeContext(state = {}) {
  const loadable = (props) => ({ ...props, load: () => {} });
  return {
    sync: async () => {},
    document: {
      getSelection: () =>
        loadable({ tables: { items: state.selectionTables ?? state.tables ?? [] } }),
      body: {
        load: () => {},
        tables: { items: state.bodyTables ?? state.tables ?? [] },
        search: () => loadable({ items: state.searchItems ?? [] }),
      },
    },
  };
}

// A Word table proxy: enough of the surface findProposalTable/applyTable touch.
// Each call builds a FRESH object, because that is what Office.js does — the
// same physical table reached by two navigation paths is two client objects.
function makeTable(values) {
  return {
    rowCount: values.length,
    columnCount: values[0].length,
    load: () => {},
    getCell: (r, c) => ({
      body: {
        getRange: () => ({
          load: () => {},
          text: values[r][c],
          insertText(next) {
            values[r][c] = next;
            return { font: {} };
          },
        }),
      },
    }),
  };
}

// ---- pure helpers ----------------------------------------------------------

// applyFulldocEdits searches with { matchCase: false }, so the guard has to
// fold case too. It did not: "cat"→"Dog" then "dog"→"wolf" passed the check,
// and the second search matched the "Dog" the first edit had just written.
test("hasChainedEdits detects a chain that only matches case-insensitively", () => {
  assert.equal(
    hasChainedEdits([
      { find: "cat", replace: "Dog" },
      { find: "dog", replace: "wolf" },
    ]),
    true,
  );
});

test("hasChainedEdits ignores an edit with an empty find", () => {
  assert.equal(
    hasChainedEdits([
      { find: "", replace: "x" },
      { find: "unrelated", replace: "y" },
    ]),
    false,
  );
});

test("hasChainedEdits detects a chain", () => {
  assert.equal(
    hasChainedEdits([
      { find: "foo", replace: "bar" },
      { find: "bar", replace: "baz" },
    ]),
    true,
  );
});

test("hasChainedEdits passes independent edits", () => {
  assert.equal(
    hasChainedEdits([
      { find: "foo", replace: "bar" },
      { find: "baz", replace: "qux" },
    ]),
    false,
  );
});

test("hasChainedEdits: empty replace never chains", () => {
  assert.equal(
    hasChainedEdits([
      { find: "foo", replace: "" },
      { find: "", replace: "bar" },
    ]),
    false,
  );
});

test("cellRefToPosition converts A1 notation", () => {
  assert.deepEqual(cellRefToPosition("A1", 5, 5), { row: 0, col: 0 });
  assert.deepEqual(cellRefToPosition("B3", 5, 5), { row: 2, col: 1 });
  assert.deepEqual(cellRefToPosition("Z1", 5, 26), { row: 0, col: 25 });
  assert.deepEqual(cellRefToPosition("AA1", 5, 27), { row: 0, col: 26 });
});

test("cellRefToPosition rejects out-of-bounds refs", () => {
  assert.equal(cellRefToPosition("A10", 5, 5), null); // row out of bounds
  assert.equal(cellRefToPosition("F1", 5, 5), null); // col out of bounds
  assert.equal(cellRefToPosition("bad", 5, 5), null); // not A1 notation
  assert.equal(cellRefToPosition("", 5, 5), null);
});

// ---- applyTable (bug #4: wrong-table guard) --------------------------------

const SIG_2X2 = '2x2:[["a","b"],["c","d"]]';

// The bug this pins: findProposalTable used to merge the selection's tables with
// the body's and dedupe the merged list with `new Set(...)` — object identity.
// Office.js materialises a separate client object per navigation path, so the
// ONE table the user selected arrived as two distinct objects, both matched the
// signature, and "two matches" is the ambiguity case that refuses. Applying to a
// selected table could therefore never succeed.
test("applyTable finds the table when selection and body report it separately", async () => {
  const values = [
    ["a", "b"],
    ["c", "d"],
  ];
  const runWord = (fn) =>
    fn(
      makeContext({
        selectionTables: [makeTable(values)],
        bodyTables: [makeTable(values)],
      }),
    );
  const mgr = createProposalMgr({
    target: { kind: "table", sig: SIG_2X2 },
    runWord,
  });
  const result = await mgr.applyTable([{ cell: "A1", value: "x", old: "a" }], {
    markRed: false,
  });
  assert.equal(result.applied, 1);
});

// The ambiguity guard still has to work: two genuinely different tables holding
// identical content cannot be told apart, so writing into either is a guess.
test("applyTable still refuses when two distinct tables share the content", async () => {
  const runWord = (fn) =>
    fn(
      makeContext({
        selectionTables: [],
        bodyTables: [
          makeTable([
            ["a", "b"],
            ["c", "d"],
          ]),
          makeTable([
            ["a", "b"],
            ["c", "d"],
          ]),
        ],
      }),
    );
  const mgr = createProposalMgr({
    target: { kind: "table", sig: SIG_2X2 },
    runWord,
  });
  await assert.rejects(
    () => mgr.applyTable([{ cell: "A1", value: "x", old: "a" }], { markRed: false }),
    /Không còn tìm thấy bảng/,
  );
});

// body.tables does not descend into nested tables, so the selection is still
// needed as a fallback — dropping it would break Apply on a nested table.
test("applyTable falls back to the selection for a table the body does not list", async () => {
  const values = [
    ["a", "b"],
    ["c", "d"],
  ];
  const runWord = (fn) =>
    fn(makeContext({ selectionTables: [makeTable(values)], bodyTables: [] }));
  const mgr = createProposalMgr({
    target: { kind: "table", sig: SIG_2X2 },
    runWord,
  });
  const result = await mgr.applyTable([{ cell: "B2", value: "z", old: "d" }], {
    markRed: false,
  });
  assert.equal(result.applied, 1);
});

test("applyTable rejects when no table matches the target sig", async () => {
  const runWord = (fn) => fn(makeContext({ tables: [] }));
  const mgr = createProposalMgr({
    target: { kind: "table", sig: '2x2:[["a","b"],["c","d"]]', table: {} },
    runWord,
  });
  await assert.rejects(
    () =>
      mgr.applyTable([{ cell: "A1", value: "x", old: "a" }], {
        markRed: false,
      }),
    /Không còn tìm thấy bảng/,
  );
});

// ---- applyFulldocEdits -----------------------------------------------------

test("applyFulldocEdits rejects chained edits", async () => {
  const runWord = (fn) => fn(makeContext());
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  await assert.rejects(
    () =>
      mgr.applyFulldocEdits(
        [
          { find: "foo", replace: "bar" },
          { find: "bar", replace: "baz" },
        ],
        { markRed: false },
      ),
    /chồng chéo/,
  );
});

test("applyFulldocEdits counts skipped edits with no match", async () => {
  const runWord = (fn) => fn(makeContext({ searchItems: [] }));
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  const result = await mgr.applyFulldocEdits(
    [{ find: "missing", replace: "x" }],
    { markRed: false },
  );
  assert.equal(result.applied, 0);
  assert.equal(result.skipped, 1);
});

test("applyFulldocEdits applies a matching edit", async () => {
  const inserted = { font: {} };
  const searchItem = { insertText: () => inserted };
  const runWord = (fn) => fn(makeContext({ searchItems: [searchItem] }));
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  const result = await mgr.applyFulldocEdits(
    [{ find: "old", replace: "new" }],
    { markRed: false },
  );
  assert.equal(result.applied, 1);
  assert.equal(result.skipped, 0);
});

test("applyFulldocEdits marks red when requested", async () => {
  const inserted = { font: {} };
  const searchItem = { insertText: () => inserted };
  const runWord = (fn) => fn(makeContext({ searchItems: [searchItem] }));
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  await mgr.applyFulldocEdits([{ find: "old", replace: "new" }], {
    markRed: true,
  });
  assert.equal(inserted.font.color, "#FF0000");
});

test("applyFulldocEdits skips edits longer than MAX_SEARCH_LEN", async () => {
  const runWord = (fn) => fn(makeContext());
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  const longFind = "x".repeat(300);
  const result = await mgr.applyFulldocEdits(
    [{ find: longFind, replace: "y" }],
    { markRed: false },
  );
  assert.equal(result.applied, 0);
  assert.equal(result.skipped, 1);
});
