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
      getSelection: () => loadable({ tables: { items: state.tables ?? [] } }),
      body: {
        load: () => {},
        tables: { items: state.tables ?? [] },
        search: () => loadable({ items: state.searchItems ?? [] }),
      },
    },
  };
}

// ---- pure helpers ----------------------------------------------------------

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
