/* global globalThis */
import test from "node:test";
import assert from "node:assert/strict";
import { createProposalMgr, hasChainedEdits, cellRefToPosition } from "./proposal-mgr.js";

function makeContext(state = {}) {
  const loadable = (props) => ({ ...props, load: () => {} });
  return {
    sync: async () => {},
    document: {
      getSelection: () =>
        loadable({
          tables: { items: state.selectionTables ?? state.tables ?? [] },
        }),
      getBookmarkRangeOrNullObject: () =>
        loadable({
          isNullObject: state.bookmarkText == null,
          text: state.bookmarkText ?? "",
          insertText: (next, mode) => {
            const inserted = { font: {}, insertBookmark: () => {} };
            if (state.onBookmarkInsert) state.onBookmarkInsert(next, mode, inserted);
            return inserted;
          },
        }),
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
    true
  );
});

test("hasChainedEdits ignores an edit with an empty find", () => {
  assert.equal(
    hasChainedEdits([
      { find: "", replace: "x" },
      { find: "unrelated", replace: "y" },
    ]),
    false
  );
});

test("hasChainedEdits detects a chain", () => {
  assert.equal(
    hasChainedEdits([
      { find: "foo", replace: "bar" },
      { find: "bar", replace: "baz" },
    ]),
    true
  );
});

test("hasChainedEdits passes independent edits", () => {
  assert.equal(
    hasChainedEdits([
      { find: "foo", replace: "bar" },
      { find: "baz", replace: "qux" },
    ]),
    false
  );
});

test("hasChainedEdits: empty replace never chains", () => {
  assert.equal(
    hasChainedEdits([
      { find: "foo", replace: "" },
      { find: "", replace: "bar" },
    ]),
    false
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
      })
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
      })
    );
  const mgr = createProposalMgr({
    target: { kind: "table", sig: SIG_2X2 },
    runWord,
  });
  await assert.rejects(
    () =>
      mgr.applyTable([{ cell: "A1", value: "x", old: "a" }], {
        markRed: false,
      }),
    /Không còn tìm thấy bảng/
  );
});

// body.tables does not descend into nested tables, so the selection is still
// needed as a fallback — dropping it would break Apply on a nested table.
test("applyTable falls back to the selection for a table the body does not list", async () => {
  const values = [
    ["a", "b"],
    ["c", "d"],
  ];
  const runWord = (fn) => fn(makeContext({ selectionTables: [makeTable(values)], bodyTables: [] }));
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
    /Không còn tìm thấy bảng/
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
        { markRed: false }
      ),
    /chồng chéo/
  );
});

test("applyFulldocEdits counts skipped edits with no match", async () => {
  const runWord = (fn) => fn(makeContext({ searchItems: [] }));
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  const result = await mgr.applyFulldocEdits([{ find: "missing", replace: "x" }], {
    markRed: false,
  });
  assert.equal(result.applied, 0);
  assert.equal(result.skipped, 1);
});

test("applyFulldocEdits applies a matching edit", async () => {
  const inserted = { font: {} };
  const searchItem = { insertText: () => inserted };
  const runWord = (fn) => fn(makeContext({ searchItems: [searchItem] }));
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  const result = await mgr.applyFulldocEdits([{ find: "old", replace: "new" }], { markRed: false });
  assert.equal(result.applied, 1);
  assert.equal(result.skipped, 0);
});

// One edit can rewrite many passages — Apply replaces EVERY match. Counting
// edits and calling the result "changes" told the user 1 when the document had
// three passages rewritten, which is exactly the number they would need to
// check before deciding whether to undo.
test("applyFulldocEdits counts passages rewritten, not edits attempted", async () => {
  const searchItems = [
    { insertText: () => ({ font: {} }) },
    { insertText: () => ({ font: {} }) },
    { insertText: () => ({ font: {} }) },
  ];
  const runWord = (fn) => fn(makeContext({ searchItems }));
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  const result = await mgr.applyFulldocEdits([{ find: "old", replace: "new" }], {
    markRed: false,
  });
  assert.equal(result.applied, 1, "one edit ran");
  assert.equal(result.replaced, 3, "three passages were rewritten");
});

test("applyFulldocEdits reports zero replacements when nothing matched", async () => {
  const runWord = (fn) => fn(makeContext({ searchItems: [] }));
  const mgr = createProposalMgr({ target: { kind: "fulldoc" }, runWord });
  const result = await mgr.applyFulldocEdits([{ find: "missing", replace: "x" }], {
    markRed: false,
  });
  assert.equal(result.replaced, 0);
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
  const result = await mgr.applyFulldocEdits([{ find: longFind, replace: "y" }], {
    markRed: false,
  });
  assert.equal(result.applied, 0);
  assert.equal(result.skipped, 1);
});

// ---- applyText (Pin/bookmark staleness) -----------------------------------
// BUG-05: this was 0% covered. Three branches decide where the replacement
// goes: the bookmark still holds the captured passage (write through it), the
// bookmark is stale but the passage is unique (search fallback), or the
// passage is too long to search AND the bookmark no longer matches (hard
// refuse). Plus the older-host path where bookmark APIs don't exist.

function withBookmarks() {
  globalThis.Office = { context: { requirements: { isSetSupported: () => true } } };
  return () => {
    delete globalThis.Office;
  };
}

test("applyText writes through the bookmark when it still holds the captured passage", async () => {
  const cleanup = withBookmarks();
  try {
    let insertArgs = null;
    let rePinned = false;
    const runWord = (fn) =>
      fn(
        makeContext({
          bookmarkText: "pinned passage",
          onBookmarkInsert: (next, mode, ins) => {
            insertArgs = [next, mode];
            ins.insertBookmark = () => {
              rePinned = true;
            };
          },
        })
      );
    const mgr = createProposalMgr({
      target: { kind: "text", text: "pinned passage" },
      runWord,
    });
    const result = await mgr.applyText("replacement", { markRed: false });
    assert.equal(result.applied, 1);
    assert.deepEqual(insertArgs, ["replacement", "Replace"]);
    assert.equal(rePinned, true, "the inserted text is re-pinned for a follow-up rewrite");
  } finally {
    cleanup();
  }
});

test("applyText falls back to a unique search when the bookmark no longer matches", async () => {
  const cleanup = withBookmarks();
  try {
    let searchHit = null;
    const runWord = (fn) =>
      fn(
        makeContext({
          bookmarkText: "some other passage",
          searchItems: [
            {
              insertText: (next, mode) => {
                searchHit = [next, mode];
                // canUseBookmark is true, so applyText re-pins the result —
                // the inserted range must carry insertBookmark.
                return { font: {}, insertBookmark: () => {} };
              },
            },
          ],
        })
      );
    const mgr = createProposalMgr({
      target: { kind: "text", text: "pinned passage" },
      runWord,
    });
    const result = await mgr.applyText("replacement", { markRed: false });
    assert.equal(result.applied, 1);
    assert.deepEqual(searchHit, ["replacement", "Replace"]);
  } finally {
    cleanup();
  }
});

test("applyText refuses when the captured text is too long and the bookmark no longer matches", async () => {
  const cleanup = withBookmarks();
  try {
    const longText = "x".repeat(300); // > MAX_SEARCH_LEN (255)
    const runWord = (fn) =>
      fn(
        makeContext({
          bookmarkText: "short stale text",
        })
      );
    const mgr = createProposalMgr({ target: { kind: "text", text: longText }, runWord });
    await assert.rejects(
      () => mgr.applyText("replacement", { markRed: false }),
      /không còn được ghim/
    );
  } finally {
    cleanup();
  }
});

test("applyText refuses when the fallback search is ambiguous", async () => {
  const cleanup = withBookmarks();
  try {
    const runWord = (fn) =>
      fn(
        makeContext({
          bookmarkText: "stale",
          searchItems: [],
        })
      );
    const mgr = createProposalMgr({
      target: { kind: "text", text: "pinned passage" },
      runWord,
    });
    await assert.rejects(
      () => mgr.applyText("replacement", { markRed: false }),
      /xác định duy nhất/
    );
  } finally {
    cleanup();
  }
});

test("applyText marks the replacement red through the bookmark when requested", async () => {
  const cleanup = withBookmarks();
  try {
    const inserted = { font: {} };
    const runWord = (fn) =>
      fn(
        makeContext({
          bookmarkText: "pinned passage",
          onBookmarkInsert: (_next, _mode, ins) => {
            ins.font = inserted.font;
          },
        })
      );
    const mgr = createProposalMgr({
      target: { kind: "text", text: "pinned passage" },
      runWord,
    });
    await mgr.applyText("replacement", { markRed: true });
    assert.equal(inserted.font.color, "#FF0000");
  } finally {
    cleanup();
  }
});

// Older Office hosts lack WordApi 1.4 bookmark APIs; applyText must skip
// straight to the search fallback rather than throw.
test("applyText falls back to search on hosts without bookmark APIs", async () => {
  // Office undefined → supportsBookmarks() returns false.
  delete globalThis.Office;
  let searchHit = null;
  const runWord = (fn) =>
    fn(
      makeContext({
        searchItems: [
          {
            insertText: (next, mode) => {
              searchHit = [next, mode];
              return { font: {}, insertBookmark: () => {} };
            },
          },
        ],
      })
    );
  const mgr = createProposalMgr({
    target: { kind: "text", text: "pinned passage" },
    runWord,
  });
  const result = await mgr.applyText("replacement", { markRed: true });
  assert.equal(result.applied, 1);
  assert.deepEqual(searchHit, ["replacement", "Replace"]);
});
