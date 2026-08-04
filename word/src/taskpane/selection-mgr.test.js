import test from "node:test";
import assert from "node:assert/strict";
import { createSelectionMgr, tableSignature } from "./selection-mgr.js";

// Minimal fake Office.js context. Properties are pre-set (no real load/sync
// round-trip); load() is a no-op and sync() resolves immediately.
function makeContext(state) {
  const loadable = (props) => ({ ...props, load: () => {} });
  return {
    sync: async () => {},
    document: {
      getSelection: () =>
        loadable({
          text: state.selectionText ?? "",
          tables: { items: state.tables ?? [] },
          parentTable: loadable({
            isNullObject: (state.tables ?? []).length === 0,
          }),
          insertBookmark: () => {},
        }),
      getBookmarkRangeOrNullObject: () =>
        state.bookmarkText != null
          ? loadable({ isNullObject: false, text: state.bookmarkText })
          : loadable({ isNullObject: true, text: "" }),
      deleteBookmark: () => {},
      body: {
        getRange: () => loadable({ text: state.bodyText ?? "" }),
        search: () => loadable({ items: [] }),
      },
      onSelectionChanged: { add: () => {} },
    },
  };
}

// ---- ActivityMachine gate (bugs #2 and #3) ---------------------------------

test("gate: onSelectionChanged is a no-op while reading", async () => {
  let runWordCalled = false;
  const runWord = (fn) => {
    runWordCalled = true;
    return fn(makeContext({ selectionText: "hello" }));
  };
  const sel = createSelectionMgr({ runWord });

  const handle = sel.beginAsk(); // activity → "reading"
  await sel.onSelectionChanged(); // must be ignored
  handle.end();

  assert.equal(
    runWordCalled,
    false,
    "runWord must not be called while reading",
  );
  assert.equal(sel.getPinnedText(), "", "pin must not change while reading");
});

test("gate: onSelectionChanged is a no-op while applying", async () => {
  let runWordCalled = false;
  const runWord = (fn) => {
    runWordCalled = true;
    return fn(makeContext({ selectionText: "hello" }));
  };
  const sel = createSelectionMgr({ runWord });

  const handle = sel.beginApply(); // activity → "applying"
  await sel.onSelectionChanged();
  handle.end();

  assert.equal(runWordCalled, false);
  assert.equal(sel.getPinnedText(), "");
});

test("gate: focus-loss (empty selection) does not clear pin while reading", async () => {
  const state = { selectionText: "pinned passage", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  await sel.onSelectionChanged(); // pin "pinned passage"
  assert.equal(sel.getPinnedText(), "pinned passage");

  // Focus leaves the document → empty selectionChanged fires.
  state.selectionText = "";
  const handle = sel.beginAsk(); // activity → "reading"
  await sel.onSelectionChanged(); // must be ignored (bug #3)
  handle.end();

  assert.equal(
    sel.getPinnedText(),
    "pinned passage",
    "pin must survive focus-loss",
  );
});

// ---- serialising the handler ----------------------------------------------

// Two selectionChanged events arriving close together each ran their own
// read-then-pin sequence concurrently. Their Word.run batches interleave, so
// the OLDER selection's insertBookmark could land last and win — the stale-pin
// bug the bookmark was introduced to prevent, reached by a different route.
// getSelectionData() also awaited only the newest pin, so an Ask could start
// while an older one was still writing.
test("pin: overlapping selection changes never run two batches at once", async () => {
  const state = { selectionText: "A", bookmarkText: null };
  let inFlight = 0;
  let maxInFlight = 0;
  const runWord = async (fn) => {
    inFlight += 1;
    maxInFlight = Math.max(maxInFlight, inFlight);
    try {
      return await fn(makeContext(state));
    } finally {
      inFlight -= 1;
    }
  };
  const sel = createSelectionMgr({ runWord });

  // Fire both without awaiting the first, as Word does on a fast drag.
  const first = sel.onSelectionChanged();
  const second = sel.onSelectionChanged();
  await Promise.all([first, second]);

  assert.equal(
    maxInFlight,
    1,
    "a second selectionChanged must queue behind the first, not race it",
  );
});

// ---- the Pin outlives focus while a Proposal needs it ----------------------

// The ActivityMachine gate only holds while an Ask or Apply is running. Once
// the Ask finishes the pane goes idle with a Proposal on screen, and the next
// empty selectionChanged — which is what firing focus into the taskpane looks
// like — cleared the very Pin that Proposal is anchored to. Apply then fell
// back to a body search and refused outright if the passage was not unique.
test("pin: a locked Pin survives an empty selection", async () => {
  const state = { selectionText: "pinned passage", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  await sel.onSelectionChanged();
  assert.equal(sel.getPinnedText(), "pinned passage");

  sel.lockPin(); // a text Proposal is now on screen and depends on this Pin
  state.selectionText = "";
  await sel.onSelectionChanged();
  assert.equal(
    sel.getPinnedText(),
    "pinned passage",
    "the Proposal's Pin must survive focus leaving the document",
  );
});

test("pin: unlocking restores the normal clear-on-empty behaviour", async () => {
  const state = { selectionText: "pinned passage", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  await sel.onSelectionChanged();
  sel.lockPin();
  sel.unlockPin();

  state.selectionText = "";
  await sel.onSelectionChanged();
  assert.equal(sel.getPinnedText(), "");
});

// A lock must not freeze the Pin against the user's own intent: deliberately
// selecting a different passage still re-pins.
test("pin: a locked Pin still moves when the user selects something else", async () => {
  const state = { selectionText: "first passage", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  await sel.onSelectionChanged();
  sel.lockPin();

  state.selectionText = "second passage";
  await sel.onSelectionChanged();
  assert.equal(sel.getPinnedText(), "second passage");
});

// ---- reconcile: document beats memory (bug #1) -----------------------------

test("reconcile: bookmark text wins over stale cached memory", async () => {
  const state = { selectionText: "original text", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  // Pin the selection → pinnedText = "original text"
  await sel.onSelectionChanged();
  assert.equal(sel.getPinnedText(), "original text");

  // User edits the pinned passage in-place: bookmark now holds new text.
  state.bookmarkText = "edited text";
  state.selectionText = ""; // selection collapsed (focus in taskpane)

  // getSelectionData calls reconcile() internally.
  const handle = sel.beginAsk();
  const data = await sel.getSelectionData();
  handle.end();

  assert.equal(
    sel.getPinnedText(),
    "edited text",
    "memory must follow the document",
  );
  assert.equal(data.type, "text");
  assert.equal(data.text, "edited text");
});

test("reconcile: no-op when nothing is pinned", async () => {
  const state = { selectionText: "", bookmarkText: null, bodyText: "full doc" };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  const handle = sel.beginAsk();
  const data = await sel.getSelectionData();
  handle.end();

  assert.equal(data.type, "fulldoc");
  assert.equal(data.text, "full doc");
});

// ---- captureTarget: immutable snapshot (bug #4) ----------------------------

test("captureTarget: text snapshot is frozen at Ask time", async () => {
  const state = { selectionText: "hello world", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  await sel.onSelectionChanged();

  const handle = sel.beginAsk();
  const data = await sel.getSelectionData();
  const target = sel.captureTarget(data);
  handle.end();

  assert.equal(target.kind, "text");
  assert.equal(target.text, "hello world");

  // Selecting something else afterwards must not mutate the captured target.
  state.selectionText = "something else";
  assert.equal(target.text, "hello world");
});

test("captureTarget: table target carries a stable sig", async () => {
  const cellTexts = [
    ["a", "b"],
    ["c", "d"],
  ];
  const table = {
    rowCount: 2,
    columnCount: 2,
    load: () => {},
    getCell: (r, c) => ({
      body: { getRange: () => ({ load: () => {}, text: cellTexts[r][c] }) },
    }),
  };
  const state = { selectionText: "a b c d", tables: [table] };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  const handle = sel.beginAsk();
  const data = await sel.getSelectionData();
  const target = sel.captureTarget(data);
  handle.end();

  assert.equal(target.kind, "table");
  assert.equal(
    target.sig,
    tableSignature({ rowCount: 2, columnCount: 2, values: cellTexts }),
  );
});

test("captureTarget: fulldoc target has kind fulldoc", async () => {
  const state = {
    selectionText: "",
    bookmarkText: null,
    bodyText: "whole doc",
  };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  const handle = sel.beginAsk();
  const data = await sel.getSelectionData();
  const target = sel.captureTarget(data);
  handle.end();

  assert.equal(target.kind, "fulldoc");
});

// ---- TargetMachine classification ------------------------------------------

test("empty selection in idle clears the pin", async () => {
  const state = { selectionText: "hello", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  await sel.onSelectionChanged(); // pin "hello"
  assert.equal(sel.getPinnedText(), "hello");
  // targetKind is only classified at Ask time (getSelectionData), not on pin.

  state.selectionText = ""; // user deselects
  await sel.onSelectionChanged(); // idle + empty → clear
  assert.equal(sel.getPinnedText(), "");
  assert.equal(sel.getTargetKind(), "empty");
});

test("tableSignature is stable for identical contents", () => {
  const t = {
    rowCount: 2,
    columnCount: 2,
    values: [
      ["a", "b"],
      ["c", "d"],
    ],
  };
  assert.equal(
    tableSignature(t),
    tableSignature({
      ...t,
      values: [
        ["a", "b"],
        ["c", "d"],
      ],
    }),
  );
});

test("tableSignature differs when contents differ", () => {
  const a = {
    rowCount: 2,
    columnCount: 2,
    values: [
      ["a", "b"],
      ["c", "d"],
    ],
  };
  const b = {
    rowCount: 2,
    columnCount: 2,
    values: [
      ["a", "b"],
      ["c", "X"],
    ],
  };
  assert.notEqual(tableSignature(a), tableSignature(b));
});

// ---- reset -----------------------------------------------------------------

test("reset clears pin and returns activity to idle", async () => {
  const state = { selectionText: "hello", bookmarkText: null };
  const runWord = (fn) => fn(makeContext(state));
  const sel = createSelectionMgr({ runWord });

  await sel.onSelectionChanged();
  assert.equal(sel.getPinnedText(), "hello");

  await sel.reset();
  assert.equal(sel.getPinnedText(), "");

  // After reset the activity gate is open again, so a new selection pins.
  state.selectionText = "new text";
  await sel.onSelectionChanged();
  assert.equal(sel.getPinnedText(), "new text");
});

// ---- countOccurrences (what the Proposal card promises) ---------------------

test("countOccurrences counts each find with the options Apply will use", async () => {
  const seen = [];
  const runWord = (fn) =>
    fn({
      sync: async () => {},
      document: {
        body: {
          search: (text, opts) => {
            seen.push([text, opts]);
            return {
              items: new Array(text === "the" ? 12 : 1).fill({}),
              load: () => {},
            };
          },
        },
      },
    });
  const counts = await createSelectionMgr({ runWord }).countOccurrences([
    "the",
    "xyz",
  ]);
  assert.equal(counts.get("the"), 12);
  assert.equal(counts.get("xyz"), 1);
  assert.deepEqual(seen[0][1], { matchCase: false });
});

test("countOccurrences skips finds Word.search would reject as too long", async () => {
  const long = "x".repeat(300);
  const runWord = (fn) =>
    fn({
      sync: async () => {},
      document: { body: { search: () => ({ items: [], load: () => {} }) } },
    });
  const counts = await createSelectionMgr({ runWord }).countOccurrences([long]);
  assert.equal(
    counts.has(long),
    false,
    "no entry means the card makes no claim",
  );
});

test("countOccurrences reports null rather than a wrong number when Word fails", async () => {
  const runWord = async () => {
    throw new Error("boom");
  };
  const counts = await createSelectionMgr({ runWord }).countOccurrences(["a"]);
  assert.equal(counts.get("a"), null);
});
