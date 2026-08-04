/* global Office, setTimeout */
// Selection state for the Word task pane, extracted from taskpane.js.
//
// Two ORTHOGONAL state machines live here (see docs/adr/0001):
//
//   TargetMachine  — what the user is aiming at: empty | text | table | multi
//   ActivityMachine — what the pane is doing:    idle | reading | applying
//
// The four historical selection bugs are each forbidden by structure:
//   #1 stale pin wins      → reconcile(): the document bookmark beats memory
//   #2 race vs in-flight pin → ActivityMachine gate: onSelectionChanged is a
//                            no-op unless activity === "idle"
//   #3 focus-loss clears pin → an empty selectionChanged only clears when idle
//   #4 apply targets wrong   → captureTarget(): an immutable snapshot handed to
//                            proposal-mgr, which never reads selection state back
//
// Word.run is injected as `runWord` (DI) so the module is unit-testable without
// the Office host — see selection-mgr.test.js. The module OWNS the selection
// gate: taskpane.js calls beginAsk()/beginApply() and never tracks it itself.

// Word's Range.search() rejects find strings longer than 255 characters.
export const MAX_SEARCH_LEN = 255;
// Keep full-document prompts bounded so large files do not exceed model context.
export const MAX_FULLDOC_CHARS = 60000;
// Bookmark used to anchor the user's selection in the document. A bookmark is a
// live document object, so re-reading it later is always accurate — unlike a
// captured text string, which can go stale. Exported so proposal-mgr can re-pin
// the replaced range after a text Apply.
export const PIN_BOOKMARK_NAME = "HermesPinnedSelection";

// Only transient Office error codes are retried; a real failure (permission
// denied, invalid argument, etc.) surfaces immediately.
const TRANSIENT_CODES = new Set(["ActivitySuspended", "ItemNotFound"]);

function supportsBookmarks() {
  try {
    return Office.context.requirements.isSetSupported("WordApi", "1.4");
  } catch {
    return false;
  }
}

// Content fingerprint of a table snapshot ({rowCount, columnCount, values}).
// Word's JS API gives no stable table identifier we can hold across runWord
// contexts, so the table's own pre-edit contents are what let Apply re-find the
// exact table a proposal was generated against.
export function tableSignature(t) {
  return `${t.rowCount}x${t.columnCount}:${JSON.stringify(t.values)}`;
}

/**
 * @param {{ runWord: (fn: (context: any) => Promise<any>) => Promise<any> }} deps
 */
export function createSelectionMgr({ runWord }) {
  // Word.run with retry — the first read right after a document switch can fail
  // with a transient "document busy / call cancelled" error, so retry before
  // giving up.
  function runWithRetry(fn, retries = 3) {
    return new Promise((resolve, reject) => {
      const attempt = (n) => {
        runWord(fn)
          .then(resolve)
          .catch((err) => {
            if (n <= 1 || !TRANSIENT_CODES.has(err && err.code)) reject(err);
            else setTimeout(() => attempt(n - 1), 350);
          });
      };
      attempt(retries);
    });
  }

  // ---- TargetMachine + ActivityMachine state -------------------------------
  let targetKind = "empty"; // empty | text | table | multi
  let activity = "idle"; // idle | reading | applying
  let pinnedText = ""; // cached text of the pinned selection (memory)
  let capturedText = ""; // snapshot taken at Ask time, fallback Apply target
  let selectionIsPinned = false;
  // Serialises the selectionChanged handler AND is what getSelectionData waits
  // on, so an Ask never starts while a pin is still being written.
  let pendingPin = Promise.resolve();
  // Set while a rendered Proposal is anchored to the Pin. See lockPin().
  let pinLocked = false;

  // ---- reading the selection (text / table / full doc) ---------------------

  function readSelectedText() {
    return runWithRetry((context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      return context.sync().then(() => (sel.text || "").trim());
    }).catch(() => "");
  }

  // Fallback whole-document reader using the Office.js v1 file API. It binds to
  // the live document directly, so it works even when the runWord context is
  // stale after opening a new file.
  function readFullBodyViaFile() {
    return new Promise((resolve) => {
      Office.context.document.getFileAsync(
        Office.FileType.Text,
        { sliceSize: 65536 },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded)
            return resolve("");
          const file = result.value;
          const sliceCount = file.sliceCount;
          let cur = 0;
          let text = "";
          const next = () => {
            if (cur >= sliceCount) {
              file.closeAsync();
              return resolve(text.trim());
            }
            file.getSliceAsync(cur, (slice) => {
              if (slice.status === Office.AsyncResultStatus.Succeeded) {
                text += slice.data;
                cur += 1;
                next();
              } else {
                file.closeAsync();
                resolve(text.trim());
              }
            });
          };
          next();
        },
      );
    });
  }

  async function readFullBody() {
    try {
      return await runWithRetry((context) => {
        const range = context.document.body.getRange();
        range.load("text");
        return context.sync().then(() => (range.text || "").trim());
      });
    } catch (e) {
      // runWord failed (stale context on a freshly opened file, etc.) — fall
      // back to the v1 file API before reporting a failure.
      const viaFile = await readFullBodyViaFile();
      if (viaFile) return viaFile;
      throw e;
    }
  }

  // ---- bookmark machinery (the document-side source of truth) --------------

  // Anchor the current selection with a bookmark so the exact range survives
  // focus leaving the document. Re-create on every change so only the LATEST
  // selection is pinned.
  async function pinCurrentSelection() {
    try {
      await runWithRetry(async (context) => {
        const sel = context.document.getSelection();
        sel.load("text");
        await context.sync();
        const text = (sel.text || "").trim();
        // Only (re)pin on a NON-empty selection. An empty/collapsed event also
        // fires when focus merely leaves the document — deleting the bookmark
        // there would wipe the pin a pending proposal depends on.
        if (text.length > 0) {
          pinnedText = text;
          selectionIsPinned = true;
          context.document.deleteBookmark(PIN_BOOKMARK_NAME);
          sel.insertBookmark(PIN_BOOKMARK_NAME);
          await context.sync();
        }
      });
    } catch {
      // Best-effort (bookmark APIs need WordApi 1.4); fallbacks still work.
    }
  }

  // Read the pinned selection back from its bookmark. Returns null when no
  // selection is pinned.
  async function getPinnedSelectionFromBookmark() {
    try {
      return await runWithRetry((context) => {
        const range =
          context.document.getBookmarkRangeOrNullObject(PIN_BOOKMARK_NAME);
        range.load(["isNullObject", "text"]);
        return context.sync().then(() => {
          if (range.isNullObject) return null;
          const t = (range.text || "").trim();
          return t.length > 0 ? t : null;
        });
      });
    } catch {
      return null;
    }
  }

  // Guarantee the bookmark holds exactly `expectedText`, regardless of stray
  // selectionChanged noise fired while focus moved into the taskpane.
  async function ensurePinnedBookmark(expectedText) {
    if (!expectedText || !supportsBookmarks()) return;
    try {
      await runWithRetry(async (context) => {
        const bmRange =
          context.document.getBookmarkRangeOrNullObject(PIN_BOOKMARK_NAME);
        bmRange.load(["isNullObject", "text"]);
        const sel = context.document.getSelection();
        sel.load("text");
        await context.sync();
        if (
          !bmRange.isNullObject &&
          (bmRange.text || "").trim() === expectedText
        )
          return;
        let target = null;
        if ((sel.text || "").trim() === expectedText) {
          target = sel;
        } else if (expectedText.length <= MAX_SEARCH_LEN) {
          const ranges = context.document.body.search(expectedText, {
            matchCase: false,
          });
          ranges.load("items");
          await context.sync();
          if (ranges.items.length > 0) target = ranges.items[0];
        }
        if (target) {
          context.document.deleteBookmark(PIN_BOOKMARK_NAME);
          target.insertBookmark(PIN_BOOKMARK_NAME);
          await context.sync();
        }
      });
    } catch {
      // Best-effort — Apply still has the bookmark/search fallbacks.
    }
  }

  async function clearSelectionBookmark() {
    try {
      await runWithRetry(async (context) => {
        context.document.deleteBookmark(PIN_BOOKMARK_NAME);
        await context.sync();
      });
    } catch {
      // best-effort cleanup
    }
  }

  // ---- selectionChanged handler (gated by ActivityMachine) -----------------

  async function handleSelectionChange() {
    // Re-checked here, not only in onSelectionChanged: an event that queued
    // behind an earlier one may reach the front after an Ask has started.
    if (activity !== "idle") return;

    const text = await readSelectedText();
    if (!text) {
      // Idle + empty selection. Normally the old pin is stale and gets cleared
      // — but "empty" is also what focus moving into the taskpane looks like,
      // and if a Proposal is on screen its Pin is the thing Apply will write
      // through. A Pin that does not survive focus leaving the document is not
      // doing the job the bookmark exists for.
      if (pinLocked) return;
      pinnedText = "";
      selectionIsPinned = false;
      capturedText = "";
      targetKind = "empty";
      await clearSelectionBookmark();
      return;
    }
    // A non-empty event is the user deliberately aiming somewhere else, so it
    // re-pins even under a lock. The pending Proposal is unaffected: its target
    // was frozen by captureTarget() at Ask time.
    await pinCurrentSelection();
  }

  async function onSelectionChanged() {
    // THE GATE (bug #2/#3): while reading or applying, ignore selectionChanged
    // entirely. Focus leaving the document fires an empty selection event that
    // would otherwise wipe the pin a pending Ask/Apply depends on.
    if (activity !== "idle") return;

    // Queue behind whatever is already in flight instead of racing it. Two
    // events arriving together used to run their read-then-pin sequences
    // concurrently; the Word.run batches interleave, so the OLDER selection's
    // insertBookmark could land last and win. Neither handler ever rejects, but
    // the chain is defended anyway so one failure cannot wedge the queue.
    pendingPin = pendingPin.then(handleSelectionChange, handleSelectionChange);
    await pendingPin;
  }

  // Note: the handler is never explicitly removed. The taskpane's JS realm is
  // torn down and reloaded on every document switch, so no previous handler
  // dangles.
  function registerSelectionChanged(afterChange) {
    runWord(async (context) => {
      context.document.onSelectionChanged.add(async () => {
        await onSelectionChanged();
        if (afterChange) afterChange();
      });
      await context.sync();
    }).catch(() => {});
  }

  // ---- reconcile: document beats memory (bug #1) ---------------------------

  // If the user edited the pinned passage's CONTENT directly (without changing
  // the selection range), the cached pinnedText is stale. The bookmark is a live
  // document object, so it is authoritative — pull memory back in line with it.
  // Called at the start of getSelectionData(), which runs inside beginAsk(), so
  // the gate guarantees no in-flight pin races this read.
  async function reconcile() {
    if (!selectionIsPinned) return;
    const bm = await getPinnedSelectionFromBookmark();
    if (bm && bm !== pinnedText) {
      pinnedText = bm;
      capturedText = bm;
    }
  }

  // ---- table detection -----------------------------------------------------

  async function readTableCells(context, tables) {
    const rows = [];
    for (const table of tables) {
      table.load(["rowCount", "columnCount"]);
      await context.sync();
      // Load every cell's range up front and sync ONCE per table, instead of
      // round-tripping per cell.
      const rangeRows = [];
      for (let r = 0; r < table.rowCount; r++) {
        const rangeRow = [];
        for (let c = 0; c < table.columnCount; c++) {
          const range = table.getCell(r, c).body.getRange();
          range.load("text");
          rangeRow.push(range);
        }
        rangeRows.push(rangeRow);
      }
      await context.sync();
      const values = rangeRows.map((rangeRow) =>
        rangeRow.map((range) => (range.text || "").trim()),
      );
      rows.push({
        rowCount: table.rowCount,
        columnCount: table.columnCount,
        values,
      });
    }
    return rows;
  }

  // Detect a table in the current selection. Loads are isolated so a missing
  // parentTable (cursor outside any table) never throws and breaks the flow.
  async function detectSelectedTable() {
    const fromSelection = await runWord(async (context) => {
      const sel = context.document.getSelection();
      sel.load("tables/items");
      await context.sync();
      // eslint-disable-next-line office-addins/load-object-before-read
      if (sel.tables.items.length === 0) return null;
      return readTableCells(context, sel.tables.items);
    }).catch(() => null);
    if (fromSelection) return fromSelection;

    const fromParent = await runWord(async (context) => {
      const sel = context.document.getSelection();
      sel.load("parentTable/isNullObject");
      await context.sync();
      if (sel.parentTable.isNullObject) return null;
      // eslint-disable-next-line office-addins/load-object-before-read
      return readTableCells(context, [sel.parentTable]);
    }).catch(() => null);
    if (fromParent) return fromParent;

    return null;
  }

  // ---- the main read (called inside beginAsk) ------------------------------

  async function getSelectionData() {
    // If a selectionChanged re-pin is still in flight (the user selected a
    // passage and clicked Ask right away), let it finish first — reading the
    // bookmark mid-pin would resolve the previously pinned passage.
    await pendingPin;
    await reconcile();

    const selText = await readSelectedText();

    // A table selection still carries text, so check tables FIRST.
    const tableRows = await detectSelectedTable();
    // Only enter table mode if the table actually has data. An empty
    // placeholder table must not hijack the flow.
    const tableHasData =
      tableRows &&
      tableRows.length > 0 &&
      tableRows.some((t) =>
        t.values.flat().some((v) => (v || "").trim().length > 0),
      );
    if (tableHasData && selText.length > 0) {
      if (tableRows.length > 1) {
        targetKind = "multi";
        return { type: "multi-table", text: "" };
      }
      targetKind = "table";
      return { type: "table", tables: tableRows, rawText: selText };
    }

    // The LIVE selection is authoritative at Ask time; the pinned bookmark is
    // only a fallback for when the selection collapsed as focus moved into the
    // taskpane. Preferring the bookmark unconditionally let a STALE pin win.
    const pinnedFromBookmark = await getPinnedSelectionFromBookmark();
    let effectiveSelText;
    if (selText.length > 0 && selText !== pinnedFromBookmark) {
      // Fresh selection the pin missed — re-anchor the bookmark on it now via
      // the live selection RANGE (body.search rejects >255-char finds).
      await pinCurrentSelection();
      effectiveSelText = selText;
    } else if (selText.length > 0) {
      effectiveSelText = selText;
    } else if (
      selectionIsPinned &&
      pinnedFromBookmark &&
      pinnedFromBookmark.length > 0
    ) {
      // Live selection empty but a valid pin exists (focus left the doc).
      effectiveSelText = pinnedFromBookmark;
    } else if (selectionIsPinned && pinnedText.trim().length > 0) {
      // Pin flag set but bookmark couldn't be read — fall back to cached text.
      effectiveSelText = pinnedText;
    } else {
      effectiveSelText = "";
    }

    if (effectiveSelText.length > 0) {
      capturedText = effectiveSelText;
      targetKind = "text";
      return {
        type: "text",
        text: effectiveSelText,
        capturedText: effectiveSelText,
      };
    }

    // Nothing selected → operate on the whole open document.
    const bodyText = await readFullBody();
    if (bodyText.length > 0) {
      targetKind = "empty";
      return { type: "fulldoc", text: bodyText };
    }

    targetKind = "empty";
    return { type: "empty", text: "" };
  }

  // ---- counting matches for the Proposal card ------------------------------

  // How many times each find string occurs in the body, using the SAME search
  // options Apply will use. The card has to state this before the user commits:
  // Apply replaces every match, and learning that from a toast mid-Apply is
  // learning it after the document already changed.
  //
  // Best-effort — on failure a find maps to null and the card simply makes no
  // claim about the count rather than a wrong one.
  async function countOccurrences(finds) {
    const counts = new Map();
    const wanted = [];
    for (const find of finds) {
      if (!find || find.length > MAX_SEARCH_LEN || counts.has(find)) continue;
      counts.set(find, null);
      wanted.push(find);
    }
    if (wanted.length === 0) return counts;
    try {
      await runWithRetry(async (context) => {
        const searches = wanted.map((find) => {
          const ranges = context.document.body.search(find, {
            matchCase: false,
          });
          ranges.load("items");
          return { find, ranges };
        });
        await context.sync();
        for (const { find, ranges } of searches)
          counts.set(find, ranges.items.length);
      });
    } catch {
      // Leave the reserved nulls in place.
    }
    return counts;
  }

  // ---- immutable snapshot for proposal-mgr (bug #4) ------------------------

  // Captured at Ask time and handed to proposal-mgr, which NEVER reads selection
  // state back. For tables this carries the pre-edit signature Apply matches
  // against, so selecting a different table before Apply can't redirect the
  // write. For text it carries the exact passage the proposal was generated for.
  function captureTarget(selectionData) {
    if (selectionData.type === "text") {
      return { kind: "text", text: capturedText };
    }
    if (selectionData.type === "table") {
      return {
        kind: "table",
        sig: tableSignature(selectionData.tables[0]),
        table: selectionData.tables[0],
      };
    }
    return { kind: "fulldoc" };
  }

  // ---- lifecycle (owned by this module, called by taskpane) ----------------

  function beginAsk() {
    activity = "reading";
    return { end: () => (activity = "idle") };
  }

  function beginApply() {
    activity = "applying";
    return { end: () => (activity = "idle") };
  }

  // Hold the Pin against the empty selectionChanged that focus leaving the
  // document produces. The taskpane locks once a text Proposal is on screen and
  // unlocks when that Proposal is applied or discarded — the Pin is what Apply
  // writes through, so it has to outlive the pane taking focus.
  function lockPin() {
    pinLocked = true;
  }

  function unlockPin() {
    pinLocked = false;
  }

  // Drop the in-memory pin + document bookmark WITHOUT touching activity — used
  // after a text Apply, whose replacement invalidates the original passage. An
  // explicit clear overrides the lock; that is the point of it being explicit.
  async function clearPin() {
    pinLocked = false;
    pinnedText = "";
    selectionIsPinned = false;
    capturedText = "";
    targetKind = "empty";
    await clearSelectionBookmark();
  }

  // Full reset for a new chat: clears the pin and returns activity to idle.
  async function reset() {
    activity = "idle";
    await clearPin();
  }

  function getTargetKind() {
    return targetKind;
  }

  function getPinnedText() {
    return pinnedText;
  }

  return {
    getSelectionData,
    countOccurrences,
    captureTarget,
    beginAsk,
    beginApply,
    lockPin,
    unlockPin,
    clearPin,
    reset,
    ensurePinnedBookmark,
    registerSelectionChanged,
    onSelectionChanged,
    getTargetKind,
    getPinnedText,
    readSelectedText,
  };
}
