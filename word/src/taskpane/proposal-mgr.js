/* global Office */
// Apply logic for the Word task pane, extracted from taskpane.js.
//
// This module is constructed with an IMMUTABLE `target` snapshot (from
// selection-mgr.captureTarget()) and NEVER reads selection state back. That is
// what kills bug #4 (Apply writing into whichever table/range happens to be
// selected at click time): the target is frozen at Ask time.
//
// Word.run is injected as `runWord` (DI) so the module is unit-testable without
// the Office host — see proposal-mgr.test.js.

import { columnLettersToIndex } from "../../../shared/parsers.js";
import { MAX_SEARCH_LEN, PIN_BOOKMARK_NAME } from "./selection-mgr.js";

function supportsBookmarks() {
  try {
    return Office.context.requirements.isSetSupported("WordApi", "1.4");
  } catch {
    return false;
  }
}

export function cellRefToPosition(cellRef, rowCount, columnCount) {
  const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return null;
  const col = columnLettersToIndex(match[1].toUpperCase());
  const row = parseInt(match[2], 10) - 1;
  if (col < 0 || col >= columnCount || row < 0 || row >= rowCount) return null;
  return { row, col };
}

// Detect edits whose replacement text contains another edit's find string —
// applying those in sequence would cascade unpredictably.
export function hasChainedEdits(edits) {
  return edits.some(
    (edit, i) =>
      edit.replace &&
      edits.some((other, j) => i !== j && edit.replace.includes(other.find)),
  );
}

/**
 * @param {{
 *   target: { kind: "text"|"table"|"fulldoc", text?: string, sig?: string, table?: object },
 *   runWord: (fn: (context: any) => Promise<any>) => Promise<any>,
 *   showToast?: (text: string, opts?: object) => void,
 * }} deps
 */
export function createProposalMgr({ target, runWord, showToast }) {
  const notify = showToast || (() => {});

  // Locate the table whose current contents still match the target signature.
  // Checks the tables in the live selection first (common case, cheap), then
  // falls back to scanning the body. Returns null when no table matches, which
  // the caller turns into a "select the table again" error rather than writing
  // into whatever table happens to be selected now.
  async function findProposalTable(context, sig) {
    const sel = context.document.getSelection();
    sel.load("tables/items");
    const body = context.document.body;
    body.load("tables/items");
    await context.sync();

    const matches = [];
    for (const table of new Set(sel.tables.items.concat(body.tables.items))) {
      table.load(["rowCount", "columnCount"]);
      await context.sync();
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
      const values = rangeRows.map((row) => row.map((range) => (range.text || "").trim()));
      const snapSig = `${table.rowCount}x${table.columnCount}:${JSON.stringify(values)}`;
      if (snapSig === sig) matches.push(table);
    }
    // Same contents can appear in multiple tables. Picking the first match would
    // write a proposal into an unrelated table.
    return matches.length === 1 ? matches[0] : null;
  }

  /**
   * Apply a table proposal. `changes` are {cell, value, old} records.
   * @returns {Promise<{applied: number}>}
   */
  async function applyTable(changes, { markRed }) {
    await runWord(async (context) => {
      const table = await findProposalTable(context, target.sig);
      if (!table)
        throw new Error(
          "Không còn tìm thấy bảng của đề xuất này (bảng đã bị sửa hoặc xoá). Hãy chọn lại bảng rồi hỏi lại.",
        );
      table.load(["rowCount", "columnCount"]);
      await context.sync();

      const cells = [];
      for (const change of changes) {
        const pos = cellRefToPosition(change.cell, table.rowCount, table.columnCount);
        if (!pos || change.old === undefined)
          throw new Error("Đề xuất có ô không hợp lệ. Hãy hỏi lại Hermes.");
        const range = table.getCell(pos.row, pos.col).body.getRange();
        range.load("text");
        cells.push({ change, range });
      }
      await context.sync();
      if (cells.some(({ change, range }) => (range.text || "").trim() !== change.old))
        throw new Error("Bảng đã thay đổi sau khi tạo đề xuất. Hãy hỏi lại Hermes.");
      for (const { change, range } of cells) {
        const inserted = range.insertText(String(change.value), "Replace");
        if (markRed) inserted.font.color = "#FF0000";
      }
      await context.sync();
    });
    return { applied: changes.length };
  }

  /**
   * Apply inline search-and-replace edits across the whole document. Each edit
   * is applied and sync'd independently so one bad edit (search text too long,
   * or no longer found) can't abort the rest of the batch.
   * @returns {Promise<{applied: number, skipped: number}>}
   */
  async function applyFulldocEdits(edits, { markRed }) {
    if (hasChainedEdits(edits))
      throw new Error("Đề xuất có các chỉnh sửa chồng chéo. Hãy hỏi lại Hermes.");
    return runWord(async (context) => {
      let applied = 0;
      let skipped = 0;
      for (const edit of edits) {
        if (!edit.find || edit.find.length > MAX_SEARCH_LEN) {
          skipped++;
          continue;
        }
        try {
          const ranges = context.document.body.search(edit.find, { matchCase: false });
          ranges.load("items");
          await context.sync();
          // A search that matched nothing changed the document in no way, so it
          // must count as skipped.
          if (ranges.items.length === 0) {
            skipped++;
            continue;
          }
          if (ranges.items.length > 1) {
            notify(`"${edit.find}" xuất hiện ${ranges.items.length} lần — sẽ thay TẤT CẢ.`, {
              tone: "warn",
              timeout: 4000,
            });
          }
          ranges.items.forEach((r) => {
            const inserted = r.insertText(String(edit.replace), "Replace");
            if (markRed) inserted.font.color = "#FF0000";
          });
          await context.sync();
          applied++;
        } catch {
          skipped++;
        }
      }
      return { applied, skipped };
    });
  }

  /**
   * Apply a text proposal: replace the pinned passage with `text`. Prefers the
   * bookmark if it still holds the captured passage; otherwise falls back to a
   * unique search for the captured text. Re-pins the inserted text so a
   * follow-up "rewrite it again" has something anchored.
   * @returns {Promise<{applied: number}>}
   */
  async function applyText(text, { markRed }) {
    const captured = target.text;
    await runWord(async (context) => {
      // Bookmark APIs need WordApi 1.4 — on older hosts skip to the search
      // fallback instead of throwing mid-batch.
      const canUseBookmark = supportsBookmarks();
      let bmRange = null;
      let bmText = null;
      if (canUseBookmark) {
        bmRange = context.document.getBookmarkRangeOrNullObject(PIN_BOOKMARK_NAME);
        bmRange.load(["isNullObject", "text"]);
        await context.sync();
        if (!bmRange.isNullObject) bmText = (bmRange.text || "").trim();
      }
      const bookmarkMatchesProposal =
        bmText !== null && (!captured || bmText === captured);
      let inserted = null;
      if (bookmarkMatchesProposal) {
        inserted = bmRange.insertText(text, "Replace");
      } else if (captured && captured.length <= MAX_SEARCH_LEN) {
        const ranges = context.document.body.search(captured, { matchCase: false });
        ranges.load("items");
        await context.sync();
        if (ranges.items.length === 1) {
          inserted = ranges.items[0].insertText(text, "Replace");
        } else {
          throw new Error(
            "Không thể xác định duy nhất đoạn văn bản gốc. Hãy chọn lại đoạn đó rồi hỏi lại Hermes.",
          );
        }
      } else if (captured) {
        // Too long to search for and the bookmark no longer matches it.
        throw new Error(
          "Đoạn văn bản gốc không còn được ghim. Hãy chọn lại đoạn cần thay rồi hỏi lại.",
        );
      } else {
        throw new Error("Không có văn bản được chọn để áp dụng.");
      }
      if (markRed) inserted.font.color = "#FF0000";
      // Re-pin the inserted text: replacing a bookmarked range removes the
      // bookmark, so without this a follow-up rewrite would have nothing pinned.
      if (canUseBookmark) inserted.insertBookmark(PIN_BOOKMARK_NAME);
      await context.sync();
    });
    return { applied: 1 };
  }

  return { applyTable, applyFulldocEdits, applyText };
}
