/* global Office, Word */

import { askHermes } from "../shared/hermes.js";
// Pure helpers deduped with the Excel taskpane via the repo-root shared/
// folder (no npm workspace between the two add-ins) — see shared/parsers.js.
import { columnIndexToLetters, columnLettersToIndex, parseEdits, parseTableChanges } from "../../../shared/parsers.js";

// Full-document mode sends the whole open document to Hermes as context;
// cap it so a large file doesn't blow up the request payload / token budget.
const MAX_FULLDOC_CHARS = 30000;

// Bookmark used to anchor the user's selection in the document at the moment
// they make it. A bookmark is a live document object, so re-reading its text
// (or range) later is always accurate — unlike a captured text string, which
// can go stale when the selectionChanged proxy resolves against an earlier
// snapshot. This is what makes "select A, rewrite, then select B, rewrite"
// reliably target B instead of revisiting A.
const PIN_BOOKMARK_NAME = "HermesPinnedSelection";

Office.onReady().then(() => {
  const log = document.getElementById("log");
  const input = document.getElementById("prompt");
  const askBtn = document.getElementById("ask");
  const applyBtn = document.getElementById("apply");
  const markRedWrap = document.getElementById("markRedWrap");
  const newChatBtn = document.getElementById("newchat");
  const statusEl = document.getElementById("status");
  const preview = document.getElementById("preview");

  let messages = [];
  let lastProposal = null;
  // Snapshot of the text range captured at send time, used as a fallback
  // target if the user's selection is lost by the time they click Apply.
  let capturedSelectionText = "";
  // Selection captured at the moment it happens in the document (via
  // selectionChanged), NOT at click time — clicking the Ask button moves
  // focus into the taskpane and drops the in-document selection, so reading
  // it fresh inside sendMessage() would always see "" and fall back to
  // treating the request as whole-document.
  let pinnedSelectionText = "";
  // The in-flight pinCurrentSelection() promise. Ask awaits it before reading
  // the bookmark, so selecting a passage and clicking Ask immediately can't
  // read the pin mid-flight and resolve the PREVIOUS passage.
  let pendingPin = Promise.resolve();

  function addMsg(role, text) {
    const div = document.createElement("div");
    div.className = `msg ${role === "user" ? "user" : "bot"}`;
    div.textContent = text;
    log.appendChild(div);
    log.scrollTop = log.scrollHeight;
  }

  function setStatus(text) {
    statusEl.textContent = text;
  }

  // Escapes text before it is interpolated into an innerHTML template, so
  // model output or document text can never inject markup into the preview.
  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  // ---- reading the selection (text / table / full doc) ---------------------

  // Word.run with retry — the first read right after a document switch can
  // fail with a transient "document busy / call cancelled" error, so retry
  // before giving up.
  function wordRunWithRetry(fn, retries = 3) {
    return new Promise((resolve, reject) => {
      const attempt = (n) => {
        Word.run(fn).then(resolve).catch((err) => {
          if (n <= 1) reject(err);
          else setTimeout(() => attempt(n - 1), 350);
        });
      };
      attempt(retries);
    });
  }

  function readSelectedText() {
    return wordRunWithRetry((context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      return context.sync().then(() => (sel.text || "").trim());
    }).catch(() => "");
  }

  // Fallback whole-document reader using the Office.js v1 file API. It binds
  // to the live document directly, so it works even when Word.run's context
  // is stale after opening a new file.
  function readFullBodyViaFile() {
    return new Promise((resolve) => {
      Office.context.document.getFileAsync(
        Office.FileType.Text,
        { sliceSize: 65536 },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) return resolve("");
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
        }
      );
    });
  }

  async function readFullBody() {
    try {
      return await wordRunWithRetry((context) => {
        const range = context.document.body.getRange();
        range.load("text");
        return context.sync().then(() => (range.text || "").trim());
      });
    } catch (e) {
      // Word.run failed (stale context on a freshly opened file, etc.) —
      // fall back to the v1 file API before reporting a failure.
      const viaFile = await readFullBodyViaFile();
      if (viaFile) return viaFile;
      throw e;
    }
  }

  // Anchor the current selection in the document with a bookmark so the exact
  // range survives focus leaving the document. Re-create it on every change so
  // only the LATEST selection is pinned — this is what keeps "select B, rewrite"
  // from re-targeting the previously selected passage A.
  async function pinCurrentSelection() {
    try {
      await wordRunWithRetry(async (context) => {
        const sel = context.document.getSelection();
        sel.load("text");
        await context.sync();
        const text = (sel.text || "").trim();
        // Only (re)pin on a NON-empty selection. An empty/collapsed selection
        // event also fires when focus merely leaves the document (e.g. the
        // user clicks Ask or Apply in the taskpane) — deleting the bookmark
        // there would wipe the pin the pending proposal depends on. Stale
        // pins are cleared by newChat() / replaced by the next real selection.
        if (text.length > 0) {
          pinnedSelectionText = text;
          // deleteBookmark is a no-op when the bookmark doesn't exist;
          // delete-then-insert in one batch replaces any previous pin.
          context.document.deleteBookmark(PIN_BOOKMARK_NAME);
          sel.insertBookmark(PIN_BOOKMARK_NAME);
          await context.sync();
        }
      });
    } catch (_) {
      // Selection pinning is best-effort (the bookmark APIs need WordApi 1.4);
      // the pinnedSelectionText / fresh-read fallbacks still work.
    }
  }

  // Read the text of the pinned selection back from its bookmark. Returns null
  // when no selection is pinned (empty selection, brand-new chat, etc.).
  async function getPinnedSelectionFromBookmark() {
    try {
      return await wordRunWithRetry((context) => {
        // Note: the Word JS API has no document.bookmarks collection (that's
        // the VBA object model) — bookmarks are reached via this Document
        // method, which returns the bookmark's Range directly.
        const range = context.document.getBookmarkRangeOrNullObject(PIN_BOOKMARK_NAME);
        range.load(["isNullObject", "text"]);
        return context.sync().then(() => {
          if (range.isNullObject) return null;
          const t = (range.text || "").trim();
          return t.length > 0 ? t : null;
        });
      });
    } catch (_) {
      return null;
    }
  }

  // Make the pin authoritative at Ask time: guarantee the bookmark holds
  // exactly the passage being sent to Hermes, regardless of any stray
  // selectionChanged noise fired while focus moved into the taskpane. If the
  // bookmark already matches, this is a no-op; otherwise re-anchor it on the
  // live selection (or a search hit) whose text matches what was captured.
  async function ensurePinnedBookmark(expectedText) {
    if (!expectedText) return;
    // Bookmark APIs need WordApi 1.4 — pre-1.4 hosts rely on the
    // captured-text search fallback at Apply time instead.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) return;
    try {
      await wordRunWithRetry(async (context) => {
        const bmRange = context.document.getBookmarkRangeOrNullObject(PIN_BOOKMARK_NAME);
        bmRange.load(["isNullObject", "text"]);
        const sel = context.document.getSelection();
        sel.load("text");
        await context.sync();
        if (!bmRange.isNullObject && (bmRange.text || "").trim() === expectedText) return;
        let target = null;
        if ((sel.text || "").trim() === expectedText) {
          target = sel;
        } else if (expectedText.length <= MAX_SEARCH_LEN) {
          const ranges = context.document.body.search(expectedText, { matchCase: false });
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
    } catch (_) {
      // Best-effort — Apply still has the bookmark/search fallbacks.
    }
  }

  // Remove the pin bookmark (e.g. on a fresh chat, or when the flow doesn't
  // need it) so it never lingers in the document.
  async function clearSelectionBookmark() {
    try {
      await wordRunWithRetry(async (context) => {
        // deleteBookmark is a no-op when the bookmark doesn't exist.
        context.document.deleteBookmark(PIN_BOOKMARK_NAME);
        await context.sync();
      });
    } catch (_) {
      // best-effort cleanup
    }
  }

  // Fires on every in-document selection change; pins the new selection with a
  // bookmark so getSelectionData() later resolves the EXACT passage the user
  // just selected. Runs outside a Word.run context, so it opens its own.
  async function onSelectionChangedHandler() {
    // pinCurrentSelection never rejects (it catches internally), so this
    // promise is always safe to await from getSelectionData().
    pendingPin = pinCurrentSelection();
    await pendingPin;
  }

  // Note: this handler is never explicitly removed (no
  // context.document.onSelectionChanged.remove call anywhere). That's
  // intentional — the taskpane's JS realm is torn down and reloaded fresh
  // every time the user switches/opens a document, so there is no previous
  // handler left dangling to leak.
  function registerSelectionChangedHandler() {
    Word.run(async (context) => {
      context.document.onSelectionChanged.add(onSelectionChangedHandler);
      await context.sync();
    }).catch(() => {});
  }

  async function getSelectionData() {
    // If a selectionChanged re-pin is still in flight (the user selected a
    // passage and clicked Ask right away), let it finish first — reading the
    // bookmark mid-pin would resolve the previously pinned passage.
    await pendingPin;

    const selText = await readSelectedText();

    // A table selection still carries text, so check tables FIRST.
    const tableRows = await detectSelectedTable();
    // Only enter table mode if the table actually has data. An empty
    // placeholder table (e.g. a lone blank row) must not hijack the flow
    // and hide the rest of the document — fall through to full-doc instead.
    const tableHasData = tableRows && tableRows.length > 0 &&
      tableRows.some((t) => t.values.flat().some((v) => (v || "").trim().length > 0));
    if (tableHasData) {
      return { type: "table", tables: tableRows, rawText: selText };
    }

    // The LIVE selection is authoritative at Ask time; the pinned bookmark is
    // only a fallback for when the selection collapsed as focus moved into
    // the taskpane. Preferring the bookmark unconditionally let a STALE pin
    // win — e.g. Apply re-anchors the bookmark on the previous passage's
    // rewritten text, so if the pin handler hadn't caught up with a newly
    // selected passage yet, Ask would silently target the old one again.
    const pinnedFromBookmark = await getPinnedSelectionFromBookmark();
    let effectiveSelText;
    if (selText.length > 0 && selText !== pinnedFromBookmark) {
      // Fresh selection the pin missed (or hasn't caught up with) — re-anchor
      // the bookmark on it now via the live selection RANGE, which works for
      // any length (body.search rejects find strings over 255 chars). This
      // also updates pinnedSelectionText.
      await pinCurrentSelection();
      effectiveSelText = selText;
    } else if (pinnedFromBookmark && pinnedFromBookmark.length > 0) {
      effectiveSelText = pinnedFromBookmark;
    } else if (pinnedSelectionText.trim().length > 0) {
      effectiveSelText = pinnedSelectionText;
    } else {
      effectiveSelText = selText;
    }

    // Plain text selection (even short) → operate on exactly that text.
    if (effectiveSelText.length > 0) {
      capturedSelectionText = effectiveSelText;
      return { type: "text", text: effectiveSelText };
    }

    // Nothing selected → operate on the whole open document.
    const bodyText = await readFullBody();
    if (bodyText.length > 0) {
      return { type: "fulldoc", text: bodyText };
    }

    return { type: "empty", text: "" };
  }

  async function readTableCells(context, tables) {
    const rows = [];
    for (const table of tables) {
      table.load(["rowCount", "columnCount"]);
      await context.sync();
      // Load every cell's range up front and sync ONCE per table, instead of
      // round-tripping per cell — and read .text off the same range object
      // that was loaded (a fresh getRange() call returns an unloaded proxy).
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
      const values = rangeRows.map((rangeRow) => rangeRow.map((range) => (range.text || "").trim()));
      rows.push({ rowCount: table.rowCount, columnCount: table.columnCount, values });
    }
    return rows;
  }

  // Detect a table in the current selection. Loads are isolated so a missing
  // parentTable (cursor outside any table) never throws and breaks the flow.
  async function detectSelectedTable() {
    const fromSelection = await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.tables.load("items");
      await context.sync();
      if (sel.tables.items.length === 0) return null;
      return readTableCells(context, sel.tables.items);
    }).catch(() => null);
    if (fromSelection) return fromSelection;

    const fromParent = await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.parentTable.load("isNullObject");
      await context.sync();
      if (sel.parentTable.isNullObject) return null;
      return readTableCells(context, [sel.parentTable]);
    }).catch(() => null);
    if (fromParent) return fromParent;

    return null;
  }

  function formatTablePreview(tbl) {
    const flat = tbl.values.map((row, ri) => {
      const cells = row.map((cell, ci) => `  [${columnIndexToLetters(ci)}${ri + 1}] ${cell}`).join("\n");
      return `Row ${ri + 1}:\n${cells}`;
    }).join("\n\n");
    return `${tbl.rowCount} rows x ${tbl.columnCount} cols\n\n${flat}`;
  }

  // ---- system prompt --------------------------------------------------------

  function buildSystemPrompt(data) {
    if (!data || data.type === "empty") {
      return "You are editing a Word document. The user has not selected any text. Generate the content they request. Respond with the text only — no markdown fences, no explanations.";
    }

    if (data.type === "table") {
      let tableDesc = "";
      if (data.tables && data.tables.length > 0) {
        tableDesc = data.tables.map((t, i) =>
          `TABLE ${i + 1} (${t.rowCount} rows x ${t.columnCount} cols):\n${formatTablePreview(t)}`
        ).join("\n\n");
      }
      return `You are editing a table in a Word document. Each cell is labeled [ColRow] (e.g. [A1] = first column, first row). Below is the current table data. The user will ask you to modify it.

When returning edited table data, output EXACTLY this JSON format at the end of your reply, on a new line:
\`\`\`json
{"cells": [{"cell": "A1", "value": "new value"}, ...]}
\`\`\`
Only include cells that changed. Also give a brief natural-language reply explaining what you did.

CURRENT TABLE DATA:
${tableDesc}`;
    }

    // Full-document mode (nothing selected) — apply changes to the WHOLE file,
    // but ONLY the spots that change. Never paste a review list over the doc.
    if (data.type === "fulldoc") {
      return `You are editing a Word document. The user did NOT select any text — they want changes applied to the ENTIRE open document below.

INSTRUCTIONS — pick the match:

- FIX / CORRECT / REPLACE text (spelling, grammar, wording, reformat): you MUST output a fenced JSON block at the end of your reply:
\`\`\`json
{"edits":[{"find":"wrong text exactly as written","replace":"corrected text"}]}
\`\`\`
  • List EVERY change. Apply to ALL matching occurrences (case-insensitive), INLINE, preserving formatting.
  • Do NOT rewrite or repeat the whole document. Do NOT output a separate prose list of errors.
  • One short sentence above the JSON is fine.

- REWRITE or TRANSLATE the entire document: output the COMPLETE revised document text only (no JSON).

- Only a QUESTION or a review with NO changes wanted: answer in plain text (no JSON).

CURRENT DOCUMENT:
${data.text}`;
    }

    return `You are editing a Word document. The user has selected a SPECIFIC passage and wants ONLY that passage modified.

CRITICAL RULES:
- Rewrite ONLY the selected text below. Do NOT touch, repeat, or regenerate the rest of the document.
- Respond with the edited passage only — no explanations, no markdown fences, no commentary. Just the final text ready to paste into the document.
- Keep the same language, tone, and formatting style as the original.

SELECTED TEXT:
${data.text}`;
  }

  // ---- send / apply ---------------------------------------------------------

  async function sendMessage() {
    const userText = input.value.trim();
    if (!userText) return;

    addMsg("user", userText);
    input.value = "";
    askBtn.disabled = true;
    setStatus("Reading document…");
    lastProposal = null;
    applyBtn.style.display = "none";
    markRedWrap.style.display = "none";
    preview.innerHTML = "";

    try {
      const selectionData = await getSelectionData();

      // Re-assert the pin for exactly the text being sent to Hermes, so a
      // stray empty selectionChanged (focus moving into the taskpane) can
      // never leave Apply without an anchor for this proposal.
      if (selectionData.type === "text") {
        await ensurePinnedBookmark(capturedSelectionText);
      }

      if (selectionData.type === "empty") {
        addMsg("bot", "Tài liệu trống hoặc không đọc được. Hãy chọn một đoạn văn bản, hoặc gõ nội dung cần xử lý.");
        setStatus("Không có văn bản (doc trống / chưa sync xong). Thử lại.");
        return;
      }

      let docText = selectionData.text || "";
      const fulldocTruncated = selectionData.type === "fulldoc" && docText.length > MAX_FULLDOC_CHARS;
      if (fulldocTruncated) {
        docText = docText.slice(0, MAX_FULLDOC_CHARS) + `\n\n[... còn ${docText.length - MAX_FULLDOC_CHARS} ký tự nữa ...]`;
      }
      const displayData = { ...selectionData, text: docText };

      const statusText = selectionData.type === "fulldoc"
        ? `${selectionData.text.length} chars from full document`
        : (selectionData.text ? `${selectionData.text.length} chars selected` : "Table selected");
      const truncationWarning = fulldocTruncated
        ? ` ⚠ Tài liệu vượt quá ${MAX_FULLDOC_CHARS} ký tự — chỉ ${MAX_FULLDOC_CHARS} ký tự đầu được gửi tới Hermes.`
        : "";
      setStatus(statusText + " — Hermes is thinking…" + truncationWarning);

      const sysPrompt = buildSystemPrompt(displayData);
      const payload = [
        { role: "system", content: sysPrompt },
        ...messages,
        { role: "user", content: userText },
      ];

      setStatus("Hermes is thinking…");
      const reply = await askHermes(payload);
      addMsg("bot", reply);

      if (selectionData.type === "table") {
        const tableChanges = parseTableChanges(reply);
        if (tableChanges.length > 0) {
          lastProposal = { type: "table", changes: tableChanges };
          preview.innerHTML = tableChanges.map(c =>
            `<div class="act">Set ${escapeHtml(c.cell)} → "${escapeHtml(c.value)}"</div>`
          ).join("");
          applyBtn.style.display = "block";
          markRedWrap.style.display = "flex";
        }
      } else if (selectionData.type === "text") {
        lastProposal = { type: "text", text: reply };
        preview.innerHTML = `
          <div class="act">Proposed edit:</div>
          <div class="msg bot">${escapeHtml(reply).replace(/\n/g, "<br>")}</div>
        `;
        applyBtn.style.display = "block";
        markRedWrap.style.display = "flex";
      } else if (selectionData.type === "fulldoc") {
        // Prefer inline edits (fix spelling etc.) — only the wrong spots change.
        const edits = parseEdits(reply);
        if (edits.length > 0) {
          lastProposal = { type: "fulldoc-edits", edits };
          preview.innerHTML = edits.map(e =>
            `<div class="act">"${escapeHtml(e.find)}" → "${escapeHtml(e.replace)}"</div>`
          ).join("") + `<div class="act">Áp dụng cho mọi vị trí (giữ nguyên định dạng).</div>`;
          applyBtn.style.display = "block";
          markRedWrap.style.display = "flex";
        } else if (reply.trim().length > selectionData.text.length * 0.5) {
          // Looks like a full rewrite / translation → replace whole doc.
          lastProposal = { type: "fulldoc-full", text: reply };
          preview.innerHTML = `<div class="act">Sẽ THAY THẾ TOÀN BỘ văn bản bằng kết quả trên.</div>`;
          applyBtn.style.display = "block";
          markRedWrap.style.display = "flex";
        } else {
          // Plain answer / review — nothing to apply.
          lastProposal = null;
          preview.innerHTML = `<div class="act">Đây là câu trả lời / nhận xét — không áp dụng trực tiếp.</div>`;
          applyBtn.style.display = "none";
          markRedWrap.style.display = "none";
        }
      }

      messages.push({ role: "user", content: userText });
      messages.push({ role: "assistant", content: reply });
      setStatus("Ready.");
    } catch (err) {
      const errMsg = err.message || String(err);
      addMsg("bot", errMsg);
      setStatus("Error.");
    } finally {
      askBtn.disabled = false;
      input.focus();
    }
  }

  function cellRefToPosition(cellRef, rowCount, columnCount) {
    const match = cellRef.match(/^([A-Z]+)(\d+)$/i);
    if (!match) return null;
    const col = columnLettersToIndex(match[1].toUpperCase());
    const row = parseInt(match[2], 10) - 1;
    if (col < 0 || col >= columnCount || row < 0 || row >= rowCount) return null;
    return { row, col };
  }

  // Word's Range.search() rejects find strings longer than 255 characters.
  const MAX_SEARCH_LEN = 255;

  async function applyEdit() {
    if (!lastProposal) return;

    // Đọc trạng thái toggle "đánh dấu đỏ" một lần — nếu bật, mọi đoạn văn bản
    // được chèn/sửa trong lượt Apply này sẽ được tô màu đỏ để dễ nhận biết.
    const markRedEl = document.getElementById("markRed");
    const markRed = markRedEl && markRedEl.checked;

    setStatus("Applying…");
    let editStats = null;
    try {
      if (lastProposal.type === "table") {
        await Word.run(async (context) => {
          const sel = context.document.getSelection();
          sel.tables.load("items");
          await context.sync();
          const table = sel.tables.items[0];
          if (!table) throw new Error("Không tìm thấy bảng đã chọn. Hãy chọn lại bảng rồi Apply.");
          table.load(["rowCount", "columnCount"]);
          await context.sync();

          for (const change of lastProposal.changes) {
            const pos = cellRefToPosition(change.cell, table.rowCount, table.columnCount);
            if (!pos) continue;
            const cell = table.getCell(pos.row, pos.col);
            const inserted = cell.body.getRange().insertText(String(change.value), "Replace");
            if (markRed) inserted.font.color = "#FF0000";
          }
          await context.sync();
        });
      } else if (lastProposal.type === "fulldoc-edits") {
        // Inline search-and-replace for each edit — only wrong spots change,
        // surrounding formatting is preserved. Each edit is applied and
        // sync'd independently so one bad edit (search text too long, or no
        // longer found) can't abort the rest of the batch.
        editStats = await Word.run(async (context) => {
          let applied = 0;
          let skipped = 0;
          for (const edit of lastProposal.edits) {
            if (!edit.find || edit.find.length > MAX_SEARCH_LEN) { skipped++; continue; }
            try {
              const ranges = context.document.body.search(edit.find, { matchCase: false });
              ranges.load("items");
              await context.sync();
              ranges.items.forEach((r) => {
                const inserted = r.insertText(String(edit.replace), "Replace");
                if (markRed) inserted.font.color = "#FF0000";
              });
              await context.sync();
              applied++;
            } catch (_) {
              skipped++;
            }
          }
          return { applied, skipped };
        });
      } else if (lastProposal.type === "fulldoc-full") {
        await Word.run(async (context) => {
          const range = context.document.body.getRange();
          const inserted = range.insertText(lastProposal.text, "Replace");
          if (markRed) inserted.font.color = "#FF0000";
          await context.sync();
        });
      } else {
        // Plain text: replace the pinned selection if its bookmark is still in
        // the document AND still holds the passage this proposal was generated
        // for — the pin tracks the LATEST selection, so if the user selected a
        // different passage after asking, blindly inserting into the bookmark
        // would overwrite the wrong text. On mismatch (or a lost bookmark),
        // fall back to a search for the captured text so Apply never fails
        // just because focus moved to the task pane.
        await Word.run(async (context) => {
          // Bookmark APIs need WordApi 1.4 — on older hosts skip straight to
          // the search fallback instead of throwing mid-batch.
          const canUseBookmark = Office.context.requirements.isSetSupported("WordApi", "1.4");
          let bmRange = null;
          let bmText = null;
          if (canUseBookmark) {
            bmRange = context.document.getBookmarkRangeOrNullObject(PIN_BOOKMARK_NAME);
            bmRange.load(["isNullObject", "text"]);
            await context.sync();
            if (!bmRange.isNullObject) bmText = (bmRange.text || "").trim();
          }
          const bookmarkMatchesProposal =
            bmText !== null && (!capturedSelectionText || bmText === capturedSelectionText);
          let inserted = null;
          if (bookmarkMatchesProposal) {
            inserted = bmRange.insertText(lastProposal.text, "Replace");
          } else if (capturedSelectionText && capturedSelectionText.length <= MAX_SEARCH_LEN) {
            const ranges = context.document.body.search(capturedSelectionText, { matchCase: false });
            ranges.load("items");
            await context.sync();
            if (ranges.items.length > 0) {
              inserted = ranges.items[0].insertText(lastProposal.text, "Replace");
            } else {
              throw new Error("Không tìm thấy đoạn văn bản đã chọn để thay thế. Hãy chọn lại đoạn đó rồi Apply.");
            }
          } else if (capturedSelectionText) {
            // Too long to search for and the bookmark no longer matches it.
            throw new Error("Đoạn văn bản gốc không còn được ghim. Hãy chọn lại đoạn cần thay rồi hỏi lại.");
          } else {
            throw new Error("Không có văn bản được chọn để áp dụng.");
          }
          if (markRed) inserted.font.color = "#FF0000";
          // Re-pin the inserted text: replacing a bookmarked range removes the
          // bookmark, so without this a follow-up "rewrite it again" in the
          // same chat would have nothing pinned and fall back to stale text.
          if (canUseBookmark) inserted.insertBookmark(PIN_BOOKMARK_NAME);
          await context.sync();
        });
        pinnedSelectionText = lastProposal.text.trim();
        capturedSelectionText = pinnedSelectionText;
      }
      const n = lastProposal.type === "table" ? lastProposal.changes.length
        : editStats ? editStats.applied : 1;
      const skippedNote = editStats && editStats.skipped > 0
        ? ` (${editStats.skipped} bỏ qua — không tìm thấy hoặc quá dài để tìm kiếm)`
        : "";
      addMsg("bot", `Applied ${n} action(s).${skippedNote}`);
      lastProposal = null;
      preview.innerHTML = "";
      applyBtn.style.display = "none";
      markRedWrap.style.display = "none";
      setStatus("Ready.");
    } catch (err) {
      setStatus("Apply failed: " + (err.message || err));
      addMsg("bot", "⚠ Apply failed: " + (err.message || err));
    }
  }

  function newChat() {
    messages = [];
    lastProposal = null;
    capturedSelectionText = "";
    // Cleared so a fresh chat never reuses a stale pinned selection; the
    // next selectionChanged event (or an existing live selection) repins it.
    pinnedSelectionText = "";
    // Drop the pin bookmark from the document so a new chat starts clean.
    clearSelectionBookmark();
    log.innerHTML = "";
    preview.innerHTML = "";
    applyBtn.style.display = "none";
    markRedWrap.style.display = "none";
    // Reset to checked so a new chat always starts from the predictable
    // default rather than carrying over whatever the user last toggled.
    const markRedEl = document.getElementById("markRed");
    if (markRedEl) markRedEl.checked = true;
    setStatus("New chat. Select text and ask me to edit it.");
  }

  registerSelectionChangedHandler();
  onSelectionChangedHandler(); // prime pinnedSelectionText for a selection made before Office.onReady resolved

  askBtn.addEventListener("click", sendMessage);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
  input.addEventListener("input", () => {
    input.style.height = "auto";
    input.style.height = Math.min(input.scrollHeight, 120) + "px";
  });
  applyBtn.addEventListener("click", applyEdit);
  newChatBtn.addEventListener("click", newChat);
});
