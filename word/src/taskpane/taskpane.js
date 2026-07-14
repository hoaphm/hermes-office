/* global Office, Word */

import { askHermes } from "../shared/hermes.js";

Office.onReady().then(() => {
  const log = document.getElementById("log");
  const input = document.getElementById("prompt");
  const askBtn = document.getElementById("ask");
  const applyBtn = document.getElementById("apply");
  const newChatBtn = document.getElementById("newchat");
  const statusEl = document.getElementById("status");
  const preview = document.getElementById("preview");

  let messages = [];
  let lastProposal = null;
  // Snapshot of the text range captured at send time, used as a fallback
  // target if the user's selection is lost by the time they click Apply.
  let capturedSelectionText = "";

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

  async function getSelectionData() {
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

    // Plain text selection (even short) → operate on exactly that text.
    if (selText.length > 0) {
      capturedSelectionText = selText;
      return { type: "text", text: selText };
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

  // Base-26 column index (0-based) <-> letters, e.g. 0 -> "A", 25 -> "Z",
  // 26 -> "AA". Plain `String.fromCharCode(65 + i)` only covers single
  // letters and silently wraps/collides past column Z.
  function columnIndexToLetters(index) {
    let n = index + 1;
    let letters = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      letters = String.fromCharCode(65 + rem) + letters;
      n = Math.floor((n - 1) / 26);
    }
    return letters;
  }

  function columnLettersToIndex(letters) {
    let n = 0;
    for (let i = 0; i < letters.length; i++) {
      n = n * 26 + (letters.charCodeAt(i) - 64);
    }
    return n - 1;
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
    preview.innerHTML = "";

    try {
      const selectionData = await getSelectionData();

      if (selectionData.type === "empty") {
        addMsg("bot", "Tài liệu trống hoặc không đọc được. Hãy chọn một đoạn văn bản, hoặc gõ nội dung cần xử lý.");
        setStatus("Không có văn bản (doc trống / chưa sync xong). Thử lại.");
        return;
      }

      const statusText = selectionData.type === "fulldoc"
        ? `${selectionData.text.length} chars from full document`
        : (selectionData.text ? `${selectionData.text.length} chars selected` : "Table selected");
      setStatus(statusText + " — Hermes is thinking…");

      let docText = selectionData.text || "";
      if (selectionData.type === "fulldoc" && docText.length > 30000) {
        docText = docText.slice(0, 30000) + `\n\n[... còn ${docText.length - 30000} ký tự nữa ...]`;
      }
      const displayData = { ...selectionData, text: docText };

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
        }
      } else if (selectionData.type === "text") {
        lastProposal = { type: "text", text: reply };
        preview.innerHTML = `
          <div class="act">Proposed edit:</div>
          <div class="msg bot">${escapeHtml(reply).replace(/\n/g, "<br>")}</div>
        `;
        applyBtn.style.display = "block";
      } else if (selectionData.type === "fulldoc") {
        // Prefer inline edits (fix spelling etc.) — only the wrong spots change.
        const edits = parseEdits(reply);
        if (edits.length > 0) {
          lastProposal = { type: "fulldoc-edits", edits };
          preview.innerHTML = edits.map(e =>
            `<div class="act">"${escapeHtml(e.find)}" → "${escapeHtml(e.replace)}"</div>`
          ).join("") + `<div class="act">Áp dụng cho mọi vị trí (giữ nguyên định dạng).</div>`;
          applyBtn.style.display = "block";
        } else if (reply.trim().length > selectionData.text.length * 0.5) {
          // Looks like a full rewrite / translation → replace whole doc.
          lastProposal = { type: "fulldoc-full", text: reply };
          preview.innerHTML = `<div class="act">Sẽ THAY THẾ TOÀN BỘ văn bản bằng kết quả trên.</div>`;
          applyBtn.style.display = "block";
        } else {
          // Plain answer / review — nothing to apply.
          lastProposal = null;
          preview.innerHTML = `<div class="act">Đây là câu trả lời / nhận xét — không áp dụng trực tiếp.</div>`;
          applyBtn.style.display = "none";
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

  function parseTableChanges(reply) {
    const fenced = reply.match(/```json\s*([\s\S]*?)```/i);
    const jsonStr = fenced ? fenced[1] : reply.match(/\{"cells"\s*:\s*\[[\s\S]*?\]\s*\}/);
    if (!jsonStr) return [];
    try {
      const obj = JSON.parse(fenced ? fenced[1] : jsonStr[0]);
      return obj.cells || [];
    } catch (_) {
      return [];
    }
  }

  function parseEdits(reply) {
    // Accept either a fenced ```json block OR a bare { "edits": [...] } object
    // that may be embedded in prose. Try fenced first, then a loose match.
    const candidates = [];
    const fenced = reply.match(/```(?:json)?\s*([\s\S]*?)```/i);
    if (fenced) candidates.push(fenced[1]);
    const bare = reply.match(/\{[^{}]*"edits"\s*:\s*\[[\s\S]*?\]\s*\}/);
    if (bare) candidates.push(bare[0]);

    for (const c of candidates) {
      try {
        const obj = JSON.parse(c.trim());
        const edits = obj && obj.edits;
        if (Array.isArray(edits)) {
          return edits
            .filter((e) => e && typeof e.find === "string" && e.find.length > 0)
            .map((e) => ({ find: e.find, replace: e.replace === undefined ? "" : String(e.replace) }));
        }
      } catch (_) { /* try next candidate */ }
    }
    return [];
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
            cell.body.getRange().insertText(String(change.value), "Replace");
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
              ranges.items.forEach((r) => r.insertText(String(edit.replace), "Replace"));
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
          range.insertText(lastProposal.text, "Replace");
          await context.sync();
        });
      } else {
        // Plain text: replace the selection if it is still intact, otherwise
        // fall back to replacing the first occurrence of the captured text so
        // Apply never fails just because focus moved to the task pane.
        await Word.run(async (context) => {
          const sel = context.document.getSelection();
          sel.load("text");
          await context.sync();
          const stillSelected = (sel.text || "").trim();

          if (stillSelected.length > 0) {
            sel.insertText(lastProposal.text, "Replace");
          } else if (capturedSelectionText) {
            const ranges = context.document.body.search(capturedSelectionText, { matchCase: false });
            ranges.load("items");
            await context.sync();
            if (ranges.items.length > 0) {
              ranges.items[0].insertText(lastProposal.text, "Replace");
            } else {
              throw new Error("Không tìm thấy đoạn văn bản đã chọn để thay thế.");
            }
          } else {
            throw new Error("Không có văn bản được chọn để áp dụng.");
          }
          await context.sync();
        });
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
    log.innerHTML = "";
    preview.innerHTML = "";
    applyBtn.style.display = "none";
    setStatus("New chat. Select text and ask me to edit it.");
  }

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
