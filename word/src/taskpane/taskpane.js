/* global Office, Word, document */
// Word task pane — thin UI + orchestration layer.
//
// All selection/bookmark state lives in selection-mgr.js (two orthogonal state
// machines). All Apply logic lives in proposal-mgr.js (immutable target). This
// file wires DOM events to those modules and renders results. No Word.run here.

import { askHermes } from "../shared/hermes.js";
import {
  columnIndexToLetters,
  parseWordEdits,
  parseWordTableChanges,
  appendMessage,
  appendTypingIndicator,
  removeTypingIndicator,
  setStatus as setStatusUi,
  setBusy as setBusyUi,
  showToast,
  mountContextBar,
  renderProposalCard,
} from "../../../shared/proposal-card.js";
import { createSelectionMgr, MAX_FULLDOC_CHARS } from "./selection-mgr.js";
import { createProposalMgr, cellRefToPosition } from "./proposal-mgr.js";

const MAX_HISTORY_MESSAGES = 20;

Office.onReady().then(() => {
  const log = document.getElementById("log");
  const input = document.getElementById("prompt");
  const askBtn = document.getElementById("ask");
  const applyBtn = document.getElementById("apply");
  const markRedWrap = document.getElementById("markRedWrap");
  const newChatBtn = document.getElementById("newchat");
  const statusEl = document.getElementById("status");
  const statusRowEl = document.getElementById("statusRow");
  const preview = document.getElementById("preview");
  const emptyEl = document.getElementById("empty");
  let typingEl = null;
  const contextHost =
    document.getElementById("contextRow") ||
    document.querySelector(".ds-header") ||
    document.querySelector("header") ||
    document.body;

  // ---- state (UI-level only; selection state is in selection-mgr) ----------
  let messages = [];
  let lastProposal = null; // { type, ... } — what Apply will act on
  let busy = false; // UI re-entrancy guard (distinct from selection-mgr gate)

  // ---- selection module (owns Word.run, bookmark, gate) --------------------
  const sel = createSelectionMgr({ runWord: (fn) => Word.run(fn) });

  // ---- context bar (pure UI) -----------------------------------------------
  function refreshContextBar(extra) {
    const chips = [];
    const pinned = sel.getPinnedText();
    if (pinned) {
      chips.push({
        label: "Pinned",
        value: pinned.slice(0, 60),
        state: "pinned",
      });
    }
    if (extra && extra.snapshot)
      chips.push({ label: "Snapshot", value: extra.snapshot });
    if (extra && extra.willTruncate)
      chips.push({
        label: "⚠ Cắt bớt",
        value: extra.willTruncate,
        state: "warn",
      });
    if (!chips.length)
      chips.push({ label: "Sẵn sàng", value: "", state: "idle" });
    mountContextBar(contextHost, chips);
  }
  refreshContextBar();

  // ---- UI helpers ----------------------------------------------------------
  function addMsg(role, text, opts) {
    if (emptyEl) emptyEl.classList.add("is-hidden");
    return appendMessage(log, role, text, opts);
  }

  function setStatus(text, tone) {
    setStatusUi(statusEl, text, { tone });
    if (statusRowEl) {
      if (tone) statusRowEl.dataset.tone = tone;
      else delete statusRowEl.dataset.tone;
    }
  }

  // ---- system prompt builder (pure) ----------------------------------------
  function buildSystemPrompt(data) {
    if (!data || data.type === "empty") {
      return "You are editing a Word document. The user has not selected any text. Generate the content they request. Respond with the text only — no markdown fences, no explanations.";
    }
    if (data.type === "table") {
      let tableDesc = "";
      if (data.tables && data.tables.length > 0) {
        tableDesc = data.tables
          .map((t, i) => {
            const flat = t.values
              .map((row, ri) => {
                const cells = row
                  .map(
                    (cell, ci) =>
                      `  [${columnIndexToLetters(ci)}${ri + 1}] ${cell}`,
                  )
                  .join("\n");
                return `Row ${ri + 1}:\n${cells}`;
              })
              .join("\n\n");
            return `TABLE ${i + 1} (${t.rowCount} rows x ${t.columnCount} cols):\n${flat}`;
          })
          .join("\n\n");
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
    if (data.type === "fulldoc") {
      return `You are checking and editing the ENTIRE open Word document. The user selected no text.

For spelling, grammar, or wording corrections, return ONLY one JSON object. No markdown fence. No explanation. Exact format:
{"edits":[{"find":"the exact incorrect text from the document","replace":"the corrected text"}]}

Rules:
- Include every correction.
- Each find value must be copied exactly from the document and be no longer than 255 characters.
- Replace only incorrect fragments. Never return the whole document.
- If no correction is needed, return {"edits":[]}.

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

  // A reply that is nothing but one fenced block — unwrap it.
  function stripWrappingFence(text) {
    const t = String(text).trim();
    const m = t.match(/^```[a-zA-Z]*[ \t]*\r?\n([\s\S]*?)\r?\n?```$/);
    return m ? m[1].trim() : t;
  }

  // ---- send ----------------------------------------------------------------
  async function sendMessage() {
    const userText = input.value.trim();
    if (!userText || busy) return;

    busy = true;
    addMsg("user", userText);
    input.value = "";
    setBusyUi(askBtn, true);
    setStatus("Đang đọc tài liệu…", "busy");
    lastProposal = null;
    // The previous Proposal is gone, so its hold on the Pin goes with it.
    sel.unlockPin();
    applyBtn.hidden = true;
    markRedWrap.hidden = true;
    preview.innerHTML = "";
    refreshContextBar({ snapshot: "Đang đọc tài liệu…" });

    const handle = sel.beginAsk();
    try {
      let selectionData;
      try {
        selectionData = await sel.getSelectionData();
      } catch (e) {
        throw new Error(`[stage:read-doc] ${e.message || e}`);
      }

      if (selectionData.type === "text") {
        await sel.ensurePinnedBookmark(selectionData.capturedText);
      }

      if (selectionData.type === "multi-table") {
        addMsg(
          "bot",
          "Đang chọn nhiều bảng. Hãy chọn một bảng duy nhất rồi hỏi lại Hermes.",
          { tone: "warn" },
        );
        setStatus("Cần chọn một bảng.", "warn");
        return;
      }
      if (selectionData.type === "empty") {
        addMsg(
          "bot",
          "Tài liệu trống hoặc không đọc được. Hãy chọn một đoạn văn bản, hoặc gõ nội dung cần xử lý.",
        );
        setStatus("Không có văn bản (doc trống / chưa sync xong). Thử lại.");
        return;
      }

      // Bound full-document prompts.
      const displayData = { ...selectionData };
      let fullDocTruncated = false;
      if (
        displayData.type === "fulldoc" &&
        displayData.text.length > MAX_FULLDOC_CHARS
      ) {
        displayData.text = displayData.text.slice(0, MAX_FULLDOC_CHARS);
        fullDocTruncated = true;
      }

      const statusText =
        selectionData.type === "fulldoc"
          ? `${selectionData.text.length} ký tự từ toàn văn bản`
          : selectionData.text
            ? `${selectionData.text.length} ký tự được chọn`
            : "Đã chọn bảng";
      setStatus(
        statusText +
          (fullDocTruncated
            ? ` — đã giới hạn ${MAX_FULLDOC_CHARS} ký tự đầu`
            : ""),
      );
      if (fullDocTruncated) {
        refreshContextBar({
          snapshot: `Phân tích ${MAX_FULLDOC_CHARS}/${selectionData.text.length} ký tự đầu`,
          willTruncate: "Tài liệu dài",
        });
      }

      const sysPrompt = buildSystemPrompt(displayData);
      const payload = [
        { role: "system", content: sysPrompt },
        ...messages,
        { role: "user", content: userText },
      ];

      setStatus("Hermes đang suy nghĩ…", "busy");
      typingEl = appendTypingIndicator(log);
      let reply;
      try {
        reply = await askHermes(payload);
      } catch (e) {
        throw new Error(`[stage:call-api] ${e.message || e}`);
      }
      removeTypingIndicator(typingEl);
      typingEl = null;
      addMsg("bot", reply);

      // Capture immutable target for Apply BEFORE rendering.
      const target = sel.captureTarget(selectionData);

      if (selectionData.type === "table") {
        const tableChanges = parseWordTableChanges(reply);
        if (tableChanges.length > 0) {
          const table = selectionData.tables[0];
          lastProposal = {
            type: "table",
            target,
            changes: tableChanges.map((change) => {
              const pos = cellRefToPosition(
                change.cell,
                table.rowCount,
                table.columnCount,
              );
              return {
                ...change,
                old: pos ? table.values[pos.row][pos.col] : undefined,
              };
            }),
          };
          renderProposalCard(preview, {
            title: "Cập nhật bảng đã chọn",
            actions: lastProposal.changes.map((c) => ({
              type: "setCell",
              cell: c.cell,
              old: c.old,
              new: c.value,
            })),
          });
          applyBtn.hidden = false;
          markRedWrap.hidden = false;
        }
      } else if (selectionData.type === "text") {
        const passage = stripWrappingFence(reply);
        lastProposal = { type: "text", target, text: passage };
        renderProposalCard(preview, {
          title: "Đề xuất chỉnh sửa đoạn đã chọn",
          actions: [{ type: "replace", find: target.text, replace: passage }],
        });
        // Apply writes through the bookmark, so the Pin has to outlive the pane
        // taking focus. Without this the next empty selectionChanged clears it
        // and Apply falls back to a body search that refuses a repeated passage.
        sel.lockPin();
        applyBtn.hidden = false;
        markRedWrap.hidden = false;
      } else if (selectionData.type === "fulldoc") {
        const edits = parseWordEdits(reply);
        if (edits.length > 0) {
          lastProposal = { type: "fulldoc-edits", target, edits };
          // Count matches BEFORE the card is drawn: Apply replaces every
          // occurrence, so "how many" is part of what the user is approving.
          const counts = await sel.countOccurrences(edits.map((e) => e.find));
          renderProposalCard(preview, {
            title: "Sửa nhanh toàn văn bản",
            actions: edits.map((e) => ({
              type: "replace",
              find: e.find,
              replace: e.replace,
              matchCount: counts.get(e.find),
            })),
          });
          applyBtn.hidden = false;
          markRedWrap.hidden = false;
        } else {
          lastProposal = null;
          preview.innerHTML = `<div class="ds-card-action"><div class="label">Đây là câu trả lời / nhận xét — không áp dụng trực tiếp.</div></div>`;
          applyBtn.hidden = true;
          markRedWrap.hidden = true;
        }
      }

      messages.push({ role: "user", content: userText });
      messages.push({ role: "assistant", content: reply });
      if (messages.length > MAX_HISTORY_MESSAGES) {
        messages = messages.slice(-MAX_HISTORY_MESSAGES);
      }
      setStatus("Sẵn sàng.");
    } catch (err) {
      removeTypingIndicator(typingEl);
      typingEl = null;
      addMsg("bot", (err.message || String(err)) + "\n" + (err.stack || ""), {
        tone: "err",
      });
      setStatus("Lỗi.", "err");
    } finally {
      handle.end();
      busy = false;
      setBusyUi(askBtn, false);
      input.focus();
    }
  }

  // ---- apply ---------------------------------------------------------------
  async function applyEdit() {
    if (!lastProposal || busy) return;

    const markRedEl = document.getElementById("markRed");
    const markRed = markRedEl && markRedEl.checked;

    busy = true;
    setBusyUi(applyBtn, true);
    setStatus("Đang áp dụng…", "busy");

    const handle = sel.beginApply();
    try {
      const mgr = createProposalMgr({
        target: lastProposal.target,
        runWord: (fn) => Word.run(fn),
        showToast,
      });
      let stats;
      if (lastProposal.type === "table") {
        stats = await mgr.applyTable(lastProposal.changes, { markRed });
      } else if (lastProposal.type === "fulldoc-edits") {
        stats = await mgr.applyFulldocEdits(lastProposal.edits, { markRed });
      } else {
        stats = await mgr.applyText(lastProposal.text, { markRed });
        // Text proposal complete — its bookmark was re-pinned on the new text
        // by proposal-mgr, but the in-memory selection state is now stale.
        await sel.clearPin();
      }

      const n =
        lastProposal.type === "table"
          ? lastProposal.changes.length
          : stats.applied;
      const skippedNote =
        stats.skipped > 0
          ? ` (${stats.skipped} bỏ qua — không tìm thấy hoặc quá dài để tìm kiếm)`
          : "";
      addMsg("bot", `Đã áp dụng ${n} thay đổi.${skippedNote}`, { tone: "ok" });
      showToast(`Đã áp dụng ${n} thay đổi.${skippedNote}`, { tone: "ok" });
      lastProposal = null;
      // No Proposal left to anchor: the text path already cleared the Pin, the
      // others just release the hold. A failed Apply deliberately keeps both,
      // because the Proposal is still on screen and still retryable.
      sel.unlockPin();
      preview.innerHTML = "";
      applyBtn.hidden = true;
      markRedWrap.hidden = true;
      setStatus("Sẵn sàng.");
      refreshContextBar();
    } catch (err) {
      const errText = "⚠ Áp dụng thất bại: " + (err.message || err);
      setStatus(errText, "err");
      addMsg("bot", errText, { tone: "err" });
      showToast(errText, { tone: "err", timeout: 6000 });
    } finally {
      handle.end();
      busy = false;
      setBusyUi(applyBtn, false);
    }
  }

  // ---- new chat ------------------------------------------------------------
  async function newChat() {
    if (busy) return;
    messages = [];
    lastProposal = null;
    log.innerHTML = "";
    if (emptyEl) emptyEl.classList.remove("is-hidden");
    preview.innerHTML = "";
    applyBtn.hidden = true;
    markRedWrap.hidden = true;
    const markRedEl = document.getElementById("markRed");
    if (markRedEl) markRedEl.checked = true;
    setStatus("Cuộc trò chuyện mới. Chọn văn bản và yêu cầu Hermes chỉnh sửa.");
    // UI first so the button feels instant; the document-side clear is a
    // Word.run round-trip. Awaited so an unawaited bookmark delete cannot land
    // after the next selectionChanged has already re-pinned.
    await sel.reset();
    refreshContextBar();
  }

  // ---- wiring --------------------------------------------------------------
  sel.registerSelectionChanged(() => refreshContextBar());
  sel.onSelectionChanged(); // prime for a selection made before Office.onReady

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
