/* global Office, Excel, document, window */
import { askHermes } from "../shared/hermes";
// Pure helpers deduped with the Word taskpane via the repo-root shared/
// folder (no npm workspace between the two add-ins) — see shared/parsers.js.
import {
  signature,
  contentSignature,
  resolveRange,
  chartType,
  trimHistory,
} from "../../../shared/parsers.js";
// Reply parsing, action descriptions and cell-value literalisation live in a
// pure module so they can be tested without Office.js — see actions.test.js.
import {
  MAX_ACTIONS,
  literalCellValue,
  literalizeGrid,
  splitReply,
  describe,
  tableName,
  retargetsSheets,
} from "./actions.js";
// UI helpers (proposal card, toast, context bar) live in shared/proposal-card.js
// so both add-ins get the same design system + a11y treatment.
import {
  appendMessage,
  appendTypingIndicator,
  removeTypingIndicator,
  setStatus as setStatusUi,
  setBusy as setBusyUi,
  showToast,
  renderProposalCard,
  writesBlind,
  partialApplyNotice,
} from "../../../shared/proposal-card.js";

const MAX_ROWS = 500;
// Very wide sheets can otherwise blow up the snapshot payload sent to Hermes.
const MAX_COLS = 100;
// Belt-and-suspenders byte cap on top of the row/col caps, in case a sheet
// is dense (long strings) rather than just tall/wide.
const MAX_SNAPSHOT_BYTES = 200000;
// Cap on replayed conversation turns (excluding the system message at index
// 0). Each turn can carry a full sheet snapshot, so an unbounded history grew
// the request payload without limit across a long session.
const MAX_HISTORY_MESSAGES = 20;

const SYSTEM = `You are Hermes, embedded in an Excel task pane. Chat naturally and concisely.

When a message starts with a bracketed [Active sheet ...] note, that is the current sheet's data — use it. If a [Selected range ...] note is present, the user has highlighted specific cell(s); operate on that selection by default unless they name a different range. The data is only re-sent when it changes, so for follow-up questions rely on the data already shown earlier in the conversation.

When the user wants to modify the workbook, append EXACTLY ONE fenced block at the very end of your reply:
\`\`\`json
{"actions":[ ... ]}
\`\`\`
Supported action types:
- {"type":"setCell","cell":"B2","old":"<current>","new":"<new value>"}
- {"type":"setCells","range":"A1:B3","values":[["x","y"],["x","y"]]}
- {"type":"format","range":"C2:C50","numberFormat":"$#,##0.00","bold":false,"fill":"#FFF2CC"}
- {"type":"createTable","range":"A1:F200","name":"Leads","hasHeaders":true}
- {"type":"createChart","chartType":"ColumnClustered","dataRange":"A1:B20","title":"By status"}   // chartType: ColumnClustered | Bar | Line | Pie | XYScatter
- {"type":"newSheet","name":"Dashboard"}
- {"type":"renameSheet","to":"Acquisitions 2021"}
Always also give a short natural-language reply. Use absolute A1 refs; qualify cross-sheet refs as Sheet!A1. Omit the block entirely when no change is requested.`;

const history = [{ role: "system", content: SYSTEM }];
let pendingActions = [];
let lastSig = null;
let busy = false;
// Sheet the current pendingActions proposal was generated against — apply()
// must target this sheet, not whatever happens to be active by click time.
let proposalSheetName = null;
// Content fingerprint of that sheet at the moment the Proposal was made. Only
// setCell carries a per-Action old-value guard; setCells, format and the rest
// write blind, so this is what stops a Proposal reviewed against one sheet
// from being applied to a sheet that has since changed.
let proposalSig = null;
// True once a press has written part of this Proposal. The Remainder on the
// card was reviewed against the sheet as it was BEFORE those writes, so the
// card has to say so rather than look identical to a fresh Proposal. ADR-0004.
let afterPartialApply = false;

Office.onReady(() => {
  const askBtn = document.getElementById("ask");
  const newChatBtn = document.getElementById("newchat");
  const promptEl = document.getElementById("prompt");
  const applyBtn = document.getElementById("apply");
  const emptyEl = document.getElementById("empty");
  // Typing indicator shown while Hermes is thinking — created on send,
  // removed when the reply lands (or on error).
  let typingEl = null;

  askBtn.addEventListener("click", ask);
  newChatBtn.addEventListener("click", newChat);
  applyBtn.addEventListener("click", apply);
  promptEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      ask();
    }
  });
  promptEl.addEventListener("input", () => {
    promptEl.style.height = "auto";
    promptEl.style.height = Math.min(promptEl.scrollHeight, 120) + "px";
  });

  // One-line compact meta row: Sheet · Range · Selection. Each item is
  // hidden if empty so the separator doesn't dangle. Cheap to re-render
  // on every state change; no DOM thrash because we update the same nodes.
  const sheetEl = document.querySelector('.ds-meta-item[data-key="sheet"]');
  const rangeEl = document.querySelector('.ds-meta-item[data-key="range"]');
  const selEl = document.querySelector('.ds-meta-item[data-key="selection"]');
  const separators = [...document.querySelectorAll(".ds-meta-sep")];

  function setMetaItem(el, value, { active = false } = {}) {
    if (!el) return;
    const trimmed = (value || "").trim();
    el.textContent = trimmed || "—";
    el.dataset.empty = trimmed ? "false" : "true";
    if (active && trimmed) {
      el.dataset.state = "active";
    } else {
      delete el.dataset.state;
    }
    // Hide the element when nothing meaningful to show, plus its trailing
    // separator so the row stays balanced.
    el.style.display = trimmed ? "" : "none";
  }

  function refreshMeta(snap) {
    if (!snap) {
      setMetaItem(sheetEl, "");
      setMetaItem(rangeEl, "");
      setMetaItem(selEl, "");
    } else {
      setMetaItem(sheetEl, snap.name);
      setMetaItem(rangeEl, snap.address || "");
      setMetaItem(selEl, snap.selection ? snap.selection.address : "", {
        active: true,
      });
    }
    const visible = [sheetEl, rangeEl, selEl].filter((el) => el && el.dataset.empty === "false");
    separators.forEach((sep, i) => {
      sep.style.display = i < visible.length - 1 ? "" : "none";
    });
  }
  refreshMeta(null);
  window.__hermesRefreshContext = refreshMeta;

  // First message of a chat replaces the empty state; newChat() brings
  // it back. Exposed so ask()/newChat() (defined below) can toggle it.
  window.__hermesSetEmptyHidden = (hidden) => {
    if (emptyEl) emptyEl.classList.toggle("is-hidden", hidden);
  };
  window.__hermesTyping = (show) => {
    if (show) {
      typingEl = appendTypingIndicator(document.getElementById("log"));
    } else {
      removeTypingIndicator(typingEl);
      typingEl = null;
    }
  };
});

// ---- conversation ----------------------------------------------------------

async function ask() {
  const input = document.getElementById("prompt");
  const prompt = input.value.trim();
  if (!prompt || busy) return;
  input.value = "";
  addBubble("user", prompt);
  setBusy(true, "Đang đọc bảng tính…");
  try {
    const snap = await getSnapshot();
    if (typeof window.__hermesRefreshContext === "function") window.__hermesRefreshContext(snap);
    const sig = signature(snap);
    const dataChanged = sig !== lastSig;
    const content = dataChanged ? `${dataNote(snap)}\n\n${prompt}` : prompt;
    // Built, not pushed: a turn joins the conversation only once its reply is
    // in hand. Pushing first left an orphan user turn behind every failed call,
    // so a second attempt sent two user messages in a row.
    const outgoing = [...history, { role: "user", content }];

    setBusy(true, "Hermes đang suy nghĩ…");
    if (typeof window.__hermesTyping === "function") window.__hermesTyping(true);
    const raw = await askHermes(outgoing);
    if (typeof window.__hermesTyping === "function") window.__hermesTyping(false);

    history.push({ role: "user", content }, { role: "assistant", content: raw });
    // Likewise deferred: advancing lastSig before the call meant a failed turn
    // recorded a Snapshot as delivered that the Provider never saw, and the
    // next turn withheld it as unchanged.
    if (dataChanged) lastSig = sig;
    // Trim oldest turns, always preserving the system message at index 0.
    if (trimHistory(history, MAX_HISTORY_MESSAGES, { keepSystem: true }).trimmed) {
      // The dropped turns may have carried the only copy of the sheet
      // snapshot, so force the next turn to re-send fresh data.
      lastSig = null;
    }

    const { prose, actions } = splitReply(raw);
    addBubble("bot", prose);
    if (actions.length > MAX_ACTIONS) {
      pendingActions = [];
      addBubble(
        "bot",
        "Đề xuất có quá nhiều thay đổi. Hãy yêu cầu Hermes chia thành vài bước nhỏ hơn.",
        "warn"
      );
      setStatus("Đề xuất quá lớn.", "warn");
      renderActions([]);
      return;
    }
    // Baseline the staleness guard here, not at the top of ask(): it must
    // describe the Snapshot THIS Proposal was generated from. Setting it before
    // the call meant a failed ask re-baselined the Proposal still sitting on the
    // card against a newer sheet, and Apply then measured that Proposal's
    // freshness against a state its author never saw.
    proposalSheetName = snap.name;
    proposalSig = contentSignature(snap);
    pendingActions = actions;
    // A fresh Proposal, reviewed against a sheet nothing has written to since.
    afterPartialApply = false;
    renderActions(actions);
    setStatus(
      actions.length
        ? `${actions.length} hành động được đề xuất — xem lại rồi Áp dụng.`
        : "Sẵn sàng."
    );
  } catch (e) {
    if (typeof window.__hermesTyping === "function") window.__hermesTyping(false);
    addBubble("bot", "⚠ " + e.message, "err");
    setStatus("Lỗi.", "err");
  } finally {
    setBusy(false);
  }
}

function newChat() {
  history.length = 1; // keep system message
  lastSig = null;
  proposalSheetName = null;
  proposalSig = null;
  if (typeof window.__hermesTyping === "function") window.__hermesTyping(false);
  clearPending();
  document.getElementById("log").innerHTML = "";
  if (typeof window.__hermesSetEmptyHidden === "function") window.__hermesSetEmptyHidden(false);
  setStatus("Cuộc trò chuyện mới. Mở sheet dữ liệu và đặt câu hỏi.");
  if (typeof window.__hermesRefreshContext === "function") window.__hermesRefreshContext(null);
}

// ---- reading the sheet (only sent when changed) ----------------------------

// Reads the active sheet, or a named one when the caller has a specific sheet
// in mind — re-baselining a Remainder must read the Proposal's OWN sheet, and
// the active sheet is not it once a newSheet Action has activated the sheet it
// created. Returns null when the named sheet is gone.
async function getSnapshot(sheetName) {
  return Excel.run(async (context) => {
    const named = Boolean(sheetName);
    const sheet = named
      ? context.workbook.worksheets.getItemOrNullObject(sheetName)
      : context.workbook.worksheets.getActiveWorksheet();
    // isNullObject only exists on the OrNullObject result; loading it off
    // getActiveWorksheet() would throw.
    sheet.load(named ? ["name", "isNullObject"] : ["name"]);
    const used = sheet.getUsedRangeOrNullObject(true); // valuesOnly
    used.load(["address", "values"]);
    // The sheet name and used range are the parts that must not fail, so they
    // get a sync of their own. They used to share one with the best-effort
    // selection read below — and a sync that rejects does not re-queue the
    // loads it carried, so a selection Excel refused to hand over (a multi-area
    // pick, a chart focused) left sheet.name permanently unloaded. Reading it
    // afterwards threw PropertyNotLoaded and the whole Ask died on a detail
    // the code had already decided was optional.
    await context.sync();
    if (named && sheet.isNullObject) return null;

    // Read the user's current selection (single cell or range), then cap it
    // by the same row/column/byte limits as the used-range snapshot — a user
    // who selects A1:XFD1048576 would otherwise freeze the task pane.
    let selection = null;
    try {
      const sel = context.workbook.getSelectedRange();
      sel.load(["address", "values"]);
      await context.sync();
      if (sel.address) {
        let sv = sel.values;
        if (sv && sv.length > MAX_ROWS) {
          sv = sv.slice(0, MAX_ROWS);
        }
        if (sv && sv.length > 0 && sv[0].length > MAX_COLS) {
          sv = sv.map((r) => r.slice(0, MAX_COLS));
        }
        if (sv && JSON.stringify(sv).length > MAX_SNAPSHOT_BYTES) {
          sv = [];
        }
        selection = { address: sel.address, values: sv || [] };
      }
    } catch {
      /* no selection or multi-area — ignore */
    }

    if (used.isNullObject) {
      return {
        name: sheet.name,
        address: null,
        values: [],
        rowsTruncated: false,
        colsTruncated: false,
        bytesTruncated: false,
        selection,
      };
    }
    let values = used.values;
    let rowsTruncated = false;
    let colsTruncated = false;
    let bytesTruncated = false;
    if (values.length > MAX_ROWS) {
      values = values.slice(0, MAX_ROWS);
      rowsTruncated = true;
    }
    if (values.length > 0 && values[0].length > MAX_COLS) {
      values = values.map((row) => row.slice(0, MAX_COLS));
      colsTruncated = true;
    }
    // Belt-and-suspenders: even within the row/col caps a sheet of long
    // strings can blow the payload up. Halve rows until under the byte cap.
    let payloadBytes = JSON.stringify(values).length;
    while (payloadBytes > MAX_SNAPSHOT_BYTES && values.length > 1) {
      values = values.slice(0, Math.ceil(values.length / 2));
      payloadBytes = JSON.stringify(values).length;
      bytesTruncated = true;
    }
    return {
      name: sheet.name,
      address: used.address,
      values,
      rowsTruncated,
      colsTruncated,
      bytesTruncated,
      selection,
    };
  });
}

function dataNote(s) {
  let head;
  if (!s.address) {
    head = `[Active sheet "${s.name}" is empty.]`;
  } else {
    const notes = [];
    if (s.rowsTruncated) notes.push(`first ${MAX_ROWS} rows`);
    if (s.colsTruncated) notes.push(`first ${MAX_COLS} cols`);
    if (s.bytesTruncated) notes.push("further truncated to fit size limit");
    const suffix = notes.length ? ` (${notes.join(", ")})` : "";
    head = `[Active sheet "${s.name}", range ${s.address}${suffix}. Current data:]\n${JSON.stringify(s.values)}`;
  }

  if (s.selection) {
    head += `\n\n[Selected range ${s.selection.address}. Current selection values:]\n${JSON.stringify(s.selection.values)}`;
  }
  return head;
}

// ---- preview + apply -------------------------------------------------------

function renderActions(actions) {
  const box = document.getElementById("preview");
  const card = renderProposalCard(box, {
    title: actions.length
      ? `${actions.length} hành động đề xuất cho sheet "${proposalSheetName || ""}"`
      : "",
    actions: actions,
    // Only after a press has already written part of this Proposal. A fresh
    // Proposal carries neither, so the marks mean something when they appear.
    notice: afterPartialApply ? partialApplyNotice(actions) : null,
    badgeAction: afterPartialApply ? (a) => (writesBlind(a) ? "⚠ ghi đè" : null) : null,
  });
  // renderProposalCard renders no CTA of its own — the single Apply button is
  // the footer-level #apply, which we show only when there is a card.
  document.getElementById("apply").hidden = !card;
  document.getElementById("highlightWrap").hidden = !card;
}

async function apply() {
  if (!pendingActions.length || busy) return;
  // Tinting applied cells used to be unconditional, permanently overwriting
  // whatever fill the user had there with no way to opt out. Now it mirrors
  // Word's "mark edits red" toggle: on by default, but the user's choice.
  const highlightEl = document.getElementById("highlight");
  const highlight = highlightEl && highlightEl.checked;
  setBusy(true, "Đang áp dụng…");
  let applied = 0;
  let skipped = 0;
  const failures = [];
  // Actions that did not run. Kept so a partial Apply can be retried instead of
  // being thrown away with the rest of the Proposal.
  const remaining = [];
  // Actions that DID run — what decides whether a Remainder is still meaningful.
  const ran = [];
  try {
    // The user reviewed this Proposal against the sheet as it looked at Ask
    // time. If the content changed since, the review no longer describes what
    // Apply would write — refuse rather than write blind. Compared on content
    // only: moving the selection between Ask and Apply is normal.
    if (proposalSig) {
      const fresh = await getSnapshot();
      // A wrong active sheet has its own, more precise error inside Excel.run.
      if (fresh.name === proposalSheetName && contentSignature(fresh) !== proposalSig) {
        throw new Error(
          `Sheet "${proposalSheetName}" đã thay đổi kể từ khi đề xuất được tạo, nên thẻ ` +
            `bạn vừa xem không còn mô tả đúng thứ sẽ được ghi. Hãy hỏi lại Hermes.`
        );
      }
    }
    await Excel.run(async (context) => {
      const wb = context.workbook;
      const activeSheet = wb.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      let sheet = proposalSheetName
        ? wb.worksheets.getItemOrNullObject(proposalSheetName)
        : activeSheet;
      sheet.load(["name", "isNullObject"]);
      await context.sync();

      if (sheet.isNullObject) {
        throw new Error(`Sheet "${proposalSheetName}" không còn tồn tại. Hãy hỏi lại Hermes.`);
      }
      if (proposalSheetName && activeSheet.name !== proposalSheetName) {
        throw new Error(
          `Sheet đang mở là "${activeSheet.name}" nhưng đề xuất này được tạo cho "${proposalSheetName}". ` +
            `Hãy quay lại sheet "${proposalSheetName}" rồi Apply.`
        );
      }

      // Pre-load row/column counts for every numberFormat action up front,
      // in ONE sync, so the main loop below doesn't need a mid-loop
      // context.sync() just to read dimensions for the common case. Skipped
      // for actions after a newSheet/renameSheet using an unqualified range,
      // since the target sheet isn't resolvable yet — those fall back to a
      // per-action load further down.
      const formatDims = new Map();
      let sheetMayChange = false;
      pendingActions.forEach((a, i) => {
        if (a.type === "newSheet" || a.type === "renameSheet") {
          sheetMayChange = true;
          return;
        }
        if (a.type !== "format" || !a.numberFormat) return;
        const isQualified = String(a.range || "").includes("!");
        if (!isQualified && sheetMayChange) return;
        try {
          const r = resolveRange(wb, sheet, a.range);
          r.load(["rowCount", "columnCount"]);
          formatDims.set(i, r);
        } catch {
          /* handled per-action below */
        }
      });
      if (formatDims.size > 0) await context.sync();

      for (let i = 0; i < pendingActions.length; i++) {
        const a = pendingActions[i];
        try {
          switch (a.type) {
            case "newSheet": {
              sheet = wb.worksheets.add(a.name || "Sheet");
              sheet.activate();
              break;
            }
            case "renameSheet": {
              sheet.name = a.to || a.name;
              break;
            }
            case "setCell": {
              const r = resolveRange(wb, sheet, a.cell);
              // Check the value hasn't changed since the proposal was made.
              if (a.old !== undefined) {
                r.load("values");
                await context.sync();
                const cur = String((r.values || [[null]])[0][0] ?? "");
                if (cur !== String(a.old)) {
                  skipped++;
                  remaining.push(a);
                  failures.push(
                    `${describe(a)}: bị bỏ qua — giá trị hiện tại ("${cur}") khác với giá trị gốc ("${a.old}") của đề xuất.`
                  );
                  continue;
                }
              }
              r.values = [[literalCellValue(a.new)]];
              if (highlight) r.format.fill.color = "#C6EFCE";
              break;
            }
            case "setCells": {
              resolveRange(wb, sheet, a.range).values = literalizeGrid(a.values);
              break;
            }
            case "format": {
              const r = resolveRange(wb, sheet, a.range);
              if (a.bold !== undefined) r.format.font.bold = !!a.bold;
              if (a.fill) r.format.fill.color = a.fill;
              if (a.numberFormat) {
                const dims = formatDims.get(i);
                let rowCount, columnCount;
                if (dims) {
                  rowCount = dims.rowCount;
                  columnCount = dims.columnCount;
                } else {
                  r.load(["rowCount", "columnCount"]);
                  await context.sync();
                  rowCount = r.rowCount;
                  columnCount = r.columnCount;
                }
                const fmt = Array.from({ length: rowCount }, () =>
                  Array.from({ length: columnCount }, () => a.numberFormat)
                );
                r.numberFormat = fmt;
              }
              break;
            }
            case "createTable": {
              const t = wb.tables.add(resolveRange(wb, sheet, a.range), a.hasHeaders !== false);
              if (a.name) t.name = tableName(a.name);
              break;
            }
            case "createChart": {
              const ch = sheet.charts.add(
                chartType(a.chartType),
                resolveRange(wb, sheet, a.dataRange),
                Excel.ChartSeriesBy.auto
              );
              if (a.title) ch.title.text = a.title;
              break;
            }
            default: {
              // Not silently dropped: the card showed this row, so the summary
              // has to say why it did nothing.
              skipped++;
              failures.push(`${describe(a)}: loại hành động không được hỗ trợ — bỏ qua.`);
              continue;
            }
          }
          // Sync after each action individually (instead of once at the end)
          // so a single bad range/malformed action surfaces its own error
          // and can be skipped, rather than aborting — or silently losing
          // track of which action failed in — the whole batch.
          await context.sync();
          applied++;
          ran.push(a);
        } catch (actionErr) {
          skipped++;
          remaining.push(a);
          failures.push(`${describe(a)}: ${actionErr.message || actionErr}`);
        }
      }
    });
    const total = pendingActions.length;
    const summary =
      skipped > 0
        ? `Đã áp dụng ${applied}/${total} hành động (${skipped} bị bỏ qua).`
        : `Đã áp dụng ${applied} hành động.`;
    addBubble("bot", summary, "ok");
    showToast(summary, { tone: skipped > 0 ? "warn" : "ok" });
    if (failures.length > 0) addBubble("bot", "⚠ " + failures.join("; "), "err");
    lastSig = null; // workbook changed — re-send fresh data on the next turn

    // A Partial Apply: keep the unrun Actions on the card instead of discarding
    // the whole Proposal — they used to be unrecoverable. See ADR-0004 for what
    // this costs and why the card has to say so.
    // Re-baseline only against a sheet we can still NAME. Without the
    // proposalSheetName guard, getSnapshot() falls back to the active sheet,
    // which is the retargeting this whole branch exists to prevent.
    const canKeepRemainder =
      remaining.length > 0 && Boolean(proposalSheetName) && !retargetsSheets(ran);
    const after = canKeepRemainder ? await getSnapshot(proposalSheetName) : null;

    if (after) {
      // Re-baselined against the PROPOSAL's sheet, read by name. Reading the
      // active sheet instead silently retargeted the Remainder: a newSheet
      // Action activates the sheet it creates, so proposalSheetName became the
      // new sheet and the unrun Actions' bare ranges would have been written
      // there on the next press — past a guard that had just been re-baselined
      // to agree.
      proposalSheetName = after.name;
      proposalSig = contentSignature(after);
      pendingActions = remaining;
      afterPartialApply = true;
      renderActions(remaining);
      addBubble(
        "bot",
        `${remaining.length} hành động chưa chạy vẫn còn trên thẻ — sửa nguyên nhân ở trên rồi bấm Áp dụng lần nữa.`,
        "warn"
      );
      setStatus(`Còn ${remaining.length} hành động chưa áp dụng.`, "warn");
      if (typeof window.__hermesRefreshContext === "function") window.__hermesRefreshContext(after);
    } else {
      if (remaining.length > 0) {
        // A bare range in the unrun Actions no longer has a target we can name,
        // and guessing one is exactly what Apply exists to prevent. Say which
        // of the two reasons it was — "it changed underneath you" and "I cannot
        // find your sheet any more" call for different next moves.
        const reason = retargetsSheets(ran)
          ? "đề xuất này tạo hoặc đổi tên sheet"
          : `không còn đọc được sheet "${proposalSheetName || "?"}"`;
        addBubble(
          "bot",
          `${remaining.length} hành động chưa chạy đã bị bỏ: ${reason}, nên không còn ` +
            `xác định được chúng sẽ ghi vào đâu. Hãy hỏi lại Hermes.`,
          "warn"
        );
      }
      clearPending();
      proposalSheetName = null;
      proposalSig = null;
      setStatus(remaining.length > 0 ? "Cần hỏi lại Hermes." : "Sẵn sàng.");
      if (typeof window.__hermesRefreshContext === "function") window.__hermesRefreshContext(null);
    }
  } catch (e) {
    const errText = "⚠ " + e.message;
    addBubble("bot", errText, "err");
    setStatus("Lỗi.", "err");
    showToast(errText, { tone: "err", timeout: 6000 });
  } finally {
    setBusy(false);
  }
}

// ---- helpers ---------------------------------------------------------------

function clearPending() {
  pendingActions = [];
  afterPartialApply = false;
  document.getElementById("preview").innerHTML = "";
  document.getElementById("apply").hidden = true;
  document.getElementById("highlightWrap").hidden = true;
}

function setBusy(b, msg) {
  busy = b;
  setBusyUi(document.getElementById("ask"), b);
  if (msg) setStatus(msg, b ? "busy" : undefined);
}

function addBubble(who, text, tone) {
  // The first message replaces the empty state for this chat.
  if (typeof window.__hermesSetEmptyHidden === "function") window.__hermesSetEmptyHidden(true);
  appendMessage(document.getElementById("log"), who, text, { tone });
}

function setStatus(s, tone) {
  setStatusUi(document.getElementById("status"), s, { tone });
  // Mirror the tone onto the row so the bar tints while busy — the dot sits
  // outside #status, so setStatusUi's textContent write can't wipe it.
  const row = document.getElementById("statusRow");
  if (row) {
    if (tone) row.dataset.tone = tone;
    else delete row.dataset.tone;
  }
}
