/* global Office, Excel, document */
import { askHermes } from "../shared/hermes";
// Pure helpers deduped with the Word taskpane via the repo-root shared/
// folder (no npm workspace between the two add-ins) — see shared/parsers.js.
import { signature, resolveRange, chartType } from "../../../shared/parsers.js";
// UI helpers (proposal card, toast, context bar) live in shared/proposal-card.js
// so both add-ins get the same design system + a11y treatment.
import {
  appendMessage,
  setStatus as setStatusUi,
  setBusy as setBusyUi,
  showToast,
  renderProposalCard,
} from "../../../shared/proposal-card.js";

const MAX_ROWS = 500;
// Very wide sheets can otherwise blow up the snapshot payload sent to Hermes.
const MAX_COLS = 100;
// Belt-and-suspenders byte cap on top of the row/col caps, in case a sheet
// is dense (long strings) rather than just tall/wide.
const MAX_SNAPSHOT_BYTES = 200000;

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

Office.onReady(() => {
  const askBtn = document.getElementById("ask");
  const newChatBtn = document.getElementById("newchat");
  const promptEl = document.getElementById("prompt");
  const applyBtn = document.getElementById("apply");

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
  const metaSep = document.querySelectorAll(".ds-meta-sep");

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
  }
  refreshMeta(null);
  window.__hermesRefreshContext = refreshMeta;
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
    proposalSheetName = snap.name;
    if (typeof window.__hermesRefreshContext === "function") window.__hermesRefreshContext(snap);
    const sig = signature(snap);
    let content = prompt;
    if (sig !== lastSig) {
      lastSig = sig;
      content = `${dataNote(snap)}\n\n${prompt}`;
    }
    history.push({ role: "user", content });

    setBusy(true, "Hermes đang suy nghĩ…");
    const raw = await askHermes(history);
    history.push({ role: "assistant", content: raw });

    const { prose, actions } = splitReply(raw);
    addBubble("bot", prose);
    pendingActions = actions;
    renderActions(actions);
    setStatus(
      actions.length
        ? `${actions.length} hành động được đề xuất — xem lại rồi Áp dụng.`
        : "Sẵn sàng."
    );
  } catch (e) {
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
  clearPending();
  document.getElementById("log").innerHTML = "";
  setStatus("Cuộc trò chuyện mới. Mở sheet dữ liệu và đặt câu hỏi.");
  if (typeof window.__hermesRefreshContext === "function") window.__hermesRefreshContext(null);
}

// ---- reading the sheet (only sent when changed) ----------------------------

async function getSnapshot() {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");

    // Read the user's current selection (single cell or range)
    let selection = null;
    try {
      const sel = context.workbook.getSelectedRange();
      sel.load(["address", "values"]);
      await context.sync();
      if (sel.address) selection = { address: sel.address, values: sel.values };
    } catch (_) {
      /* no selection or multi-area — ignore */
    }

    const used = sheet.getUsedRangeOrNullObject(true); // valuesOnly
    used.load(["address", "values"]);
    await context.sync();
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
    // strings can serialize to a huge payload — halve the row count until it
    // fits rather than sending a giant blob to Hermes.
    while (values.length > 1 && JSON.stringify(values).length > MAX_SNAPSHOT_BYTES) {
      values = values.slice(0, Math.ceil(values.length / 2));
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

// ---- parsing the agent's reply ---------------------------------------------

function splitReply(raw) {
  let actions = [];
  let prose = raw;
  const fenced = raw.match(/```json\s*([\s\S]*?)```/i);
  const target = fenced ? fenced[1] : (raw.match(/\{[\s\S]*"actions"[\s\S]*\}/) || [])[0];
  if (target) {
    try {
      const obj = JSON.parse(target);
      actions =
        obj.actions || (obj.editPlan ? obj.editPlan.map((e) => ({ type: "setCell", ...e })) : []);
    } catch (_) {
      /* leave actions empty */
    }
    prose = raw.replace(fenced ? fenced[0] : target, "").trim();
  }
  return {
    prose: prose || "(các thay đổi đề xuất bên dưới)",
    actions: Array.isArray(actions) ? actions : [],
  };
}

// ---- preview + apply -------------------------------------------------------

function describe(a) {
  switch (a.type) {
    case "setCell":
      return `Set ${a.cell}:  "${a.old ?? ""}" → "${a.new}"`;
    case "setCells":
      return `Fill ${a.range} (${(a.values || []).length} rows)`;
    case "format":
      return `Format ${a.range}${a.numberFormat ? ` as ${a.numberFormat}` : ""}${a.bold ? " (bold)" : ""}`;
    case "createTable":
      return `Create table "${a.name || "Table"}" over ${a.range}`;
    case "createChart":
      return `Create ${a.chartType || "Column"} chart from ${a.dataRange}${a.title ? ` — "${a.title}"` : ""}`;
    case "newSheet":
      return `New sheet "${a.name}"`;
    case "renameSheet":
      return `Rename active tab → "${a.to || a.name}"`;
    default:
      return JSON.stringify(a);
  }
}

function renderActions(actions) {
  const box = document.getElementById("preview");
  const card = renderProposalCard(box, {
    title: actions.length
      ? `${actions.length} hành động đề xuất cho sheet "${proposalSheetName || ""}"`
      : "",
    actions: actions,
    primaryLabel: "Áp dụng",
  });
  // renderProposalCard inserts an in-card Apply button — but we still need
  // the standalone #apply button visible too, because the rest of the layout
  // (and the existing keybindings) assume a footer-level CTA. The card's
  // button is decorative and the real action lives in #apply.
  document.getElementById("apply").style.display = card ? "block" : "none";
}

async function apply() {
  if (!pendingActions.length || busy) return;
  setBusy(true, "Đang áp dụng…");
  let applied = 0;
  let skipped = 0;
  const failures = [];
  try {
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
        } catch (_) {
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
              r.values = [[a.new]];
              r.format.fill.color = "#C6EFCE";
              break;
            }
            case "setCells": {
              resolveRange(wb, sheet, a.range).values = a.values;
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
              skipped++;
              continue;
            }
          }
          // Sync after each action individually (instead of once at the end)
          // so a single bad range/malformed action surfaces its own error
          // and can be skipped, rather than aborting — or silently losing
          // track of which action failed in — the whole batch.
          await context.sync();
          applied++;
        } catch (actionErr) {
          skipped++;
          failures.push(`${describe(a)}: ${actionErr.message || actionErr}`);
        }
      }
    });
    const summary =
      skipped > 0
        ? `Đã áp dụng ${applied}/${pendingActions.length} hành động (${skipped} bị bỏ qua).`
        : `Đã áp dụng ${applied} hành động.`;
    addBubble("bot", summary, "ok");
    showToast(summary, { tone: skipped > 0 ? "warn" : "ok" });
    if (failures.length > 0) addBubble("bot", "⚠ " + failures.join("; "), "err");
    clearPending();
    lastSig = null; // workbook changed — re-send fresh data on the next turn
    proposalSheetName = null;
    setStatus("Sẵn sàng.");
    if (typeof window.__hermesRefreshContext === "function") window.__hermesRefreshContext(null);
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

function tableName(n) {
  return String(n)
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/^[^A-Za-z_]/, "_");
}

function clearPending() {
  pendingActions = [];
  document.getElementById("preview").innerHTML = "";
  document.getElementById("apply").style.display = "none";
}

function setBusy(b, msg) {
  busy = b;
  setBusyUi(document.getElementById("ask"), b);
  if (msg) setStatus(msg, b ? "busy" : undefined);
}

function addBubble(who, text, tone) {
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
