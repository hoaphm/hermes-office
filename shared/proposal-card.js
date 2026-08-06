// Shared UI helpers for the Word + Excel task panes. Both panes import from
// their own src/shared/ re-export shim — see word/src/shared/hermes.js for
// the pattern; we mirror it here for proposal-card and toast helpers.
//
// Anything that talks to the Office host (Word.run / Excel.run) lives in
// the addin; this file is pure DOM and HTML strings. Two reasons:
//   1. Keeps the shared bundle tiny and CSS-only friendly.
//   2. Makes it trivial to unit-test the rendering in plain node.

import {
  columnIndexToLetters,
  columnLettersToIndex,
  parseEdits,
  parseTableChanges,
} from "./parsers.js";

const esc = (s) =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

// ---- Message bubbles --------------------------------------------------------

/**
 * Append a chat message to the log.
 *
 * @param {HTMLElement} log the scrollable log element
 * @param {"user"|"bot"} role
 * @param {string} text already-escaped (we set textContent)
 * @param {{tone?: "ok"|"warn"|"err"}} [opts]
 * @returns {HTMLElement} the appended .ds-msg element
 */
export function appendMessage(log, role, text, { tone } = {}) {
  const wrap = document.createElement("div");
  wrap.className = "ds-msg";
  wrap.dataset.role = role;
  if (tone) wrap.dataset.tone = tone;

  const bubble = document.createElement("div");
  bubble.className = "ds-bubble";
  bubble.textContent = text;
  wrap.appendChild(bubble);

  log.appendChild(wrap);
  const scroller = log.closest(".ds-content") || log;
  scroller.scrollTop = scroller.scrollHeight;
  return wrap;
}

// ---- Typing indicator ------------------------------------------------------

/**
 * Append a "Hermes is thinking" indicator to the log. It is a bot-side
 * message with three pulsing dots (styled by .ds-typing in design-system.css).
 * The caller removes it with removeTypingIndicator() once the reply lands.
 *
 * @param {HTMLElement} log the scrollable log element
 * @returns {HTMLElement} the appended .ds-msg element
 */
export function appendTypingIndicator(log) {
  const wrap = document.createElement("div");
  wrap.className = "ds-msg ds-typing";
  wrap.dataset.role = "bot";
  wrap.setAttribute("aria-label", "Hermes đang suy nghĩ");

  const bubble = document.createElement("div");
  bubble.className = "ds-bubble";
  for (let i = 0; i < 3; i++) {
    const dot = document.createElement("span");
    dot.className = "ds-dot";
    dot.setAttribute("aria-hidden", "true");
    bubble.appendChild(dot);
  }
  wrap.appendChild(bubble);

  log.appendChild(wrap);
  const scroller = log.closest(".ds-content") || log;
  scroller.scrollTop = scroller.scrollHeight;
  return wrap;
}

/**
 * Remove a typing indicator previously returned by appendTypingIndicator().
 * Safe to call with null / detached elements.
 *
 * @param {HTMLElement|null} el
 */
export function removeTypingIndicator(el) {
  if (el && el.parentElement) el.parentElement.removeChild(el);
}

// ---- Context bar -----------------------------------------------------------

/**
 * Mount the context chip row inside a header element (or any container).
 * Re-renders cheaply on every call — the caller passes the new chip set.
 *
 * @param {HTMLElement} host a `ds-header` element (or any container)
 * @param {Array<{label: string, value?: string, state?: "pinned"|"warn"}>} chips
 * @returns {HTMLElement} the chip-row element
 */
export function mountContextBar(host, chips) {
  let row = host.querySelector(".ds-context");
  if (!row) {
    row = document.createElement("div");
    row.className = "ds-context";
    row.setAttribute("role", "list");
    host.appendChild(row);
  }
  row.innerHTML = chips
    .filter((c) => c && (c.value || c.label))
    .map((c) => {
      const state = c.state ? ` data-state="${esc(c.state)}"` : "";
      const value = c.value ? `<strong>${esc(c.value)}</strong>` : "";
      const label = c.label ? `<span>${esc(c.label)}</span>` : "";
      return `<span class="ds-chip"${state} role="listitem">${label}${value}</span>`;
    })
    .join("");
  return row;
}

// ---- Proposal card (the core component) -----------------------------------

// A grid Action can carry thousands of cells. The card must still SHOW what
// will be written — Apply is only a review boundary if the payload is on
// screen — so render a bounded window of it rather than a row count alone.
const PREVIEW_ROWS = 8;
const PREVIEW_COLS = 8;

// Action types that write without first checking what is already there. Only
// setCell compares against a captured old value, so after a Partial Apply has
// moved the document underneath the Remainder, setCell is the only Action still
// covered by the review the user gave. These are the ones the card badges.
// See CONTEXT.md ("Partial Apply") and ADR-0004.
const BLIND_WRITE_TYPES = new Set([
  "setCells",
  "format",
  "createTable",
  "createChart",
]);

export function writesBlind(action) {
  return Boolean(action) && BLIND_WRITE_TYPES.has(action.type);
}

/**
 * The banner a Remainder carries. Says why the card changed and how much of it
 * the earlier review no longer covers — not a bare count, and not silence.
 *
 * @param {object[]} actions the Remainder
 * @returns {string} plain text; renderProposalCard escapes it
 */
export function partialApplyNotice(actions) {
  const blind = (actions || []).filter(writesBlind).length;
  if (blind === 0) {
    return (
      "Sheet đã đổi sau lần áp dụng trước. Các hành động còn lại vẫn kiểm tra " +
      "giá trị hiện tại trước khi ghi."
    );
  }
  return (
    `Sheet đã đổi sau lần áp dụng trước. ${blind} hành động dưới đây ghi đè ` +
    "không kiểm tra giá trị hiện tại."
  );
}

/**
 * Flatten a 2-D grid into a bounded, human-readable block for the card.
 *
 * @param {any[][]} values
 * @returns {string} plain text; the caller escapes it
 */
export function previewGrid(values) {
  const rows = Array.isArray(values) ? values : [];
  const lines = rows.slice(0, PREVIEW_ROWS).map((row) => {
    const cells = Array.isArray(row) ? row : [row];
    const shown = cells.slice(0, PREVIEW_COLS).map((c) => String(c ?? ""));
    const restCols = cells.length - PREVIEW_COLS;
    return shown.join("  │  ") + (restCols > 0 ? `  │  … +${restCols} cột` : "");
  });
  const restRows = rows.length - PREVIEW_ROWS;
  if (restRows > 0) lines.push(`… +${restRows} dòng nữa`);
  return lines.join("\n");
}

/**
 * Describe a single action.
 *
 * Returns PLAIN, UNESCAPED text. renderProposalCard escapes every field it
 * interpolates, so escaping here too produced visible double-encoding —
 * a title of `R&D "2024"` rendered as `R&amp;D &quot;2024&quot;`. Escaping
 * belongs at the single point where a string enters innerHTML, not here.
 *
 * `detail` is the payload the action will write, shown verbatim under the
 * summary. An action that writes content the summary cannot fit MUST set it.
 *
 * @param {object} a action (Word edits[] or Excel actions[])
 * @returns {{summary: string, kind: "diff"|"info", diff?: {old, new}, detail?: string}}
 */
export function describeAction(a) {
  if (!a || typeof a !== "object") return { summary: String(a), kind: "info" };
  switch (a.type) {
    case "setCell":
      return {
        summary: `Set ${String(a.cell ?? "")}`,
        kind: "diff",
        diff: { old: String(a.old ?? ""), new: String(a.new ?? "") },
      };
    case "setCells": {
      const rows = Array.isArray(a.values) ? a.values : [];
      return {
        summary: `Fill ${a.range} · ${rows.length} rows`,
        kind: "info",
        detail: previewGrid(rows),
      };
    }
    case "format":
      return {
        summary:
          `Format ${a.range}` +
          (a.numberFormat ? ` as ${a.numberFormat}` : "") +
          (a.bold ? " (bold)" : "") +
          (a.fill ? ` fill ${a.fill}` : ""),
        kind: "info",
      };
    case "createTable":
      return {
        summary: `Create table "${a.name || "Table"}" over ${a.range}`,
        kind: "info",
      };
    case "createChart":
      return {
        summary:
          `Create ${a.chartType || "Column"} chart from ${a.dataRange}` +
          (a.title ? ` — "${a.title}"` : ""),
        kind: "info",
      };
    case "newSheet":
      return { summary: `New sheet "${a.name}"`, kind: "info" };
    case "renameSheet":
      return { summary: `Rename active tab → "${a.to || a.name}"`, kind: "info" };
    case "replace": {
      // Word edit shape: { find, replace, matchCount }.
      const find = a.find ?? a.find_text ?? "";
      const replace = a.replace ?? a.replace_text ?? "";
      // Apply replaces EVERY match, case-insensitively. The count has to be on
      // the card: warning about it during Apply is a warning after the write.
      // Undefined (the host could not count) simply shows no claim.
      const n = typeof a.matchCount === "number" ? a.matchCount : null;
      let summary = "Replace";
      if (n === 0) summary += " · không còn tìm thấy trong tài liệu";
      else if (n === 1) summary += " · 1 chỗ";
      else if (n !== null) summary += ` · TẤT CẢ ${n} chỗ`;
      return {
        summary,
        kind: "diff",
        diff: { old: String(find), new: String(replace) },
      };
    }
    case "insert":
      return {
        summary: `Insert at ${a.location || a.at || "cursor"}`,
        kind: "info",
      };
    case "delete":
      return { summary: `Delete ${a.range || a.text || ""}`, kind: "info" };
    default:
      return { summary: JSON.stringify(a), kind: "info" };
  }
}

/**
 * Render a proposal card into a host element. The host's previous children
 * are removed and a new card is appended.
 *
 * The card is header (inverted) + action list (each row is label + optional
 * diff old/new chips). It deliberately renders NO call-to-action button: both
 * task panes own a footer-level #apply button, and a second in-card button
 * would be a decorative duplicate of it. Callers should show/hide their own
 * #apply based on whether this returns a card.
 *
 * `notice` renders as a banner directly under the header, and `badgeAction`
 * returns a short label to pin on an individual action (or falsy for none).
 * Both exist for the Remainder a Partial Apply leaves behind: the banner says
 * why the card is no longer what the user reviewed, the badges say which rows
 * that applies to. Reporting only the banner would summarise into a single
 * claim exactly what the Apply boundary requires per-action — see CONTEXT.md.
 *
 * @param {HTMLElement} host any container (e.g. #preview)
 * @param {{
 *   title?: string,
 *   actions: object[],
 *   notice?: string,
 *   badgeAction?: (action: object) => string|null,
 * }} proposal
 * @returns {HTMLElement|null} the rendered card element, or null if no actions
 */
export function renderProposalCard(host, proposal) {
  host.innerHTML = "";
  if (!proposal || !Array.isArray(proposal.actions) || proposal.actions.length === 0) {
    return null;
  }
  const card = document.createElement("section");
  card.className = "ds-card";
  card.setAttribute("role", "group");
  card.setAttribute("aria-label", proposal.title || "Thay đổi đề xuất");

  // Header (inverted bg — focal point)
  const head = document.createElement("div");
  head.className = "ds-card-head";
  const titleText = proposal.title || `${proposal.actions.length} hành động đề xuất`;
  head.innerHTML =
    `<span class="ds-card-title">${esc(titleText)}</span>` +
    `<span class="ds-card-tag">${proposal.actions.length} mục</span>`;
  card.appendChild(head);

  // Banner (Remainder of a Partial Apply — normally absent)
  if (proposal.notice) {
    const notice = document.createElement("div");
    notice.className = "ds-card-notice";
    notice.setAttribute("role", "note");
    notice.textContent = proposal.notice;
    card.appendChild(notice);
  }

  // Action list
  const list = document.createElement("ul");
  list.className = "ds-card-list";
  const badgeAction = proposal.badgeAction;
  proposal.actions.forEach((a) => {
    const { summary, kind, diff, detail } = describeAction(a);
    const li = document.createElement("li");
    li.className = "ds-card-action";
    const badge = badgeAction ? badgeAction(a) : null;
    let html =
      `<div class="label">${esc(summary)}` +
      (badge ? `<span class="ds-card-badge">${esc(badge)}</span>` : "") +
      `</div>`;
    if (kind === "diff" && diff) {
      html +=
        `<div class="diff">` +
        `<div class="ds-diff-old">${esc(diff.old)}</div>` +
        `<div class="ds-diff-new">${esc(diff.new)}</div>` +
        `</div>`;
    }
    if (detail) html += `<div class="ds-card-detail">${esc(detail)}</div>`;
    li.innerHTML = html;
    list.appendChild(li);
  });
  card.appendChild(list);

  host.appendChild(card);
  return card;
}

// ---- Toast (replaces inline "result" bubbles for Apply) --------------------

let toastHost = null;
function ensureToastHost() {
  if (toastHost && document.body.contains(toastHost)) return toastHost;
  toastHost = document.createElement("div");
  toastHost.className = "ds-toast-host";
  toastHost.setAttribute("role", "region");
  toastHost.setAttribute("aria-label", "Thông báo");
  toastHost.setAttribute("aria-live", "polite");
  document.body.appendChild(toastHost);
  return toastHost;
}

/**
 * Show a transient toast at the bottom of the pane.
 *
 * @param {string} text the message (we set textContent; no HTML injection)
 * @param {{tone?: "ok"|"warn"|"err"|"info", timeout?: number}} [opts]
 *   timeout defaults to 3500ms; pass 0 to make the toast sticky
 * @returns {HTMLElement} the toast element
 */
export function showToast(text, { tone = "info", timeout = 3500 } = {}) {
  const host = ensureToastHost();
  const t = document.createElement("div");
  t.className = "ds-toast";
  t.dataset.tone = tone;
  t.setAttribute("role", tone === "err" ? "alert" : "status");
  const body = document.createElement("span");
  body.textContent = text;
  t.appendChild(body);
  const dismiss = document.createElement("button");
  dismiss.type = "button";
  dismiss.className = "ds-toast-dismiss";
  dismiss.setAttribute("aria-label", "Đóng");
  dismiss.textContent = "×";
  dismiss.addEventListener("click", () => removeToast(t));
  t.appendChild(dismiss);
  host.appendChild(t);
  if (timeout > 0) {
    setTimeout(() => removeToast(t), timeout);
  }
  return t;
}

function removeToast(t) {
  if (!t || !t.parentElement) return;
  t.parentElement.removeChild(t);
}

// ---- Status row ------------------------------------------------------------

/**
 * Set the status line. tone="err" / "busy" / "ok" drives the dot animation
 * and the aria-live announcement.
 */
export function setStatus(statusEl, text, { tone } = {}) {
  if (!statusEl) return;
  statusEl.textContent = text || "";
  if (tone) {
    statusEl.dataset.tone = tone;
  } else {
    delete statusEl.dataset.tone;
  }
}

/**
 * Toggle the busy state on a button. We use aria-disabled (not the native
 * `disabled` attribute) so the button stays focusable and the spinner
 * pseudo-element renders.
 */
export function setBusy(btn, busy) {
  if (!btn) return;
  if (busy) {
    btn.setAttribute("aria-disabled", "true");
    btn.dataset.busy = "true";
  } else {
    btn.removeAttribute("aria-disabled");
    delete btn.dataset.busy;
  }
}

// ---- Word-specific helpers -------------------------------------------------

export function parseWordEdits(raw) {
  return parseEdits(raw);
}

export function parseWordTableChanges(raw) {
  return parseTableChanges(raw);
}

// ---- Reviewability limit ---------------------------------------------------

// A Proposal larger than this is refused rather than rendered: the card
// becomes unreviewable, and Apply is only a security boundary if the user can
// read what it will write. Canonical value — Word guards on it and Excel
// re-exports it from actions.js so the public import surface stays stable.
export const MAX_ACTIONS = 100;

const REVIEW_LIMIT_ERROR =
  "Đề xuất có quá nhiều thay đổi để xem xét an toàn";

/**
 * Throw the UI-safe Vietnamese error when a proposal holds more Actions than
 * the reviewable limit; otherwise return the array unchanged so the callsite
 * can chain. MUST run before any Proposal is created/rendered — the oversized
 * payload is rejected, never truncated.
 * @param {Array<unknown>} actions parsed Actions to review
 * @returns {Array<unknown>} the same array, when within the limit
 */
export function assertReviewableActions(actions) {
  if (actions.length > MAX_ACTIONS) {
    throw new Error(REVIEW_LIMIT_ERROR);
  }
  return actions;
}

export { columnIndexToLetters, columnLettersToIndex };
