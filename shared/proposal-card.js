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

/**
 * Describe a single action. Returns a small object so the caller (or
 * renderProposalCard) can decide how to render it.
 *
 * @param {object} a action (Word edits[] or Excel actions[])
 * @returns {{summary: string, kind: "diff"|"info", diff?: {old, new}, label?: string}}
 */
export function describeAction(a) {
  if (!a || typeof a !== "object") return { summary: String(a), kind: "info" };
  switch (a.type) {
    case "setCell":
      return {
        summary: `Set ${escCell(a.cell)}`,
        kind: "diff",
        diff: { old: String(a.old ?? ""), new: String(a.new ?? "") },
      };
    case "setCells":
      return {
        summary: `Fill ${esc(a.range)} · ${(a.values || []).length} rows`,
        kind: "info",
      };
    case "format":
      return {
        summary:
          `Format ${esc(a.range)}` +
          (a.numberFormat ? ` as ${esc(a.numberFormat)}` : "") +
          (a.bold ? " (bold)" : ""),
        kind: "info",
      };
    case "createTable":
      return {
        summary: `Create table "${esc(a.name || "Table")}" over ${esc(a.range)}`,
        kind: "info",
      };
    case "createChart":
      return {
        summary:
          `Create ${esc(a.chartType || "Column")} chart from ${esc(a.dataRange)}` +
          (a.title ? ` — "${esc(a.title)}"` : ""),
        kind: "info",
      };
    case "newSheet":
      return { summary: `New sheet "${esc(a.name)}"`, kind: "info" };
    case "renameSheet":
      return { summary: `Rename active tab → "${esc(a.to || a.name)}"`, kind: "info" };
    case "replace": {
      // Word edit shape: { find, replace, all_occurrences }.
      const find = a.find ?? a.find_text ?? "";
      const replace = a.replace ?? a.replace_text ?? "";
      return {
        summary: `Replace`,
        kind: "diff",
        diff: { old: String(find), new: String(replace) },
      };
    }
    case "insert":
      return {
        summary: `Insert at ${esc(a.location || a.at || "cursor")}`,
        kind: "info",
      };
    case "delete":
      return { summary: `Delete ${esc(a.range || a.text || "")}`, kind: "info" };
    default:
      return { summary: JSON.stringify(a), kind: "info" };
  }
}

function escCell(c) {
  if (typeof c !== "string") return String(c ?? "");
  return esc(c);
}

/**
 * Render a proposal card into a host element. The host's previous children
 * are removed and a new card is appended.
 *
 * The card is header (inverted) + action list (each row is label + optional
 * diff old/new chips) + footer with the primary CTA. The host is responsible
 * for keeping its OWN apply button if it needs a footer-level CTA — we don't
 * add one here by default.
 *
 * @param {HTMLElement} host any container (e.g. #preview)
 * @param {{ title?: string, actions: object[], primaryLabel?: string }} proposal
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
  card.setAttribute("aria-label", proposal.title || "Proposed changes");

  // Header (inverted bg — focal point)
  const head = document.createElement("div");
  head.className = "ds-card-head";
  const titleText = proposal.title || `${proposal.actions.length} hành động đề xuất`;
  head.innerHTML =
    `<span class="ds-card-title">${esc(titleText)}</span>` +
    `<span class="ds-card-tag">${proposal.actions.length} mục</span>`;
  card.appendChild(head);

  // Action list
  const list = document.createElement("ul");
  list.className = "ds-card-list";
  proposal.actions.forEach((a) => {
    const { summary, kind, diff } = describeAction(a);
    const li = document.createElement("li");
    li.className = "ds-card-action";
    if (kind === "diff" && diff) {
      li.innerHTML =
        `<div class="label">${esc(summary)}</div>` +
        `<div class="diff">` +
        `<div class="ds-diff-old">${esc(diff.old)}</div>` +
        `<div class="ds-diff-new">${esc(diff.new)}</div>` +
        `</div>`;
    } else {
      li.innerHTML = `<div class="label">${esc(summary)}</div>`;
    }
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
  toastHost.setAttribute("aria-label", "Notifications");
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
  dismiss.setAttribute("aria-label", "Dismiss");
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

export { columnIndexToLetters, columnLettersToIndex };
