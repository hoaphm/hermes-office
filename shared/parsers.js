// Pure, Office.js-free helpers shared between the Word and Excel task panes.
// No npm workspace exists between the two add-ins, so this repo-root shared/
// folder (imported via relative paths) is how the logic is deduped instead
// of being copy-pasted into word/src/taskpane/taskpane.js and
// excel/src/taskpane/taskpane.js. See shared/hermes.js for the Hermes client.

// Base-26 column index (0-based) <-> letters, e.g. 0 -> "A", 25 -> "Z",
// 26 -> "AA". Plain `String.fromCharCode(65 + i)` only covers single
// letters and silently wraps/collides past column Z.
export function columnIndexToLetters(index) {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

export function columnLettersToIndex(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}

// Index of the `}` that closes the `{` at `start`, or -1. String-aware, so
// braces inside JSON string values (or escaped quotes) don't throw off the
// depth count.
function matchBrace(text, start) {
  let depth = 0;
  let inString = false;
  let escaped = false;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === "{") depth++;
    else if (ch === "}" && --depth === 0) return i;
  }
  return -1;
}

// Pull the brace-balanced JSON object that contains `"<key>"` out of arbitrary
// prose. Scanning braces (rather than regex-matching) is what makes nested
// arrays/objects safe: a lazy `\[[\s\S]*?\]` stops at the first inner `]` and
// yields truncated, unparseable JSON.
export function extractJsonObject(text, key) {
  const keyIdx = text.indexOf(`"${key}"`);
  if (keyIdx === -1) return null;
  // Walk outward from the nearest preceding `{` until we find one whose
  // matching `}` sits after the key — i.e. the object that actually holds it.
  let start = text.lastIndexOf("{", keyIdx);
  while (start !== -1) {
    const end = matchBrace(text, start);
    if (end > keyIdx) return text.slice(start, end + 1);
    start = start > 0 ? text.lastIndexOf("{", start - 1) : -1;
  }
  return null;
}

// Accept either a fenced ```json block OR a bare { "edits": [...] } object
// that may be embedded in prose. Try fenced first, then a loose match.
export function parseEdits(reply) {
  const candidates = [];
  const fenced = reply.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced) candidates.push(fenced[1]);
  const bare = extractJsonObject(reply, "edits");
  if (bare) candidates.push(bare);

  for (const c of candidates) {
    try {
      const obj = JSON.parse(c.trim());
      const edits = obj && obj.edits;
      if (Array.isArray(edits)) {
        return edits
          .filter((e) => e && typeof e.find === "string" && e.find.length > 0)
          .map((e) => ({ find: e.find, replace: e.replace === undefined ? "" : String(e.replace) }));
      }
    } catch {
      /* try next candidate */
    }
  }
  return [];
}

export function parseTableChanges(reply) {
  const candidates = [];
  const fenced = reply.match(/```(?:json)?\s*([\s\S]*?)```/i);
  if (fenced) candidates.push(fenced[1]);
  const bare = extractJsonObject(reply, "cells");
  if (bare) candidates.push(bare);

  for (const c of candidates) {
    try {
      const obj = JSON.parse(c.trim());
      // Guard the shape: `return obj.cells || []` happily returned a non-array
      // (e.g. an object) and blew up downstream on .length / .map.
      if (obj && Array.isArray(obj.cells)) {
        return obj.cells.filter(
          (x) => x && typeof x.cell === "string" && x.value !== undefined
        );
      }
    } catch {
      /* try next candidate */
    }
  }
  return [];
}

// djb2-style string hash, used to build the Excel taskpane's change-detection
// signature (cheap and stable — not cryptographic).
export function hash(str) {
  let h = 5381;
  for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
  return h;
}

// Fingerprint of the sheet CONTENT alone. Apply compares this to decide whether
// the sheet still looks the way it did when the Proposal was reviewed. It must
// exclude the selection: moving the cursor between Ask and Apply is normal and
// must not be mistaken for the sheet changing underneath the Proposal.
export function contentSignature(s) {
  return `${s.name}|${s.address}|${s.values.length}|${hash(JSON.stringify(s.values))}`;
}

// Content plus selection — what decides whether to re-send the Snapshot to the
// Provider, where a changed selection IS a reason to re-send.
export function signature(s) {
  const selPart = s.selection ? `${s.selection.address}|${hash(JSON.stringify(s.selection.values))}` : "";
  return `${contentSignature(s)}|${selPart}`;
}

// Resolves an A1 address (optionally "Sheet!A1") to an Office.js Range,
// given the workbook and a fallback sheet for unqualified addresses. Takes
// wb/fallbackSheet as parameters rather than importing Excel itself, so this
// stays Office.js-free and unit-testable with plain object mocks.
export function resolveRange(wb, fallbackSheet, addr) {
  addr = String(addr || "").trim();
  if (addr.includes("!")) {
    const i = addr.lastIndexOf("!");
    const sn = addr.slice(0, i).replace(/^'|'$/g, "").replace(/''/g, "'");
    return wb.worksheets.getItem(sn).getRange(addr.slice(i + 1));
  }
  return fallbackSheet.getRange(addr);
}

export function chartType(t) {
  const m = {
    columnclustered: "ColumnClustered", column: "ColumnClustered", columns: "ColumnClustered",
    bar: "BarClustered", barclustered: "BarClustered",
    line: "Line", pie: "Pie", doughnut: "Doughnut", area: "Area",
    scatter: "XYScatter", xyscatter: "XYScatter",
  };
  return m[String(t || "").toLowerCase().replace(/[^a-z]/g, "")] || "ColumnClustered";
}

// Cap replayed conversation turns so a long session doesn't grow the request
// payload without limit — both task panes enforce this with the same policy
// (MAX_HISTORY_MESSAGES) but different array shapes: Excel keeps a standing
// system message at index 0 and trims in place; Word has no standing system
// message and reassigns its `messages` variable each turn.
//
// @param {object[]} history
// @param {number} max turns to keep (excluding the system message when keepSystem)
// @param {{keepSystem?: boolean}} [opts] keepSystem: never drop index 0, cap
//   everything after it (splices `history` in place). Otherwise caps the
//   whole array from the end (returns a new array; does not mutate).
// @returns {{history: object[], trimmed: boolean}}
export function trimHistory(history, max, { keepSystem = false } = {}) {
  if (keepSystem) {
    const overflow = history.length - (max + 1);
    if (overflow <= 0) return { history, trimmed: false };
    history.splice(1, overflow);
    return { history, trimmed: true };
  }
  if (history.length <= max) return { history, trimmed: false };
  return { history: history.slice(-max), trimmed: true };
}
