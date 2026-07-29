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
        return dedupeEdits(
          edits
            .filter((e) => e && typeof e.find === "string" && e.find.length > 0)
            .map((e) => ({ find: e.find, replace: e.replace === undefined ? "" : String(e.replace) }))
        );
      }
    } catch {
      /* try next candidate */
    }
  }
  return [];
}

// Models routinely list the same correction once per occurrence in the
// document ("hoc" -> "học", twice). Apply replaces ALL matches of a find
// string, so the second copy matches nothing and got counted as skipped —
// reporting a failure on a run where everything worked. A different replace
// for the same find is a real (if risky) instruction, so only exact
// find+replace pairs collapse.
function dedupeEdits(edits) {
  const seen = new Set();
  return edits.filter((e) => {
    const key = `${e.find}\u0000${e.replace}`;
    // A raw separator byte, written as an escape so this file stays plain
    // text — an actual NUL made `file` report "data" and grep skip the
    // file entirely. It must not be a space: {find:"a b", replace:"c"} and
    // {find:"a", replace:"b c"} would then share a key and the second
    // would be dropped as a false duplicate.
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

// Word's Range.search() rejects find strings longer than 255 characters.
export const MAX_SEARCH_LEN = 255;

// ...and it has no matchWholeWord in this flow, so it matches SUBSTRINGS and
// Apply rewrites every hit. A 2-character find such as "va" (-> "và") would
// therefore also corrupt the inside of "van", "vay", "vang", "vai". Refuse
// those instead of guessing.
//
// The floor is deliberately 3 rather than higher: "hoc" -> "học" is a
// 3-character correction and among the most common Vietnamese typos, so a
// stricter bar would discard far more real corrections than it prevents.
export const MIN_SEARCH_LEN = 3;

// Split edits into what Apply can safely search for and what it must refuse.
// Returning the refused ones (rather than dropping them) is the point: the
// user asked for a spell-check, so a correction we decline to make has to be
// reported, not silently swallowed.
export function partitionEdits(edits, { minLen = MIN_SEARCH_LEN, maxLen = MAX_SEARCH_LEN } = {}) {
  const applicable = [];
  const tooShort = [];
  const tooLong = [];
  for (const e of edits || []) {
    if (!e || typeof e.find !== "string" || e.find.length === 0) continue;
    // Count real characters, and ignore padding: " va " carries two
    // characters of signal, and whitespace must not buy a dangerous find
    // its way past the floor.
    const signal = [...e.find.trim()].length;
    if (signal < minLen) tooShort.push(e);
    else if (e.find.length > maxLen) tooLong.push(e);
    else applicable.push(e);
  }
  return { applicable, tooShort, tooLong };
}

// One pass over a full-document reply, returning everything the task pane
// needs to decide what to show:
//   edits      — parsed, deduped corrections
//   structured — the model actually returned an {"edits": [...]} object.
//                Distinguishes "complied, found nothing" (edits: []) from
//                "ignored the contract and wrote prose", which need different
//                messages: the first is success, the second is a miss.
//   prose      — the reply with the JSON payload removed, safe to show in the
//                chat bubble. Word used to print the whole raw reply, so a
//                correct `{"edits":[]}` surfaced to the user as literal JSON.
export function parseEditsReply(reply) {
  const raw = String(reply ?? "");
  const edits = parseEdits(raw);

  // Locate the payload we parsed so it can be cut from the displayed text.
  const fenced = raw.match(/```(?:json)?\s*([\s\S]*?)```/i);
  let payload = null;
  if (fenced && /"edits"\s*:/.test(fenced[1])) payload = fenced[0];
  else {
    const bare = extractJsonObject(raw, "edits");
    if (bare) payload = bare;
  }

  let structured = false;
  if (payload) {
    const inner = fenced && payload === fenced[0] ? fenced[1] : payload;
    try {
      structured = Array.isArray(JSON.parse(inner.trim()).edits);
    } catch {
      structured = false;
    }
  }

  const prose = (payload ? raw.replace(payload, "") : raw).trim();
  return { edits, structured, prose };
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

export function signature(s) {
  const selPart = s.selection ? `${s.selection.address}|${hash(JSON.stringify(s.selection.values))}` : "";
  return `${s.name}|${s.address}|${s.values.length}|${hash(JSON.stringify(s.values))}|${selPart}`;
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
