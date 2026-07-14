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

// Accept either a fenced ```json block OR a bare { "edits": [...] } object
// that may be embedded in prose. Try fenced first, then a loose match.
export function parseEdits(reply) {
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
    } catch (_) {
      /* try next candidate */
    }
  }
  return [];
}

export function parseTableChanges(reply) {
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
