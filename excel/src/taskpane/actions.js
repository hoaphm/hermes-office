// Pure logic behind the Excel task pane's Proposals: turning a model reply into
// Actions, describing them for the card, and making their values safe to write.
// No Office.js here on purpose — this is the part that can be tested.
import { extractJsonObject } from "../../../shared/parsers.js";

// A Proposal larger than this is refused rather than rendered: the card becomes
// unreviewable, and Apply is only a security boundary if the user can read it.
export const MAX_ACTIONS = 100;

// Range.values evaluates any string starting with =, +, -, or @ as a formula,
// exactly like Range.formulas. Model-proposed cell values can be influenced by
// sheet content that's round-tripped into the prompt, so force such strings to
// literal text (the same leading-apostrophe convention the Excel UI uses) —
// otherwise a poisoned cell could get an AI-proposed value silently applied as
// a live formula (e.g. WEBSERVICE-based exfiltration).
export function literalCellValue(v) {
  return typeof v === "string" && /^[=+\-@]/.test(v) ? "'" + v : v;
}

export function literalizeGrid(values) {
  return (values || []).map((row) => row.map(literalCellValue));
}

// Split a reply into prose and Actions. The JSON is normally fenced; the
// unfenced fallback is brace-balanced rather than a greedy /\{.*"actions".*\}/,
// which swallowed everything from the first `{` in the prose to the last `}` in
// the reply and then failed to parse.
export function splitReply(raw) {
  let actions = [];
  let prose = raw;
  const fenced = raw.match(/```json\s*([\s\S]*?)```/i);
  const target = fenced ? fenced[1] : extractJsonObject(raw, "actions");
  if (target) {
    try {
      const obj = JSON.parse(target);
      actions =
        obj.actions || (obj.editPlan ? obj.editPlan.map((e) => ({ type: "setCell", ...e })) : []);
    } catch {
      /* leave actions empty */
    }
    prose = raw.replace(fenced ? fenced[0] : target, "").trim();
  }
  return {
    prose: prose || "(các thay đổi đề xuất bên dưới)",
    actions: Array.isArray(actions) ? actions : [],
  };
}

// One line per Action for the Proposal card. This is what the user actually
// reads before pressing Apply, so an unknown Action type is shown raw rather
// than hidden.
export function describe(a) {
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

// Excel table names accept only letters, digits and underscore, and cannot
// start with a digit.
export function tableName(n) {
  return String(n)
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/^[^A-Za-z_]/, "_");
}
