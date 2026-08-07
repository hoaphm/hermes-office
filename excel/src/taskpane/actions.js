// Pure logic behind the Excel task pane's Proposals: turning a model reply into
// Actions, describing them for the card, and making their values safe to write.
// No Office.js here on purpose — this is the part that can be tested.
import { extractJsonObject } from "../../../shared/parsers.js";
import { MAX_ACTIONS, describeAction } from "../../../shared/proposal-card.js";
export { MAX_ACTIONS };

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

// One line per Action for the Apply-failure bubbles. Delegates to the shared
// describeAction() so this and the Proposal card can never drift apart — the
// card renders the richer {summary, diff, detail} shape; these flat bubbles
// re-fold the diff into the summary line.
export function describe(a) {
  const d = describeAction(a);
  if (d.diff) return `${d.summary}:  "${d.diff.old}" → "${d.diff.new}"`;
  return d.summary;
}

// Action types that change what an unqualified range means. `newSheet`
// activates the sheet it creates and `renameSheet` moves the name the Proposal
// was pinned to, so once one of them has run, the Actions of the same Proposal
// that did NOT run no longer have a well-defined target. There is no honest
// Remainder to offer in that case — refuse it. See ADR-0004.
const RETARGETING_TYPES = new Set(["newSheet", "renameSheet"]);

export function retargetsSheets(actions) {
  return (actions || []).some((a) => a && RETARGETING_TYPES.has(a.type));
}

// Excel table names accept only letters, digits and underscore, and cannot
// start with a digit.
export function tableName(n) {
  return String(n)
    .replace(/[^A-Za-z0-9_]/g, "_")
    .replace(/^[^A-Za-z_]/, "_");
}
