// Excel worksheet functions (`=HERMES.*`). These reach the Provider from a
// cell, OUTSIDE the Apply boundary the task pane enforces — a workbook that
// merely CONTAINS one of these formulas fires it on open or recalculation, with
// no card to review and no button to press.
//
// That is why every function here is gated on `customFunctions: true` in
// config.json. config.json is a local file served by the Local Gateway, so a
// document can never turn the gate on for itself. Off by default: opening an
// untrusted workbook must not be enough to spend your quota or ship its
// contents to your Provider. See CONTEXT.md and the README.
import { askHermes, customFunctionsEnabled } from "../shared/hermes";

const DISABLED_MESSAGE =
  'Hàm =HERMES.* đang tắt. Thêm "customFunctions": true vào config.json rồi tải lại Excel để bật.';

function ikey(name, args) {
  return name + ":" + JSON.stringify(args);
}
function clean(s) {
  return String(s)
    .trim()
    .replace(/^["'`]+|["'`]+$/g, "")
    .trim();
}

// Every function funnels through here so the gate cannot be forgotten on a
// newly added one. A thrown Error surfaces in the cell as an error value.
async function run(name, prompt, args) {
  if (!(await customFunctionsEnabled())) throw new Error(DISABLED_MESSAGE);
  const out = await askHermes([{ role: "user", content: prompt }], {
    idempotencyKey: ikey(name, args),
  });
  return clean(out);
}

/**
 * Classify a value with Hermes.
 * @customfunction CLASSIFY
 * @param {string} value Text to classify.
 * @param {string} instruction How to classify, e.g. "lead quality: hot/warm/cold".
 * @returns {Promise<string>} The label.
 */
export async function classify(value, instruction) {
  return run(
    "CLASSIFY",
    `Classify the text per this instruction: ${instruction}.\nReply with ONLY the label, no punctuation.\n\nText: ${value}`,
    [value, instruction]
  );
}

/**
 * Extract a field from text with Hermes.
 * @customfunction EXTRACT
 * @param {string} value Source text.
 * @param {string} what What to extract, e.g. "company name".
 * @returns {Promise<string>} The extracted value.
 */
export async function extract(value, what) {
  return run(
    "EXTRACT",
    `Extract the ${what} from the text. Reply with ONLY the value, or an empty string if none.\n\nText: ${value}`,
    [value, what]
  );
}

/**
 * Summarize a range with Hermes.
 * @customfunction SUMMARIZE
 * @param {any[][]} values Range of values.
 * @returns {Promise<string>} One-sentence summary.
 */
export async function summarize(values) {
  return run(
    "SUMMARIZE",
    `Summarize this data in ONE short sentence:\n${JSON.stringify(values)}`,
    values
  );
}

/**
 * Get an Excel formula from a description with Hermes.
 * @customfunction FORMULA_HELP
 * @param {string} goal What you want the formula to do.
 * @returns {Promise<string>} A single Excel formula.
 */
export async function formulaHelp(goal) {
  return run(
    "FORMULA_HELP",
    `Give a single Excel formula that accomplishes: ${goal}. Reply with ONLY the formula, starting with =.`,
    [goal]
  );
}
