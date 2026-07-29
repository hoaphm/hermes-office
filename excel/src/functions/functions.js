import { askHermes } from "../shared/hermes";
// Custom functions are drag-fillable, so a fill handle dragged down a column
// becomes one LLM call per cell. Every call below goes through this shared
// gate, which memoises answers, collapses duplicate concurrent requests, and
// caps how many are in flight — see shared/fn-cache.js.
import { createLimitedCache } from "../../../shared/fn-cache.js";

const cache = createLimitedCache();

function ikey(name, args) {
  return name + ":" + JSON.stringify(args);
}
function clean(s) {
  return String(s)
    .trim()
    .replace(/^["'`]+|["'`]+$/g, "")
    .trim();
}

// One call path for all four functions: same key is used both as the cache
// identity and as the provider Idempotency-Key header.
async function call(name, args, messages) {
  const key = ikey(name, args);
  const out = await cache.run(key, () => askHermes(messages, { idempotencyKey: key }));
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
  return call(
    "CLASSIFY",
    [value, instruction],
    [
      {
        role: "user",
        content: `Classify the text per this instruction: ${instruction}.\nReply with ONLY the label, no punctuation.\n\nText: ${value}`,
      },
    ]
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
  return call(
    "EXTRACT",
    [value, what],
    [
      {
        role: "user",
        content: `Extract the ${what} from the text. Reply with ONLY the value, or an empty string if none.\n\nText: ${value}`,
      },
    ]
  );
}

/**
 * Summarize a range with Hermes.
 * @customfunction SUMMARIZE
 * @param {any[][]} values Range of values.
 * @returns {Promise<string>} One-sentence summary.
 */
export async function summarize(values) {
  return call("SUMMARIZE", values, [
    {
      role: "user",
      content: `Summarize this data in ONE short sentence:\n${JSON.stringify(values)}`,
    },
  ]);
}

/**
 * Get an Excel formula from a description with Hermes.
 * @customfunction FORMULA_HELP
 * @param {string} goal What you want the formula to do.
 * @returns {Promise<string>} A single Excel formula.
 */
export async function formulaHelp(goal) {
  return call(
    "FORMULA_HELP",
    [goal],
    [
      {
        role: "user",
        content: `Give a single Excel formula that accomplishes: ${goal}. Reply with ONLY the formula, starting with =.`,
      },
    ]
  );
}
