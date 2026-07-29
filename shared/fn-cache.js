// Cache + concurrency gate for Excel's =HERMES.* custom functions.
//
// Custom functions are drag-fillable, so a single fill handle pulled down a
// column turns into one LLM request per cell. Nothing throttled or deduped
// them: 500 rows meant 500 concurrent billable calls, and re-filling the same
// values paid for the same answers again.
//
// Three behaviours, all keyed off the caller's idempotency key:
//   1. Memoise resolved answers (bounded, insertion-ordered eviction).
//   2. Share the in-flight promise for identical concurrent keys, so a fill
//      across duplicate values costs one request, not one per cell.
//   3. Cap how many distinct requests are in flight; the rest queue.
//
// Office.js-free and side-effect-free, so it unit-tests with plain functions.

export const DEFAULT_MAX_CONCURRENT = 4;
export const DEFAULT_MAX_ENTRIES = 500;

export function createLimitedCache({
  maxConcurrent = DEFAULT_MAX_CONCURRENT,
  maxEntries = DEFAULT_MAX_ENTRIES,
} = {}) {
  const resolved = new Map(); // key -> value
  const inFlight = new Map(); // key -> Promise
  const queue = []; // pending starters, drained as slots free up
  let active = 0;

  function remember(key, value) {
    // Map preserves insertion order, so the first key is the oldest.
    if (resolved.size >= maxEntries) {
      const oldest = resolved.keys().next();
      if (!oldest.done) resolved.delete(oldest.value);
    }
    resolved.set(key, value);
  }

  function pump() {
    while (active < maxConcurrent && queue.length > 0) {
      queue.shift()();
    }
  }

  /**
   * Run `task` for `key`, reusing a cached or in-flight result when possible.
   *
   * @param {string} key stable identity of the request
   * @param {() => Promise<any>} task performs the actual call
   * @returns {Promise<any>}
   */
  function run(key, task) {
    if (resolved.has(key)) return Promise.resolve(resolved.get(key));
    const pending = inFlight.get(key);
    if (pending) return pending;

    const promise = new Promise((resolve, reject) => {
      const start = () => {
        active++;
        Promise.resolve()
          .then(task)
          .then(
            (value) => {
              remember(key, value);
              resolve(value);
            },
            // Failures are deliberately NOT cached: a transient provider error
            // would otherwise pin a permanent error onto every cell using that
            // value until the workbook reloads.
            reject
          )
          .finally(() => {
            active--;
            inFlight.delete(key);
            pump();
          });
      };
      if (active < maxConcurrent) start();
      else queue.push(start);
    });

    inFlight.set(key, promise);
    return promise;
  }

  return {
    run,
    clear() {
      resolved.clear();
    },
    get stats() {
      return { cached: resolved.size, inFlight: inFlight.size, queued: queue.length, active };
    },
  };
}
