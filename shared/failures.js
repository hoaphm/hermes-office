// Disclosed Failures — the only vocabulary in which a failed call is described
// to the user. See CONTEXT.md and ADR-0005.
//
// Two rules the rest of this file exists to enforce:
//
//   1. Nothing from the Provider is disclosed. Not its response body, not its
//      status line, not a stack, not a local path or URL. Every string below is
//      written here and contains no interpolation.
//   2. The set is closed, and two situations calling for the same user action
//      are the same Disclosed Failure. Adding a case means adding a named one
//      with its own action and its own test — not widening an existing one.
//
// This module holds no DOM. It used to live in proposal-card.js, a module that
// renders cards, which is why the Excel task pane never got the fix and went on
// printing upstream response bodies into its chat log.

/**
 * @typedef {object} Failure
 * @property {string} code   stable identifier, for tests and logs
 * @property {"local"|"upstream"} hop  which hop broke — see CONTEXT.md
 * @property {string} message  the only text ever shown to the user
 */

/** @type {Record<string, Failure>} */
export const FAILURES = {
  // ---- Local Hop: Task Pane → Local Gateway --------------------------------
  GATEWAY_DOWN: {
    code: "GATEWAY_DOWN",
    hop: "local",
    message: "Local Gateway chưa chạy. Mở terminal và chạy: npm run serve",
  },
  NOT_CONFIGURED: {
    code: "NOT_CONFIGURED",
    hop: "local",
    message: "Local Gateway chưa được cấu hình. Chạy: npm run setup",
  },

  // ---- Upstream Hop: Local Gateway → Provider ------------------------------
  // The Gateway said it could not reach the Provider at all. This is what a
  // change of network looks like from inside the task pane.
  UPSTREAM_UNREACHABLE: {
    code: "UPSTREAM_UNREACHABLE",
    hop: "upstream",
    message:
      "Local Gateway không kết nối được tới Provider. Kiểm tra mạng, rồi bấm Hỏi lại.",
  },
  KEY_REJECTED: {
    code: "KEY_REJECTED",
    hop: "upstream",
    message: "Provider từ chối khóa API. Chạy lại: npm run setup",
  },
  // Kept separate from KEY_REJECTED on purpose: a mistyped model id sends the
  // user hunting through network settings if it is disclosed as a connection
  // problem. Model ids are taken verbatim, so this is a routine mistake.
  MODEL_REJECTED: {
    code: "MODEL_REJECTED",
    hop: "upstream",
    message:
      "Provider không nhận tên model đang cấu hình. Kiểm tra lại tên model: npm run setup",
  },
  OVERLOADED: {
    code: "OVERLOADED",
    hop: "upstream",
    message: "Provider đang quá tải hoặc đã hết quota. Chờ một lát rồi thử lại.",
  },
  UPSTREAM_ERROR: {
    code: "UPSTREAM_ERROR",
    hop: "upstream",
    message: "Provider gặp lỗi nội bộ. Báo cho người vận hành Provider.",
  },
  UPSTREAM_SLOW: {
    code: "UPSTREAM_SLOW",
    hop: "upstream",
    message: "Provider không phản hồi kịp thời. Bấm Hỏi lại.",
  },
  BAD_PAYLOAD: {
    code: "BAD_PAYLOAD",
    hop: "upstream",
    message:
      "Provider trả về dữ liệu không đọc được. Thử đổi model hoặc kiểm tra Provider.",
  },

  // Last resort for a call-pipeline failure that matched none of the above. It
  // names no cause because we do not know one; it must never be widened to
  // absorb a case that has its own action.
  UNKNOWN: {
    code: "UNKNOWN",
    hop: "upstream",
    message: "Gọi Provider thất bại. Bấm Hỏi lại.",
  },
};

// A failure carries `err.hermes` only if it came from the call pipeline in
// shared/hermes.js. The tags are the fallback for an error that crossed a
// boundary which dropped the property.
const PIPELINE_TAG =
  /\[(?:call-api:|api-(?:first|retry)\]|stage:call-api\]|config-fetch:|config-url\])/;

const isPipelineError = (err) =>
  Boolean(err && err.hermes) || PIPELINE_TAG.test(String((err && err.message) || err));

/**
 * Which Disclosed Failure a caught error amounts to, or `null` when the error
 * did not come from the call pipeline at all — a local task-pane error
 * (validation, staleness, an oversized Proposal), whose own message is already
 * written for the user and is shown verbatim.
 * @param {unknown} err
 * @returns {Failure|null}
 */
export function classifyFailure(err) {
  if (!isPipelineError(err)) return null;

  const detail = (err && err.hermes) || {};
  const { stage, status, gateway, badPayload } = detail;
  const name = detail.name || (err && err.name);

  if (name === "TimeoutError" || name === "AbortError") return FAILURES.UPSTREAM_SLOW;
  if (stage === "config-model" || stage === "config-url") return FAILURES.NOT_CONFIGURED;

  // Nothing answered. On the Local Hop that means nothing is listening on the
  // Gateway's port, which is the single most common way this fails.
  if (stage === "config-fetch") {
    return status == null ? FAILURES.GATEWAY_DOWN : FAILURES.NOT_CONFIGURED;
  }
  if (stage === "fetch" || (stage === "load-config" && status == null)) {
    return FAILURES.GATEWAY_DOWN;
  }

  if (badPayload) return FAILURES.BAD_PAYLOAD;

  // The Gateway marks responses it generated itself. A 502 it produced means it
  // never reached the Provider; the same 502 without the mark came back from
  // the Provider, which is a different failure with a different action.
  if (gateway) {
    return status >= 500 || status == null
      ? FAILURES.UPSTREAM_UNREACHABLE
      : FAILURES.UNKNOWN;
  }

  if (status === 401 || status === 403) return FAILURES.KEY_REJECTED;
  if (status === 400 || status === 404) return FAILURES.MODEL_REJECTED;
  if (status === 429) return FAILURES.OVERLOADED;
  if (status >= 500) return FAILURES.UPSTREAM_ERROR;

  return FAILURES.UNKNOWN;
}

/**
 * The only user-visible representation of an error. A call-pipeline failure
 * becomes its Disclosed Failure message; every other error — raised locally by
 * the task pane, and already phrased for the user — passes through verbatim.
 * The full error, stack included, is the caller's to send to console.error and
 * belongs nowhere else.
 * @param {unknown} err
 * @returns {string}
 */
export function userErrorMessage(err) {
  const failure = classifyFailure(err);
  if (failure) return failure.message;
  return (err && err.message) || String(err);
}

/**
 * Render the result of `checkHops()` as one line per hop. Named in the
 * glossary's terms rather than in UI words, because the whole point of the
 * check is to tell the user which hop to go look at.
 * @param {{local: {ok: boolean, error: unknown}, upstream: {ok: boolean, error: unknown, skipped?: boolean}}} result
 * @returns {string[]} exactly two lines
 */
export function hopReport(result) {
  const local = result.local.ok
    ? "Local Hop: OK"
    : `Local Hop: ${userErrorMessage(result.local.error)}`;
  const upstream = result.upstream.skipped
    ? "Upstream Hop: chưa kiểm tra — sửa Local Hop trước."
    : result.upstream.ok
      ? "Upstream Hop: OK"
      : `Upstream Hop: ${userErrorMessage(result.upstream.error)}`;
  return [local, upstream];
}

/**
 * True when `err` came from the Provider call pipeline rather than from local
 * task-pane logic.
 * @param {unknown} err
 * @returns {boolean}
 */
export function isProviderError(err) {
  return classifyFailure(err) !== null;
}
