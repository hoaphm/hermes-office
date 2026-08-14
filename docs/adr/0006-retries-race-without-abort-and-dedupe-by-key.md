# A timed-out call retries once, raced without AbortController and deduped by Idempotency-Key

A chat turn can fail by timing out — the Local Gateway reached the Provider,
but no answer came back within the window. The tempting fix, abort the first
request and try again, does not exist on this platform: Office for Mac's
WKWebView rejects an `AbortSignal` passed to `fetch()` with "The string did not
match the expected pattern" (both `AbortSignal.timeout()` and
`AbortController.signal`). So `fetchWithTimeout` races the request against a
timer instead and the first request keeps running in the background. Retrying
then means two completions may be billed — the first may still complete
Provider-side after the timer fired, and its answer is silently dropped.

This ADR records how `askHermes` retries a timed-out or network-failed call:
once, no abort, with an `Idempotency-Key` that lets a Provider that honours the
header answer the retry from the first attempt instead of running (and billing)
a second completion.

## The decision

- **Retry once** on timeout or network failure — never on a Provider HTTP
  error (a 4xx/5xx means the Provider answered; retrying it would just re-send
  a request the Provider already judged, and the Disclosed Failure for that
  case asks the user to act, not to wait).
- **Do not abort the first request.** WKWebView forbids the signal; the race
  is the only portable timeout.
- **Every call carries an `Idempotency-Key`.** `askHermes` generates one fresh
  key per logical turn and sends the SAME key on the first attempt and the
  retry. A Provider that honours the header (OpenAI does for chat completions)
  can dedupe the retry against the first attempt and return its result, so the
  user pays once. The key is per-call-fresh, NOT content-derived: a user who
  re-asks the same question must get a new answer, not the cached first one.
- Callers that want content-derived keys (Excel Custom Functions, where a
  recalculation storm fires the same formula many times) pass their own; the
  per-turn default is only a fallback.

## Considered options

- **Abort the first request with AbortController, then retry.** Rejected:
  the platform forbids it. The comment in `fetchWithTimeout` documents the
  exact WKWebView error.
- **Retry without a key.** Rejected (BUG-03): two full completions, first one
  silently discarded — the double-cost this decision exists to prevent.
- **Don't retry at all; surface the Disclosed Failure and let the user press
  again.** Rejected for timeouts/network failures: a transient blip on the
  Upstream Hop is common and a single automatic retry converts most of them
  into successful turns without user action. ADR-0005 deliberately left the
  retry decision out of the failure-classification work; this is that decision,
  made where the retry lives (`askHermes`), not in the classifier.
- **Make the key content-derived.** Rejected: identical messages are also a
  user re-asking the same question, which must produce a fresh completion —
  a content hash would hand them back the first cached answer.

## Consequences

- A Provider that ignores `Idempotency-Key` (not all do) can still bill twice
  when the first attempt actually completed Provider-side after the timer
  fired. This is the accepted residual cost of not being able to abort; the
  retry is limited to ONE so the worst case is two completions, not many.
- The Gateway forwards the header as-is; nothing here authenticates, and the
  key is not a secret — it is a dedup hint only.
- The key is generated inside `askHermes`, so every call path (Word chat,
  Excel chat, Custom Functions) is covered by default; a caller passes its own
  only when content-derived dedup is the point (Custom Functions).
- "Retry once" is a behaviour of `askHermes`, not of `callApi` — the
  single-attempt primitive stays retry-free for the Hop checks and tests.
