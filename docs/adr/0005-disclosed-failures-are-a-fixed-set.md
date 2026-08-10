# A failed call is disclosed as one of a fixed set, never as the Provider's own words

An earlier fix collapsed every failure in the call pipeline to a single fixed
sentence, because upstream response bodies can carry the Provider's diagnostics,
account details, and stack traces, and those were reaching the Word chat log
verbatim. That closed the leak but bought it with the whole diagnosis: a Gateway
that is not running, a rejected API key, a wrong model id, an exhausted quota,
and a genuinely broken network all render as the same sentence, so the one thing
the user needs — which hop broke — is exactly what is withheld. Instead, the
Task Pane classifies the failure itself into a closed set of **Disclosed
Failures**, each naming the hop and the single action that resolves it, each
written by us and containing no byte of upstream text. Whether the **Upstream
Hop** failed at all is knowledge only the Local Gateway has, so Caddy marks its
own `handle_errors` responses rather than leaving the Task Pane to guess a bare
`502` that the Provider could equally have sent itself.

## Considered options

- **Show the upstream error.** Rejected: it is the leak this replaces.
- **Keep the single sentence.** Rejected: it is what sent a user hunting through
  wifi settings for what was a mistyped model id.
- **Let the Task Pane infer everything from the HTTP status.** Rejected for one
  case that happens to be the case that motivated this: a `502` raised by Caddy
  because it could not reach the Provider is indistinguishable from a `502` the
  Provider returned itself, and only the first means "your network changed".

## Consequences

- The Provider's actual words never appear in the Task Pane. Reading them means
  opening the console (the full error is still logged there) or the Gateway log.
  This is the accepted cost; do not "help" by appending the upstream message.
- The set is closed. Adding a case means adding a named Disclosed Failure with
  its own user action and its own test — not widening an existing one into
  something vaguer. Two cases sharing one user action must share one Disclosed
  Failure instead of multiplying.
- Classification is not UI. It lives in a DOM-free module both add-ins import,
  because the previous home for it was a card-rendering module and that is why
  Excel silently never got the fix.
- Retrying is a separate decision and deliberately not made here: with no
  evidence that the Provider honours `Idempotency-Key`, an automatic retry on a
  dead Upstream Hop risks paying for two completions, since the first request may
  well have reached the Provider before the return path died. The Disclosed
  Failure asks the user to press again. Flipping this later should be a change in
  one place.
