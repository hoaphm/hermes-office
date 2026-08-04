# A partially applied Proposal survives, but the card must say it is stale

Applying a Proposal is not atomic: the Excel task pane syncs each Action
separately so one bad range cannot abort the batch, which means a press can
leave some Actions written and others not. The unrun Actions stay on the card so
the work is not lost. That is a deliberate deviation from ADR-0001..0003, which
all treat Apply as a single reviewed step, and it costs something real — the
document has changed since the user reviewed those Actions, so the card must
mark what is no longer covered by that review before the second press.

## Considered options

- **Treat a touched Proposal as dead** — report what ran and clear the card, as
  the Word pane does. Rejected: a Partial Apply is usually caused by one bad
  Action among many, and discarding the other nineteen makes the user re-ask
  from scratch for someone else's mistake.
- **Make Apply all-or-nothing** — roll back on the first failure. Rejected:
  Office.js has no transaction. We would have to synthesise an undo by
  re-writing captured old values, which is more write paths, more ways to
  corrupt the document, and no better than what it replaces.
- **Keep the Remainder silently** — what the code did before. Rejected: the
  second press looked identical to the first while standing on a review that no
  longer held. See below.

## Consequences

- The staleness guard is re-baselined against the sheet *after* the writes, so
  the retry does not refuse itself. This deliberately spends the guarantee the
  guard exists to provide, and the card carries the cost instead: a banner
  saying the sheet moved, and a badge on each Action that writes blind.
- Only `setCell` carries a per-Action old-value guard, so only `setCell` is
  still covered by the original review. `setCells`, `format`, `createTable` and
  `createChart` write blind and are the ones badged.
- Re-baselining reads the *proposal's* sheet by name, never the active sheet. A
  `newSheet` Action activates the sheet it creates, so reading the active sheet
  silently retargeted the Remainder at the new sheet.
- When a batch created or renamed a sheet, there is no Remainder at all: the
  unqualified ranges in the unrun Actions no longer have a well-defined target,
  and refusing is the only honest answer. The user is told to ask again.
- The Word pane still discards the whole Proposal after a partial
  `applyFulldocEdits`. That divergence is known and left alone here — closing it
  means giving Word a Remainder UI, which is a feature, not a stability fix.
