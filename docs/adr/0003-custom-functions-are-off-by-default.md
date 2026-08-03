# Custom Functions are off unless config.json opts in

`=HERMES.*` worksheet functions reach the Provider from a cell, outside the Apply
boundary that governs everything else the add-in writes. Because a formula is
stored *in the workbook*, a file the user merely opens can fire them: opening an
untrusted `.xlsx` was enough to ship its own contents to the user's Provider and
spend their quota, with nothing shown in the Task Pane and no button pressed.
They are therefore gated on `"customFunctions": true` in `config.json` — a local
file a document cannot write — and default to off.

## Considered options

- **Leave them on.** Rejected: a gate that a hostile document can satisfy by
  existing is not a gate.
- **Remove them.** Rejected: they are genuinely useful on a sheet the user
  trusts, which is the normal case.
- **Prompt per call.** Rejected: a recalculation can fire hundreds at once, so
  the prompt would be trained away within a minute.

## Consequences

- Existing users must add the flag before `=HERMES.*` works again. Without it
  the cell shows the error text telling them exactly what to add.
- `npm run setup` preserves the flag rather than rewriting `config.json` from
  scratch — reconfiguring a Provider must not silently re-arm this.
- The gate bounds *automatic* execution, not what an enabled function can be
  asked to do. Once on, any formula in any opened workbook can call out again.
