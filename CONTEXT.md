# Hermes Office

Word and Excel add-ins that let a user chat with an LLM about the open document
or workbook, and apply the model's suggested changes after reviewing them.

## Language

### Runtime pieces

**Provider**:
The upstream OpenAI-compatible API that answers chat completions. Only the
Caddyfile knows its address.
_Avoid_: backend, Hermes server, API Server, agent

**Local Gateway**:
The Caddy instance on `localhost:8643`. It serves the built bundles and
`config.json`, and forwards `/v1/*` to the Provider with the API key attached.
_Avoid_: proxy server, backend, Hermes server

**Task Pane**:
The add-in's chat UI hosted inside Word or Excel. One per app.
_Avoid_: sidebar, panel, plugin

**Custom Function**:
An Excel worksheet formula (`=HERMES.*`) that reaches the Provider from a cell,
independently of the Task Pane — and therefore outside Apply. A workbook that
merely contains one fires it on open, so this is the one path where document
content alone can reach the Provider. Off unless explicitly enabled; see
ADR-0003.

### Reaching the Provider

**Local Hop**:
The leg from the Task Pane to the Local Gateway — `config.json` and
`/v1/chat/completions`, both same-origin. It crosses no network the user's
router can see, so it fails only when the Gateway is not running or not
configured.
_Avoid_: the connection, the request, calling the API

**Upstream Hop**:
The leg from the Local Gateway to the Provider. This is the one that crosses the
internet, so it is the leg a change of network can break — and the leg no
observer outside the Gateway can see. Whether the Gateway reached the Provider
at all is knowledge only the Gateway has; it must say so rather than let the
Task Pane guess from a status code.
_Avoid_: the backend call, upstream, the proxy request

**Disclosed Failure**:
What the user is told when a call fails: one of a fixed, named set, each naming
the hop that broke and the one thing the user can do about it. The set is closed
by construction and carries no text from the Provider — not its response body,
not a status line, not a stack. Two situations that call for the same user
action are the same Disclosed Failure; two that call for different actions are
never merged into one. See ADR-0005.
_Avoid_: error message, error string, the error

### The edit cycle

**Snapshot**:
The bounded slice of document or sheet content sent to the Provider as context.
Bounded because a whole workbook does not fit in a request.
_Avoid_: context, dump, payload

**Proposal**:
A parsed set of changes the model suggested, rendered as a card. Until Apply it
has touched nothing — see Partial Apply for what it becomes after a press that
ran only some of it.
_Avoid_: suggestion, edit plan, diff

**Action**:
One change inside a Proposal — a cell write, a format, a new sheet, a paragraph
replacement.
_Avoid_: edit, operation, command

**Apply**:
The user pressing the button that turns a Proposal into real document changes.
This is the security boundary: document content is untrusted input, so nothing
the Task Pane writes reaches the document without passing through it. The one
path outside it is a Custom Function, which is why that is off by default.

A boundary is only worth as much as the review it enables, so an Action that
Apply would write must be legible on the card before the press — its payload,
and how many places it touches — not summarised as a count.

A Proposal is reviewable only when every Action, its payload, and the number of
affected locations can be read before Apply. Oversized Proposals are rejected
before rendering; they are not truncated or summarized into an approvable card.
_Avoid_: commit, save, execute

**Partial Apply**:
An Apply that ran some of a Proposal's Actions and not others. The Proposal
survives with the unrun Actions still on the card, so a single bad Action does
not discard the rest. But the document has moved underneath those Actions, and
the review the user already gave no longer describes what an Action that writes
blind would write now. The card must say so before the second press — a boundary
the user cannot see is not a boundary.
_Avoid_: partial failure, retry, resume

**Pin**:
A Word bookmark anchoring the passage a Proposal targets, so the target survives
focus leaving the document. The bookmark, not the cached text, is authoritative.
_Avoid_: selection, anchor, marker
