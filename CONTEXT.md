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
independently of the Task Pane.

### The edit cycle

**Snapshot**:
The bounded slice of document or sheet content sent to the Provider as context.
Bounded because a whole workbook does not fit in a request.
_Avoid_: context, dump, payload

**Proposal**:
A parsed set of changes the model suggested, rendered as a card. It has touched
nothing yet.
_Avoid_: suggestion, edit plan, diff

**Action**:
One change inside a Proposal — a cell write, a format, a new sheet, a paragraph
replacement.
_Avoid_: edit, operation, command

**Apply**:
The user pressing the button that turns a Proposal into real document changes.
This is the security boundary: document content is untrusted input, so nothing
reaches the document without passing through it.
_Avoid_: commit, save, execute

**Pin**:
A Word bookmark anchoring the passage a Proposal targets, so the target survives
focus leaving the document. The bookmark, not the cached text, is authoritative.
_Avoid_: selection, anchor, marker
