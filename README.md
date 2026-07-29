# Hermes for Office

Microsoft Word + Excel add-ins that call any OpenAI-compatible LLM API directly. No Hermes server, account, or backend required.

Supported providers include OpenAI, OpenRouter, Azure OpenAI, Together, Ollama, LM Studio — anything exposing `/v1/chat/completions`.

```
                    ┌─ /word/*, /excel/*, /config.json   (static files)
Office add-in ─HTTPS─┤
        :8643        └─ /v1/*  ──HTTPS──▶  provider /v1/chat/completions
```

Caddy serves the bundles and `config.json`, and — by default — forwards `/v1/*` to your provider. Setup asks about this; see [Routing](#routing) for why it is the default and how to turn it off.

Caddy never holds or injects an API key. The add-in reads `config.json` and sends its own `Authorization` header; Caddy only forwards. A caller without a key still gets the provider's 401.

## Apps

| App | Status | Folder |
|-----|--------|--------|
| Word — task-pane chat + AI document editing | ✅ available | [`word/`](./word) |
| Excel — task-pane chat + `=HERMES.*` custom functions | ✅ available | [`excel/`](./excel) |

Each app is a self-contained Office.js add-in with its own `package.json`, webpack build, and manifest. They share client and UI code in [`shared/`](./shared), imported through relative paths — there is no npm workspace between them and nothing is published to a registry.

## Install

Requirements: Microsoft 365 desktop, Node.js 18+, and [Caddy](https://caddyserver.com/docs/install).

```bash
# macOS
bash scripts/install.sh
# Windows PowerShell
.\scripts\install.ps1
```

The script installs dependencies for the root and both add-ins, builds them, then prompts for your provider base URL, API key, model id, and whether to route through the local server (see [Routing](#routing) — say yes unless you know your provider supports browser CORS). It writes two git-ignored files:

- `config.json` (mode `600`) at the repo root — read by both task panes
- `Caddyfile`, created from `Caddyfile.example` if absent, with the provider proxy block filled in

Then start the server and sideload:

```bash
npm run serve
```

In Word and Excel: **Home → Get Add-ins → Manage My Add-ins → Upload My Add-in**, and pick `word/dist/manifest.xml` in Word, `excel/dist/manifest.xml` in Excel. Keep `npm run serve` running while using the add-ins.

### Provider config

`config.json` schema:

```json
{
  "name": "OpenAI",
  "baseUrl": "https://localhost:8643/v1",
  "upstreamUrl": "https://api.openai.com/v1",
  "apiKey": "sk-...",
  "model": "gpt-4o-mini"
}
```

`baseUrl` is what the add-in calls. With the local proxy on it is the Caddy address and `upstreamUrl` records the real provider (used by setup to regenerate the `Caddyfile` block); routing direct, `baseUrl` is the provider itself and there is no `upstreamUrl`.

`baseUrl` must be `https://` — setup rejects anything else. Enter `model` exactly as your provider names it (`gpt-4o-mini`, `anthropic/claude-sonnet-4`, `meta-llama/Llama-3-70b`); it is used verbatim.

Re-run `npm run setup` any time to change providers. The task pane caches `config.json` for the session, so reload the pane after editing it.

## Uninstall

```bash
npm run stop          # stop Caddy
rm config.json        # remove provider credentials
```

Remove the manifests from Office via **Home → Get Add-ins → Manage My Add-ins → Delete**.

## Repo-root scripts

```bash
npm run install:all   # install deps for root + word + excel
npm run build         # build word/ then excel/
npm run setup         # configure provider, create Caddyfile
npm run serve         # serve built add-ins over local HTTPS
npm run stop          # stop Caddy
npm test              # node:test suite (shared/ + scripts/)
npm run lint          # lint shared/ + scripts/ + word/ + excel/
```

Each add-in also has `npm run dev-server`, `build:dev`, `watch`, and `validate` for iterating without a full production build.

## How it works

1. The task pane reads the current selection (or table, or whole document/sheet) plus your typed prompt.
2. `askHermes()` in [`shared/hermes.js`](./shared/hermes.js) loads `/config.json`, then POSTs the conversation to `${baseUrl}/chat/completions` with a 60s timeout and one retry on timeout/network failure. HTTP errors are surfaced, not retried.
3. The reply is parsed into a structured edit proposal (`parseEdits` / `parseTableChanges` in [`shared/parsers.js`](./shared/parsers.js)) and rendered as a proposal card — nothing is written until you click **Apply**.
4. On Apply, edits go through Office.js (`Word.run` / `Excel.run`). Both panes offer a toggle to visually mark what the AI changed: red text in Word, tinted fill in Excel.

Both panes bound what they send: Word truncates full-document prompts at 120k characters, Excel caps the sheet snapshot at 500 rows / 100 columns / 200KB, and both keep the replayed history to 20 turns.

### Excel custom functions

`=HERMES.CLASSIFY`, `=HERMES.EXTRACT`, `=HERMES.SUMMARIZE`, and `=HERMES.FORMULA_HELP` call the provider straight from a cell, independent of the task pane.

They are drag-fillable, so all four go through the gate in [`shared/fn-cache.js`](./shared/fn-cache.js): resolved answers are memoised (500 entries), identical concurrent requests share one call, and at most 4 requests are in flight at once. Filling a column of repeated values costs one request per distinct value, not one per cell. Errors are not cached, so a transient provider failure doesn't stick.

Each call is still a billable request. A column of 500 distinct values is 500 calls.

## Security

### API key

The key sits in plaintext in `config.json` at the repo root, because an Office WebView cannot read arbitrary local files. Setup writes it mode `600` and it is git-ignored. It is deliberately **not** copied into `word/dist/` or `excel/dist/` — those are the folders you open when sideloading, and the easiest place to leak a key from by zipping one up.

The add-in sends your document content and your API key directly to the configured provider. Use an HTTPS provider URL (enforced) and a key scoped as narrowly as your provider allows.

### What Caddy does and does not protect

Caddy serves static files over local HTTPS and, when routing through the local proxy, forwards `/v1/*` to your provider. That is all it does. It authenticates nobody, injects no credentials, and sets no CORS header — the last one deliberately:

> The task panes are served from `https://localhost:8643/`, so their `fetch("/config.json")` is same-origin and needs no CORS. An `Access-Control-Allow-Origin: *` header on this server would let **any page open in any browser on the machine** read `config.json`, API key included.

If you edit `Caddyfile`, do not add one back. Run Caddy only while you're actually using the add-ins; `npm run stop` when you're done.

### Routing

Setup offers two ways to reach the provider.

**Through the local server (default).** `config.json` gets `baseUrl: https://localhost:8643/v1`, and setup writes a matching proxy block into `Caddyfile` between its `# >>> hermes-proxy >>>` markers. The provider call then leaves from the *same origin* as the task pane, so the browser never sends a CORS preflight — which is the single most common reason a self-hosted provider fails here. Re-running setup rewrites that block, so switching providers stays a one-command change; edit it there, not by hand.

**Direct.** `baseUrl` points straight at the provider. One less hop, but the provider must satisfy the CORS rules below. OpenAI and OpenRouter do; most self-hosted routers do not.

Proxying does not widen the key's exposure: Caddy injects nothing, the add-in still supplies its own `Authorization`, and the key already sits in a local file that any local process can read either way.

### Provider CORS

Only relevant when routing **direct**. The task pane calls your provider from a browser (the Office WebView), so the provider must allow cross-origin requests. Many self-hosted routers don't, and the symptom is confusing: `fetch` rejects instantly with a bare `TypeError` carrying no detail.

The request sends `Content-Type: application/json` and `Authorization`, which makes it a *non-simple* request — so the browser first sends a **preflight** `OPTIONS`. Two rules decide whether it works:

1. **`OPTIONS` must succeed without an API key.** Browsers never attach `Authorization` to a preflight; that is the spec, not a bug. A router that requires auth on `OPTIONS` returns 401 and the real request is never sent.
2. **Both the preflight and the actual response need `Access-Control-Allow-Origin`.**

Diagnose from a shell — this is exactly what the browser does:

```bash
curl -i -X OPTIONS https://YOUR-PROVIDER/v1/chat/completions \
  -H "Origin: https://localhost:8643" \
  -H "Access-Control-Request-Method: POST" \
  -H "Access-Control-Request-Headers: content-type,authorization"
```

You need `204` (or `200`) plus `access-control-allow-*` headers. A `401` here is the bug.

For nginx, answer the preflight *before* any auth check:

```nginx
location /v1/ {
    if ($request_method = OPTIONS) {
        add_header Access-Control-Allow-Origin  "https://localhost:8643" always;
        add_header Access-Control-Allow-Methods "GET, POST, OPTIONS"      always;
        add_header Access-Control-Allow-Headers "Content-Type, Authorization, Idempotency-Key" always;
        add_header Access-Control-Max-Age       86400                     always;
        add_header Content-Length 0;
        return 204;
    }

    # ... existing auth + proxy_pass ...

    add_header Access-Control-Allow-Origin "https://localhost:8643" always;
}
```

Providers that already work from a browser include OpenAI and OpenRouter. If you would rather not touch the provider at all, re-run `npm run setup` and accept the local-proxy default.

### Non-standard response bodies

Some OpenAI-compatible routers answer a *non-streaming* request with `Content-Type: text/event-stream`, whitespace padding (to defeat proxy buffering), and a trailing `data: [DONE]` sentinel — none of which `res.json()` can parse. `parseCompletionBody` in [`shared/hermes.js`](./shared/hermes.js) handles the padded, SSE-framed, and streamed-delta shapes, and surfaces a structured `{"error": …}` body using the provider's own wording instead of a generic parse failure.

### Prompt injection

Document and workbook content is untrusted input: text inside a file you open can influence what the model proposes. Two mitigations, both load-bearing:

- **Nothing is written without your approval.** Every change is rendered as a proposal card and applied only when you click Apply. Read the card — it is the security boundary.
- **Excel never applies a model-proposed value as a live formula.** Strings starting with `=`, `+`, `-`, or `@` are forced to literal text (`literalCellValue` in `excel/src/taskpane/taskpane.js`), so an injected `=WEBSERVICE(...)` cannot become an exfiltration channel on Apply.

Apply is also guarded against stale proposals: Excel refuses a cell whose current value no longer matches what the proposal was built from and refuses to write to a different sheet, and Word re-finds the proposal's table by content fingerprint rather than writing into whichever table happens to be selected.

## Project layout

```
hermes-office/
├── word/                  # Word add-in (Office.js, webpack)
│   ├── src/taskpane/       # chat UI + document editing logic
│   ├── src/shared/         # re-exports ../../shared/hermes.js
│   ├── manifest.xml
│   └── package.json
├── excel/                 # Excel add-in (Office.js, webpack)
│   ├── src/taskpane/       # chat UI + sheet editing logic
│   ├── src/functions/      # =HERMES.* custom functions
│   ├── src/shared/         # re-exports ../../shared/hermes.js
│   ├── manifest.xml
│   └── package.json
├── shared/                # imported by both apps via relative paths
│   ├── hermes.js           # provider client (timeout + retry + config cache)
│   ├── parsers.js          # column conversion, edit/table-edit parsing
│   ├── fn-cache.js         # cache + concurrency gate for custom functions
│   ├── proposal-card.js    # shared proposal / toast / status UI
│   └── design-system.css   # shared styles, copied into both dists at build
├── scripts/               # install, setup, serve, stop
├── config.example.json
├── Caddyfile.example      # copied to Caddyfile (git-ignored) by setup
└── eslint.config.mjs      # lint config for shared/ and scripts/
```

## Notes

No native signed MSIX/pkg installer is included — platform signing requires publisher certificates. The package ships local sideload scripts only.

## License

MIT — see [LICENSE](./LICENSE).
