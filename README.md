# Hermes for Office

Microsoft Word + Excel add-ins that call any OpenAI-compatible LLM API directly. No Hermes server, account, or backend required.

Supported providers include OpenAI, OpenRouter, Azure OpenAI, and any provider exposing `/v1/chat/completions`.

```
Office add-in ──HTTPS──▶ Caddy (:8643, local static server) ──HTTPS──▶ provider /v1/chat/completions
```

The add-in reads provider config from local `/config.json`, then calls provider directly. Caddy only serves bundles and config; it does not proxy or inject API keys.

`config.json` is git-ignored. Never share it or commit it.

## Apps

| App | Status | Folder |
|-----|--------|--------|
| Word — task-pane chat + AI document editing | ✅ available | [`word/`](./word) |
| Excel — task-pane chat + `=HERMES.*` custom functions | ✅ available | [`excel/`](./excel) |

Each app is a self-contained Office.js add-in with its own `package.json`, webpack build, and manifest. They share provider client code in [`shared/`](./shared).

## Install on Windows or macOS

Requirements: Microsoft 365 desktop, Node.js 18+, and Caddy (HTTPS static server).

1. Clone or extract project.
2. Run setup:

```bash
# macOS
bash scripts/install.sh
# Windows PowerShell
.\\scripts\\install.ps1
```

3. Enter provider base URL, API key, and model when prompted. Setup writes `config.json` to project root and both `dist/` folders.
4. Start server:

```bash
npm run serve
```

5. Sideload both manifests in Word/Excel: **Home → Get Add-ins → Manage My Add-ins → Upload My Add-in → Upload a custom add-in → manifest.xml**.

Use `word/dist/manifest.xml` in Word and `excel/dist/manifest.xml` in Excel. Keep `npm run serve` running while using add-ins.

Provider config schema:

```json
{
  "name": "OpenAI",
  "baseUrl": "https://api.openai.com/v1",
  "apiKey": "sk-...",
  "model": "gpt-4o-mini"
}
```

`apiKey` is stored in plaintext local `config.json` because Office WebView cannot read arbitrary files. Protect this file and never commit it. `config.json` is git-ignored.
## Uninstall

Stop server with `npm run stop`. Remove manifests from Office via **Home → Get Add-ins → Manage My Add-ins → Delete**. Delete `config.json` to remove provider credentials.

For development, use each app's existing `npm run dev-server` and `npm start` scripts.

## Repo-root convenience scripts

```bash
npm run build   # builds word/ then excel/
npm run setup   # configure provider
npm run serve   # serve built add-ins over local HTTPS
npm test        # shared node:test suite
npm run lint    # lint both add-ins
```

Legacy `npm start` flows are development-only and still require their original dev proxy setup.

## How it works

1. Task pane reads document/workbook content and typed prompt.
2. `shared/hermes.js` loads `/config.json`, then POSTs conversation to `${baseUrl}/chat/completions` with configured model and bearer key.
3. Reply is parsed into structured edit proposal and shown for approval.
4. Office.js applies edits only after user clicks **Apply**.

Excel exposes `=HERMES.CLASSIFY`, `=HERMES.EXTRACT`, `=HERMES.SUMMARIZE`, and `=HERMES.FORMULA_HELP` custom functions.

## Security

API key lives in plaintext local `config.json` because Office WebView cannot read arbitrary local files. Protect the file, do not share it, and never commit it. Use HTTPS provider URLs. Add-in sends document content and API key directly to configured provider.

### Prompt injection

Document/workbook content is untrusted input. Nothing is written without user approval; Excel forces model-proposed formula-like values to literal text.

## License

MIT — see [LICENSE](./LICENSE).

## Project layout

See `word/`, `excel/`, `shared/`, `scripts/`, `config.example.json`, and `Caddyfile`.

There is no npm workspace between add-ins; shared helpers are imported through relative paths.

## Notes

Native signed MSIX/pkg installers are not included. Package contains local sideload scripts; platform signing requires publisher certificates.
### Repo-root convenience scripts

From the repo root, `package.json` wraps both apps:
```
npm run build   # builds word/ then excel/
npm test        # runs the shared node:test suite (shared/*.test.js)
npm run lint    # lints word/ then excel/ (office-addin-lint)
```

## How it works

1. The task pane reads the current selection (or table, or whole document/sheet) plus your typed prompt.
2. It sends the conversation to Hermes via `askHermes()` (in [`shared/hermes.js`](./shared/hermes.js)), which POSTs to `https://localhost:8643/v1/chat/completions` with a 60s timeout and one retry on timeout/network failure.
3. Hermes' reply is parsed for a structured edit proposal (`parseEdits` / `parseTableChanges` in [`shared/parsers.js`](./shared/parsers.js)) and shown as a preview — nothing is written to the document/sheet until you click **Apply**.
4. On Apply, the add-in applies the edits via Office.js (`Word.run` / `Excel.run`). Word can optionally mark applied edits in red so you can spot AI changes at a glance.

Excel additionally exposes read-only, cacheable, drag-fillable custom functions — `=HERMES.CLASSIFY`, `=HERMES.EXTRACT`, `=HERMES.SUMMARIZE`, `=HERMES.FORMULA_HELP` — that call the agent directly from a cell, independent of the task pane.

## Project layout

```
hermes-office/
├── word/                 # Hermes for Word add-in (Office.js, webpack)
│   ├── src/
│   │   ├── taskpane/      # chat UI + document editing logic
│   │   ├── commands/      # ribbon command functions
│   │   └── shared/        # re-exports ../../shared/hermes.js
│   ├── manifest.xml
│   └── package.json
├── excel/                # Hermes for Excel add-in (Office.js, webpack)
│   ├── src/
│   │   ├── taskpane/      # chat UI + sheet editing logic
│   │   ├── functions/     # =HERMES.* custom functions
│   │   ├── commands/       # ribbon command functions
│   │   └── shared/         # re-exports ../../shared/hermes.js
│   ├── manifest.xml
│   └── package.json
├── shared/                # helpers imported by both apps via relative paths
│   ├── hermes.js           # askHermes client (timeout + retry)
│   └── parsers.js          # column conversion, edit/table-edit parsing, etc.
├── Caddyfile.example      # copy to Caddyfile (git-ignored) and fill in your key
└── package.json           # repo-root build/test/lint wrapper scripts
```

There is no npm workspace between the two add-ins (nothing is published to a registry); `shared/` is deduped purely via relative imports (`word/src/shared/hermes.js` and `excel/src/shared/hermes.js` both do `export * from "../../../shared/hermes.js"`).

## Security

⚠️ The API Server gives the agent's **full toolset, including terminal commands**. Treat `API_SERVER_KEY` like a password:
- Long and random; **never commit** your real `Caddyfile` or `.env` (both are git-ignored).
- Bind everything to `localhost`; keep `API_SERVER_CORS_ORIGINS` narrow.
- No add-in ever holds the key — Caddy injects it.

### What the Caddy hop does and does not protect

Keeping the key out of the add-in bundle solves a *distribution* problem: the secret never lands in git, in `dist/`, or in a manifest you might share. It is **not** an access control.

Caddy injects the `Authorization` header for **every** request that reaches `:8643/v1/*`, without authenticating the caller. So on a machine running this setup, any local process — another app, a shell one-liner, an npm `postinstall` script — can `curl https://localhost:8643/v1/chat/completions` and drive the full agent, tools and all, without ever seeing the key. The browser same-origin policy stops *other web pages* from reading responses (the CORS snippet only allows `https://localhost:8643`), but it stops nothing running outside a browser.

Practical consequences:
- Run `caddy` only while you are actually using the add-ins; stop it when you're done.
- Treat a machine with this running as one where any local code has agent-level privileges. Don't leave it up on a shared or untrusted host.
- The threat this design defends against is *leaking the key*, not *someone local using the agent*.

### Prompt injection

Both panes send document/workbook content to the agent, so text inside a file you open is untrusted input that can influence what the model proposes. Two mitigations are in place, and both matter:
- **Nothing is written without your approval** — every change is rendered as a proposal card and applied only when you click Apply. Read the card; it is the security boundary.
- **Excel never applies a model-proposed value as a live formula.** Strings starting with `=`, `+`, `-`, or `@` are forced to literal text (`literalCellValue` in `excel/src/taskpane/taskpane.js`), so an injected `=WEBSERVICE(...)` cannot turn into an exfiltration channel on Apply.

## License

MIT — see [LICENSE](./LICENSE).
