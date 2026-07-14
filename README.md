# Hermes for Office

Put your own [Hermes](https://github.com/NousResearch/hermes-agent) agent **inside Microsoft Office** — your models, your skills, your approval gate. Not Copilot.

```
Office add-in ──HTTPS──▶ Caddy (:8643, TLS + injects API key) ──▶ Hermes API Server (:8642) ──▶ your agent
```

The add-in never holds the API key. It calls `https://localhost:8643/v1/chat/completions`; Caddy terminates TLS (Office add-ins must be served over HTTPS) and injects the `Authorization` header before forwarding to the Hermes API Server. This keeps the secret out of the add-in bundle and out of git.

## Apps

| App | Status | Folder |
|-----|--------|--------|
| Word — task-pane chat + AI document editing | ✅ available | [`word/`](./word) |
| Excel — task-pane chat + `=HERMES.*` custom functions | ✅ available | [`excel/`](./excel) |

Each app is a self-contained Office.js add-in with its own `package.json`, webpack build, and manifest. They share the same backend (below) and the same helper code in [`shared/`](./shared).

## Shared backend (set up once, used by every app)

Hermes already exposes the **full agent** (tools, memory, skills) as an OpenAI-compatible HTTP endpoint — the **API Server** (not `hermes proxy`, which is model-only). Each add-in calls `/v1/chat/completions`; a one-line Caddy proxy adds HTTPS and injects the bearer token.

**1. Enable the Hermes API Server** — add to `~/.hermes/.env`:
```
API_SERVER_ENABLED=true
API_SERVER_KEY=<a long random secret you choose>
API_SERVER_CORS_ORIGINS=https://localhost:3000
```
Start (or restart) the gateway — the API server runs inside it:
```
hermes gateway
```
Confirm:
```
curl http://localhost:8642/v1/health
# {"status":"ok",...}
```

**2. Run Caddy** (HTTPS + auth injection):
```
cp Caddyfile.example Caddyfile     # then edit it: paste your API_SERVER_KEY
caddy run
```
Confirm:
```
curl https://localhost:8643/v1/health
```
`Caddyfile` and `.env` are both git-ignored — never commit your real secret.

**3. Run an app** — see below.

## Running an app locally

Each add-in is built and sideloaded independently.

```
cd word            # or: cd excel
npm install
npm run build       # webpack --mode production
npm start            # sideloads the add-in into Office
```

`npm start` runs `office-addin-debugging`, which provisions a local HTTPS dev certificate (via `office-addin-dev-certs`), serves the add-in, and launches Office with it sideloaded. Then, in the Office app: **Home tab → Show Taskpane**.

To stop the sideloaded session:
```
npm run stop
```

See [`word/`](./word) and [`excel/README.md`](./excel/README.md) for per-app usage, prerequisites, and gotchas.

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

## License

MIT — see [LICENSE](./LICENSE).
