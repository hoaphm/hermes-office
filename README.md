# Hermes for Office

Put your own [Hermes](https://github.com/NousResearch/hermes-agent) agent **inside Microsoft Office** — your models, your skills, your approval gate. Not Copilot.

```
Office add-in ──HTTPS──▶ Caddy (:8643, TLS + injects key) ──▶ Hermes API Server (:8642) ──▶ your agent
```

## Apps

| App | Status | Folder |
|-----|--------|--------|
| Excel — task-pane chat + `=HERMES.*` custom functions | ✅ available | [`excel/`](./excel) |
| Word | planned | — |
| PowerPoint | planned | — |

Each app is a self-contained Office.js add-in. They all share the same backend (below).

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
Confirm: `curl http://localhost:8642/v1/health` → `{"status":"ok",...}`

**2. Run Caddy** (HTTPS + auth injection):
```
cp Caddyfile.example Caddyfile     # then edit it: paste your API_SERVER_KEY
caddy run
```
Confirm: `curl https://localhost:8643/v1/health`

**3. Run an app** — `cd` into the app folder (e.g. `excel/`) and follow its README.

## Security

⚠️ The API Server gives the agent's **full toolset, including terminal commands**. Treat `API_SERVER_KEY` like a password:
- Long and random; **never commit** your real `Caddyfile` or `.env` (both are git-ignored).
- Bind everything to `localhost`; keep `API_SERVER_CORS_ORIGINS` narrow.
- No add-in ever holds the key — Caddy injects it.

## License

MIT — see [LICENSE](./LICENSE).
