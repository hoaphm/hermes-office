# The Local Gateway is the only path to the Provider

Add-ins could call the Provider directly from their WebView, and for OpenAI or
OpenRouter that works. It fails for self-hosted OpenAI-compatible routers, which
commonly do not answer the CORS preflight the browser sends for a cross-origin
POST with an `Authorization` header. Routing every call through the Local
Gateway at `https://localhost:8643/v1/*` makes the request same-origin, so no
preflight is ever sent and every Provider works the same way.

## Consequences

- `config.json` carries no `baseUrl`. The add-in always calls same-origin
  `/v1/chat/completions`; the Provider's real address lives only in the
  Caddyfile. One concept, one place — the drift between README, setup script,
  and Caddyfile that prompted this decision came from having it in two.
- The Local Gateway is required, not optional. Nothing works without
  `npm run serve`.
- Anything deriving behaviour from the Provider's address (for example
  stripping a `vendor/` prefix from the model name for non-OpenRouter
  providers) must read the Provider URL, not the URL the add-in calls — the
  latter is now always `localhost`.
