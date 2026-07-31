# The API key lives in the Caddyfile, not in config.json

`config.json` is served over HTTP by the Local Gateway, so an API key stored
there is readable by any local process (`curl https://localhost:8643/config.json`)
and can be carried off the machine and spent anywhere. Since ADR-0001 makes the
Local Gateway mandatory anyway, the key is instead written into the Caddyfile —
which is git-ignored and never served — and Caddy attaches it to each forwarded
request. `config.json` is reduced to `{name, model}` and holds no secret at all.

## Consequences

- The trade accepted: while Caddy is running, any local process can POST to
  `:8643/v1/*` anonymously and spend quota. That is strictly better than the
  alternative, where the same process could steal the key itself.
- Stop the Local Gateway when not using the add-ins (`npm run stop`), and do not
  leave it running on a shared host. Note that this mitigation disappears
  entirely when the Gateway is supervised by an always-on launchd job — the
  exposure window becomes the whole login session. Running it that way is a
  choice to accept the wider window in exchange for convenience.
- Do not "simplify" by moving the key back into `config.json`. Serving it is the
  exposure.
