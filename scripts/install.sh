#!/usr/bin/env bash
set -euo pipefail
node_v="$(node --version | sed 's/v//')" || { echo 'Node.js required'; exit 1; }
req="18.0.0"; cur="$node_v"
[ "$cur" \< "$req" ] && { echo "Node.js >= $req required, found $cur"; exit 1; }
printf 'Hermes Office setup (macOS)\n'
command -v caddy >/dev/null || { printf 'Warning: Caddy not found (optional). Install: brew install caddy\n'; }
npm run build
node scripts/setup.mjs
printf '\nBuild ready. Run npm run serve, then sideload manifests.\n'
