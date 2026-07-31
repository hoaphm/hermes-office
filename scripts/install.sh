#!/usr/bin/env bash
set -euo pipefail
node_v="$(node --version | sed 's/v//')" || { echo 'Node.js required'; exit 1; }
node -e 'const [a,b]=process.versions.node.split(".").map(Number); if (a<18||(a===18&&b<0)) process.exit(1);' || \
  { echo "Node.js >= 18.0.0 required, found $node_v"; exit 1; }
printf 'Hermes Office setup (macOS)\n'
command -v caddy >/dev/null || { printf 'Warning: Caddy not found (optional). Install: brew install caddy\n'; }
npm run build
node scripts/setup.mjs
printf '\nBuild ready. Run npm run serve, then sideload manifests.\n'
