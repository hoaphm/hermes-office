#!/usr/bin/env bash
# One-shot setup for a fresh clone: dependencies -> build -> provider config.
set -euo pipefail

cd "$(dirname "$0")/.."

command -v node >/dev/null || { echo 'Node.js 18+ required. Install from https://nodejs.org/'; exit 1; }
# Compare the MAJOR version numerically. A string compare ("9.0.0" < "18.0.0")
# is false in lexicographic order, so Node 9 used to sail past this guard.
node_major="$(node -p 'process.versions.node.split(".")[0]')"
if [ "$node_major" -lt 18 ]; then
  echo "Node.js >= 18 required, found $(node --version)"
  exit 1
fi

printf 'Hermes Office setup (macOS)\n\n'
command -v caddy >/dev/null || printf 'Warning: Caddy not found. Install it before `npm run serve`: brew install caddy\n\n'

# Each add-in has its own package.json and node_modules; a fresh clone has
# neither installed, so `npm run build` would fail on a missing webpack.
printf 'Installing dependencies…\n'
npm install
npm install --prefix word
npm install --prefix excel

printf '\nBuilding add-ins…\n'
npm run build

printf '\n'
node scripts/setup.mjs

printf '\nDone. Run `npm run serve`, then sideload word/dist/manifest.xml and excel/dist/manifest.xml.\n'
