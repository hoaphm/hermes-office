$ErrorActionPreference = "Stop"
Write-Host "Hermes Office setup (Windows)" -ForegroundColor Cyan
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw "Node.js 18+ required. Install from https://nodejs.org/" }
if ([int](node -p "process.versions.node.split('.')[0]") -lt 18) { throw "Node.js 18+ required." }
if (-not (Get-Command caddy -ErrorAction SilentlyContinue)) { Write-Warning "Caddy not found. Install: winget install CaddyServer.Caddy" }
npm run build
node scripts/setup.mjs
Write-Host "Build ready. Run 'npm run serve' in this window, then sideload word/dist/manifest.xml and excel/dist/manifest.xml." -ForegroundColor Green
