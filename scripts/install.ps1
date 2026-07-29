# One-shot setup for a fresh clone: dependencies -> build -> provider config.
$ErrorActionPreference = "Stop"

Set-Location (Join-Path $PSScriptRoot "..")

Write-Host "Hermes Office setup (Windows)" -ForegroundColor Cyan

if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw "Node.js 18+ required. Install from https://nodejs.org/" }
if ([int](node -p "process.versions.node.split('.')[0]") -lt 18) { throw "Node.js 18+ required, found $(node --version)" }
if (-not (Get-Command caddy -ErrorAction SilentlyContinue)) {
  Write-Warning "Caddy not found. Install it before 'npm run serve': winget install CaddyServer.Caddy"
}

# Each add-in has its own package.json and node_modules; a fresh clone has
# neither installed, so 'npm run build' would fail on a missing webpack.
Write-Host "`nInstalling dependencies..."
npm install
npm install --prefix word
npm install --prefix excel

Write-Host "`nBuilding add-ins..."
npm run build

Write-Host ""
node scripts/setup.mjs

Write-Host "`nDone. Run 'npm run serve', then sideload word\dist\manifest.xml and excel\dist\manifest.xml." -ForegroundColor Green
