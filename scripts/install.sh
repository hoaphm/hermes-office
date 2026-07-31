#!/usr/bin/env bash
# scripts/install.sh — cài Hermes Office trên macOS.
#
# Tự khởi động (LaunchAgent) là TÙY CHỌN và mặc định TẮT: Gateway chạy suốt
# phiên đăng nhập nghĩa là mọi tiến trình local đều gọi được Provider của bạn
# trong khoảng thời gian đó (xem docs/adr/0002). Đặt HERMES_AUTOSTART=1 để bật
# không cần hỏi, HERMES_AUTOSTART=0 để bỏ qua hẳn.
set -euo pipefail

cd "$(dirname "$0")/.."

node_v="$(node --version | sed 's/v//')" || { echo 'Node.js required'; exit 1; }
node -e 'const [a,b]=process.versions.node.split(".").map(Number); if (a<18||(a===18&&b<0)) process.exit(1);' || \
  { echo "Node.js >= 18.0.0 required, found $node_v"; exit 1; }

printf 'Hermes Office setup (macOS)\n'
command -v caddy >/dev/null || { printf 'Warning: Caddy not found. Install: brew install caddy\n'; }

npm run build
node scripts/setup.mjs

autostart="${HERMES_AUTOSTART:-}"
if [ -z "$autostart" ]; then
  if [ -t 0 ]; then
    printf '\nChạy Local Gateway tự động khi đăng nhập?\n'
    printf '  Tiện: không phải mở terminal. Đánh đổi: Gateway mở suốt phiên,\n'
    printf '  mọi tiến trình trên máy đều gọi được Provider của bạn.\n'
    printf 'Bật tự khởi động? [y/N] '
    read -r reply || reply=""
    case "$reply" in [yY]*) autostart=1 ;; *) autostart=0 ;; esac
  else
    autostart=0
  fi
fi

if [ "$autostart" = "1" ]; then
  node scripts/launchagent.mjs --install
  printf '\nXong. Sideload word/dist/manifest.xml và excel/dist/manifest.xml trong Office.\n'
else
  printf '\nXong. Chạy `npm run serve` rồi sideload word/dist/manifest.xml và\n'
  printf 'excel/dist/manifest.xml trong Office.\n'
  printf 'Bật tự khởi động sau:  node scripts/launchagent.mjs --install\n'
fi
