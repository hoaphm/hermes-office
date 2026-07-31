#!/usr/bin/env node
// scripts/launchagent.mjs — install/remove the macOS LaunchAgent that keeps the
// Local Gateway running.
//
// Opt-in, never automatic: an always-on Gateway means any local process can
// spend your Provider quota for the whole login session (ADR-0002). Callers ask
// the user first.
//
//   node scripts/launchagent.mjs --install [--force]
//   node scripts/launchagent.mjs --remove
"use strict";

import { execFile } from "child_process";
import fs from "fs";
import os from "os";
import path from "path";
import { fileURLToPath } from "url";
import { LAUNCH_AGENT_LABEL, launchAgentPlistPath } from "./gateway.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// XML text nodes cannot carry these raw. Paths with & or < are rare but a
// malformed plist fails in a way that is hard to trace back to here.
export function xmlEscape(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

export function plistContent({ caddyPath, repoRoot, logDir, label = LAUNCH_AGENT_LABEL }) {
  const e = xmlEscape;
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>${e(label)}</string>

    <key>ProgramArguments</key>
    <array>
        <string>${e(caddyPath)}</string>
        <string>run</string>
        <string>--config</string>
        <string>${e(path.join(repoRoot, "Caddyfile"))}</string>
    </array>

    <key>WorkingDirectory</key>
    <string>${e(repoRoot)}</string>

    <key>RunAtLoad</key>
    <true/>

    <key>KeepAlive</key>
    <true/>

    <key>StandardOutPath</key>
    <string>${e(path.join(logDir, "caddy.log"))}</string>

    <key>StandardErrorPath</key>
    <string>${e(path.join(logDir, "caddy.error.log"))}</string>
</dict>
</plist>
`;
}

const run = (cmd, args) =>
  new Promise((resolve) => {
    execFile(cmd, args, (err, stdout, stderr) => {
      resolve({ ok: !err, out: (stdout || "").trim(), err: (stderr || err?.message || "").trim() });
    });
  });

async function resolveCaddyPath() {
  const which = await run("command", ["-v", "caddy"]);
  if (which.ok && which.out) return which.out.split("\n")[0];
  for (const p of ["/opt/homebrew/bin/caddy", "/usr/local/bin/caddy", "/usr/bin/caddy"]) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

async function install({ force }) {
  const plistPath = launchAgentPlistPath();
  if (fs.existsSync(plistPath) && !force) {
    console.log(`LaunchAgent đã tồn tại: ${plistPath}`);
    console.log("Giữ nguyên file hiện có. Dùng --force nếu muốn ghi đè.");
    return 0;
  }

  const caddyPath = await resolveCaddyPath();
  if (!caddyPath) {
    console.error("Không tìm thấy caddy. Cài trước: brew install caddy");
    return 1;
  }

  const logDir = path.join(os.homedir(), "Library", "Logs", "hermes-office");
  fs.mkdirSync(logDir, { recursive: true });
  fs.mkdirSync(path.dirname(plistPath), { recursive: true });
  fs.writeFileSync(plistPath, plistContent({ caddyPath, repoRoot: root, logDir }));

  const lint = await run("plutil", ["-lint", plistPath]);
  if (!lint.ok) {
    console.error(`plist không hợp lệ: ${lint.err}`);
    return 1;
  }

  const target = `gui/${process.getuid()}`;
  // bootout first so re-installing over a loaded job is not an error.
  await run("launchctl", ["bootout", `${target}/${LAUNCH_AGENT_LABEL}`]);
  const boot = await run("launchctl", ["bootstrap", target, plistPath]);
  if (!boot.ok) {
    console.error(`Không nạp được LaunchAgent: ${boot.err}`);
    console.error(`Thử tay:  launchctl bootstrap ${target} ${plistPath}`);
    return 1;
  }

  console.log(`Đã cài LaunchAgent: ${plistPath}`);
  console.log(`Log: ${logDir}/caddy.error.log`);
  console.log(`Dừng hẳn khi cần:  launchctl bootout ${target}/${LAUNCH_AGENT_LABEL}`);
  return 0;
}

async function remove() {
  const plistPath = launchAgentPlistPath();
  const target = `gui/${process.getuid()}`;
  await run("launchctl", ["bootout", `${target}/${LAUNCH_AGENT_LABEL}`]);
  if (fs.existsSync(plistPath)) {
    fs.unlinkSync(plistPath);
    console.log(`Đã gỡ LaunchAgent: ${plistPath}`);
  } else {
    console.log("Không có LaunchAgent nào để gỡ.");
  }
  return 0;
}

async function main() {
  if (process.platform !== "darwin") {
    console.error("LaunchAgent chỉ dùng được trên macOS.");
    process.exit(1);
  }
  const args = process.argv.slice(2);
  const code = args.includes("--remove")
    ? await remove()
    : await install({ force: args.includes("--force") });
  process.exit(code);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
