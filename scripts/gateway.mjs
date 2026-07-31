// Helpers for talking about the Local Gateway process — shared by setup.mjs and
// serve.mjs.
//
// On this project's macOS setup the Gateway is usually run by a LaunchAgent
// (RunAtLoad + KeepAlive), not by `npm run serve`. That has two consequences
// both scripts have to handle: caddy holds the Caddyfile in memory, so editing
// it changes nothing until the job is restarted; and port 8643 is already bound,
// so spawning a second caddy silently splits requests between two instances
// running different configs.
import { execFile } from "child_process";
import net from "net";
import os from "os";
import path from "path";
import fs from "fs";

export const LAUNCH_AGENT_LABEL = "com.hermes.caddy";
export const DEFAULT_PORT = 8643;

export function launchAgentPlistPath() {
  return path.join(os.homedir(), "Library", "LaunchAgents", `${LAUNCH_AGENT_LABEL}.plist`);
}

export function hasLaunchAgent() {
  return process.platform === "darwin" && fs.existsSync(launchAgentPlistPath());
}

export function kickstartCommand() {
  return `launchctl kickstart -k gui/${process.getuid()}/${LAUNCH_AGENT_LABEL}`;
}

export function stopCommand() {
  return `launchctl bootout gui/${process.getuid()}/${LAUNCH_AGENT_LABEL}`;
}

// The port the Caddyfile actually binds, so a user who changed it does not get
// advice about the wrong number.
export function gatewayPort(caddyfileText) {
  const match = /localhost:(\d{2,5})\b/.exec(caddyfileText || "");
  return match ? Number(match[1]) : DEFAULT_PORT;
}

export function isPortBusy(port, timeoutMs = 1000) {
  return new Promise((resolve) => {
    const socket = net.connect({ host: "127.0.0.1", port });
    const done = (busy) => {
      socket.destroy();
      resolve(busy);
    };
    socket.setTimeout(timeoutMs);
    socket.once("connect", () => done(true));
    socket.once("timeout", () => done(false));
    socket.once("error", () => done(false));
  });
}

// Restart the LaunchAgent so it picks up a rewritten Caddyfile. Never throws —
// a failed restart is reported and the caller carries on.
export function restartLaunchAgent() {
  return new Promise((resolve) => {
    if (!hasLaunchAgent()) return resolve({ ok: false, reason: "absent" });
    execFile(
      "launchctl",
      ["kickstart", "-k", `gui/${process.getuid()}/${LAUNCH_AGENT_LABEL}`],
      (err, _stdout, stderr) => {
        if (err) resolve({ ok: false, reason: (stderr || err.message || "").trim() });
        else resolve({ ok: true });
      },
    );
  });
}
