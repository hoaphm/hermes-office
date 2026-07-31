#!/usr/bin/env node
// scripts/stop.mjs — Stop the Local Gateway.
//
// `caddy stop` only asks the process to exit. When the Gateway is run by the
// LaunchAgent (KeepAlive), launchd starts it straight back up — so verify
// afterwards rather than claiming success.
import { spawn } from "child_process";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { gatewayPort, isPortBusy, hasLaunchAgent, stopCommand } from "./gateway.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const caddyfile = path.join(root, "Caddyfile");
const port = fs.existsSync(caddyfile)
  ? gatewayPort(fs.readFileSync(caddyfile, "utf8"))
  : gatewayPort("");

const child = spawn("caddy", ["stop"]);
child.on("error", (err) => {
  console.error(err.message);
  process.exit(1);
});
child.on("exit", async (code) => {
  if (code !== 0) console.error(`Caddy thoát với mã ${code}`);
  // Give launchd a moment to do its thing before checking.
  await new Promise((r) => setTimeout(r, 3000));
  if (!(await isPortBusy(port))) {
    console.log("Đã dừng Local Gateway.");
    return;
  }
  if (hasLaunchAgent()) {
    console.log(`Local Gateway vẫn đang chạy trên cổng ${port}: LaunchAgent đã bật lại (KeepAlive).`);
    console.log("Dừng hẳn cho tới lần đăng nhập sau:");
    console.log(`  ${stopCommand()}`);
  } else {
    console.log(`Cổng ${port} vẫn có tiến trình đang nghe — Local Gateway chưa dừng.`);
  }
});
