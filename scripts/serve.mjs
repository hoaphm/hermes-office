#!/usr/bin/env node
// Start the Local Gateway for the built add-ins. Requires Caddy on PATH.
import { spawn } from "child_process";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import {
  gatewayPort,
  isPortBusy,
  hasLaunchAgent,
  kickstartCommand,
  stopCommand,
  LAUNCH_AGENT_LABEL,
} from "./gateway.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const caddyfile = path.join(root, "Caddyfile");
if (!fs.existsSync(caddyfile)) {
  console.error("Thiếu Caddyfile. Chạy `npm run setup` trước.");
  process.exit(1);
}

// Two caddy instances can both bind this port, after which requests are split
// between them — and the older one is still serving whatever Caddyfile it read
// at startup. That looks like a config change silently not taking effect, or
// intermittent 401s. Refuse to be the second instance.
const port = gatewayPort(fs.readFileSync(caddyfile, "utf8"));
if (await isPortBusy(port)) {
  if (hasLaunchAgent()) {
    console.log(`Local Gateway đã chạy sẵn trên cổng ${port} (LaunchAgent ${LAUNCH_AGENT_LABEL}).`);
    console.log("Không cần `npm run serve`. Nếu vừa sửa Caddyfile, nạp lại bằng:");
    console.log(`  ${kickstartCommand()}`);
    console.log("Dừng hẳn (KeepAlive sẽ bật lại nếu chỉ kill tiến trình):");
    console.log(`  ${stopCommand()}`);
    process.exit(0);
  }
  console.error(`Cổng ${port} đã có tiến trình khác đang nghe.`);
  console.error("Dừng nó trước, hoặc xem nó là gì bằng:");
  console.error(`  lsof -nP -iTCP:${port} -sTCP:LISTEN`);
  process.exit(1);
}

const child = spawn("caddy", ["run", "--config", caddyfile], {
  cwd: root,
  stdio: "inherit",
});
child.on("error", () => {
  console.error("Không tìm thấy Caddy. Cài Caddy và thêm vào PATH.");
  process.exit(1);
});
let serverStopped = false;
child.on("exit", (code) => {
  if (!serverStopped) {
    serverStopped = true;
    if (code !== 0) console.error(`Caddy thoát với mã ${code}.`);
    else console.log("Đã dừng Local Gateway.");
  }
});
process.on("SIGINT", () => {
  child.kill("SIGINT");
  serverStopped = true;
});
process.on("SIGTERM", () => {
  child.kill("SIGTERM");
  serverStopped = true;
});
