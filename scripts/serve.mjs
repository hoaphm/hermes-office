#!/usr/bin/env node
// Start local HTTPS server for built add-ins. Requires Caddy on PATH.
import { spawn } from "child_process";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
if (!fs.existsSync(path.join(root, "Caddyfile"))) {
  console.error("Missing Caddyfile. Run npm run setup first.");
  process.exit(1);
}
const child = spawn("caddy", ["run", "--config", path.join(root, "Caddyfile")], {
  cwd: root,
  stdio: "inherit",
});
child.on("error", () => {
  console.error("Caddy not found. Install Caddy and add it to PATH.");
  process.exit(1);
});
let serverStopped = false;
child.on("exit", (code) => {
  if (!serverStopped) {
    serverStopped = true;
    if (code !== 0) console.error(`Caddy exited with code ${code}.`);
    else console.log("Caddy stopped.");
  }
});
process.on("SIGINT", () => { child.kill("SIGINT"); serverStopped = true; });
process.on("SIGTERM", () => { child.kill("SIGTERM"); serverStopped = true; });
