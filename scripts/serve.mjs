#!/usr/bin/env node
// Start local HTTPS server for built add-ins. Requires Caddy on PATH.
import { spawn } from "child_process";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

if (!fs.existsSync(path.join(root, "Caddyfile"))) {
  console.error("Missing Caddyfile. Run `npm run setup` first (it creates one from Caddyfile.example).");
  process.exit(1);
}
if (!fs.existsSync(path.join(root, "config.json"))) {
  console.error("Missing config.json. Run `npm run setup` first — the add-ins cannot reach a provider without it.");
  process.exit(1);
}
for (const app of ["word", "excel"]) {
  if (!fs.existsSync(path.join(root, app, "dist", "taskpane.html"))) {
    console.error(`Missing ${app}/dist. Run \`npm run build\` first.`);
    process.exit(1);
  }
}

const child = spawn("caddy", ["run", "--config", path.join(root, "Caddyfile")], {
  cwd: root,
  stdio: "inherit",
});
child.on("error", () => {
  console.error("Caddy not found. Install Caddy and add it to PATH.");
  process.exit(1);
});
process.on("SIGINT", () => child.kill("SIGINT"));
process.on("SIGTERM", () => child.kill("SIGTERM"));
