#!/usr/bin/env node
// scripts/stop.mjs — Stop local Caddy server.
import { spawn } from "child_process";
const child = spawn("caddy", ["stop"]);
child.on("error", (err) => {
  console.error(err.message);
  process.exit(1);
});
child.on("exit", (code) => {
  if (code !== 0) console.error(`Caddy exited with code ${code}`);
  console.log("Caddy stopped.");
});
