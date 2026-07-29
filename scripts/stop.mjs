#!/usr/bin/env node
// scripts/stop.mjs — Stop local Caddy server.
import { spawn } from "child_process";
const child = spawn("caddy", ["stop"]);
child.on("error", (err) => console.error(err.message));
child.on("exit", () => console.log("Caddy stopped."));
