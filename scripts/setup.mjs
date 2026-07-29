#!/usr/bin/env node
// scripts/setup.mjs — Interactive provider configuration.
//
// Writes two files, both git-ignored:
//   config.json  — provider baseUrl / apiKey / model, served to the task panes
//   Caddyfile    — copied from Caddyfile.example if absent, so `npm run serve`
//                  has something to run (serve.mjs requires it)
//
// The API key is written to the repo root ONLY. It deliberately does not go
// into word/dist or excel/dist: the task panes fetch the root-relative
// "/config.json", which Caddy serves from the repo root, so a copy in dist
// would be unused — and dist is exactly the folder a user is told to open
// when sideloading, making it the easiest place to leak a key from.
"use strict";

import { createInterface } from "readline";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath, pathToFileURL } from "url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(scriptDir, "..");

// Two input paths.
//
// On a TTY, prompt interactively via readline. On a non-TTY stdin (answers
// piped in, or a test harness) readline is the wrong tool: the stream hits EOF
// while the script is still awaiting earlier work, the interface closes, and
// every remaining question either throws ERR_USE_AFTER_CLOSE or hangs
// unresolved until the process quietly exits without writing anything. So for
// piped input, drain stdin once and answer from that queue.
let rl;
let piped = null;

async function drainStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(chunk);
  return Buffer.concat(chunks).toString("utf8").split(/\r?\n/);
}

async function ask(q) {
  if (!process.stdin.isTTY) {
    piped ??= await drainStdin();
    const answer = piped.shift() ?? "";
    process.stdout.write(q + " " + answer + "\n");
    return answer;
  }
  rl ??= createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((res) => rl.question(q + " ", res));
}

const closeInput = () => rl?.close();

const readJsonFile = async (...parts) => {
  const p = path.resolve(...parts);
  try {
    return JSON.parse(await fs.readFile(p, "utf8"));
  } catch {
    return null;
  }
};

const exists = async (p) => {
  try {
    await fs.access(p);
    return true;
  } catch {
    return false;
  }
};

const PROXY_BEGIN = "# >>> hermes-proxy >>>";
const PROXY_END = "# <<< hermes-proxy <<<";
export const LOCAL_PROXY_BASE = "https://localhost:8643/v1";

// The Caddy block that forwards /v1/* to the real provider.
//
// handle_path strips the local /v1 prefix and the rewrite re-adds the
// provider's OWN base path, so this works for providers whose base is not
// /v1 (e.g. https://host/api/v3) — a plain reverse_proxy would send them
// /v1/chat/completions and 404.
export function buildProxyBlock(upstreamBaseUrl) {
  const u = new URL(upstreamBaseUrl);
  const basePath = u.pathname.replace(/\/+$/, "");
  const origin = `${u.protocol}//${u.host}`;
  const rewrite = basePath ? `        rewrite * ${basePath}{uri}\n` : "";
  return (
    `    handle_path /v1/* {\n` +
    rewrite +
    `        reverse_proxy ${origin} {\n` +
    `            header_up Host ${u.host}\n` +
    `        }\n` +
    `    }`
  );
}

// Replace the managed region in a Caddyfile. Idempotent: re-running setup with
// a different provider rewrites the block rather than stacking a second one.
export function applyProxyBlock(caddyfile, upstreamBaseUrl) {
  const body = upstreamBaseUrl
    ? buildProxyBlock(upstreamBaseUrl)
    : "    # (disabled — add-ins call the provider directly)";
  const region = `${PROXY_BEGIN}\n${body}\n    ${PROXY_END}`;
  const re = new RegExp(
    `${PROXY_BEGIN.replace(/[>]/g, "\\$&")}[\\s\\S]*?${PROXY_END.replace(/[<]/g, "\\$&")}`
  );
  if (!re.test(caddyfile)) return null;
  return caddyfile.replace(re, region);
}

// Create Caddyfile from the example if absent, then write the proxy block.
async function writeCaddyfile(upstreamBaseUrl) {
  const target = path.join(root, "Caddyfile");
  const example = path.join(root, "Caddyfile.example");
  let created = false;
  if (!(await exists(target))) {
    if (!(await exists(example))) {
      console.error("Warning: Caddyfile.example is missing; cannot create Caddyfile.");
      return null;
    }
    await fs.copyFile(example, target);
    created = true;
  }
  const current = await fs.readFile(target, "utf8");
  const next = applyProxyBlock(current, upstreamBaseUrl);
  if (next === null) {
    console.error(
      `Warning: ${target} has no ${PROXY_BEGIN} / ${PROXY_END} markers, so the proxy block\n` +
        "         could not be written. Copy Caddyfile.example over it, or add the markers by hand."
    );
    return { path: target, created, proxied: false };
  }
  if (next !== current) await fs.writeFile(target, next);
  return { path: target, created, proxied: !!upstreamBaseUrl };
}

// Only prompt when run as a script. The pure helpers above are imported by
// scripts/setup.test.mjs, which must not trigger an interactive session.
const isMain =
  process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href;

if (isMain)
  (async () => {
  console.log("=== Hermes Office — Provider Setup ===\n");
  console.log("This tool will configure an OpenAI-compatible LLM API provider.");
  console.log(
    "Supported: OpenAI, OpenRouter, Azure OpenAI, Ollama, LM Studio, any /v1/chat/completions endpoint.\n"
  );

  // Prefer the values already in config.json so re-running setup is a cheap
  // edit rather than a from-scratch retype.
  const src = (await readJsonFile(root, "config.json")) ??
    (await readJsonFile(root, "config.example.json")) ?? {
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o-mini",
    };

  // When the previous run enabled the proxy, config.json's baseUrl is the
  // LOCAL address; the real provider is kept in upstreamUrl. Default the
  // prompt to the provider either way, never to localhost.
  const prevProvider = src.upstreamUrl || src.baseUrl;

  let providerUrl = ((await ask(`Provider base URL [${prevProvider}]: `)) || "").trim() || prevProvider;
  providerUrl = providerUrl.replace(/\/+$/, "");
  if (!/^https:\/\//i.test(providerUrl)) {
    console.error(`\nError: baseUrl must be https://. Got "${providerUrl}".`);
    console.error("(For local providers like Ollama/LM Studio, enable TLS in their settings.)");
    process.exit(1);
  }

  const apiKey = (await ask("API key: ")).trim();
  if (!apiKey) {
    console.error("Error: API key is required.");
    process.exit(1);
  }

  // Taken verbatim. An earlier version stripped a leading "vendor/" segment
  // for non-OpenRouter base URLs, which silently corrupted legitimate ids
  // such as Together's "meta-llama/Llama-3-70b".
  const model =
    ((await ask(`Model id, exactly as your provider names it [${src.model}]: `)) || "").trim() ||
    src.model;
  if (!model) {
    console.error("Error: model is required.");
    process.exit(1);
  }

  // Routing choice. The task pane runs in a browser, so a direct provider call
  // is cross-origin and needs the provider to answer a CORS preflight without
  // an API key — which most self-hosted routers do not. Proxying through the
  // local server makes the call same-origin, so no preflight is ever sent.
  console.log(
    "\nRoute provider calls through the local server (https://localhost:8643/v1)?\n" +
      "  Recommended. The task pane runs in a browser, so calling the provider\n" +
      "  directly requires it to support CORS preflight without an API key.\n" +
      "  Proxying locally makes the call same-origin and sidesteps that entirely.\n" +
      "  Caddy only forwards — it injects no credentials."
  );
  const answer = ((await ask("Use local proxy? [Y/n]:")) || "").trim().toLowerCase();
  const useProxy = answer === "" || answer === "y" || answer === "yes";

  const config = useProxy
    ? { name: "Custom", baseUrl: LOCAL_PROXY_BASE, upstreamUrl: providerUrl, apiKey, model }
    : { name: "Custom", baseUrl: providerUrl, apiKey, model };

  const configPath = path.join(root, "config.json");
  await fs.writeFile(configPath, JSON.stringify(config, null, 2));
  // Owner-only: the file holds a plaintext API key.
  await fs.chmod(configPath, 0o600).catch(() => {});

  const caddy = await writeCaddyfile(useProxy ? providerUrl : null);

  console.log(`\nWrote ${configPath} (mode 600)`);
  console.log(
    useProxy
      ? `  add-in  -> ${LOCAL_PROXY_BASE}  (same-origin)\n  Caddy   -> ${providerUrl}`
      : `  add-in  -> ${providerUrl}  (direct; provider must support browser CORS)`
  );
  if (caddy) {
    console.log(caddy.created ? `Created ${caddy.path} from Caddyfile.example` : `Updated ${caddy.path}`);
  }
  console.log(
    "\nNext steps:\n" +
      "  1. Build add-ins (if not built yet):  npm run build\n" +
      "  2. Start server:                      npm run serve\n" +
      "  3. Sideload word/dist/manifest.xml and excel/dist/manifest.xml via Office\n"
  );
  closeInput();
})();
