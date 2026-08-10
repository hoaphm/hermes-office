#!/usr/bin/env node
// scripts/setup.mjs — Interactive Provider configuration.
//
// Asks for the Provider (upstream OpenAI-compatible API) and writes two files:
//   Caddyfile   — Local Gateway config; holds the API key (git-ignored, 0600)
//   config.json — {name, model} only; served to the task panes, no secret
// See docs/adr/0001 and docs/adr/0002.
"use strict";

import { createInterface } from "readline";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";
import { hasLaunchAgent, restartLaunchAgent, kickstartCommand } from "./gateway.mjs";

const MIN_NODE_MAJOR = 18;
const MARKER_START = "# >>> hermes-proxy >>>";
const MARKER_END = "# <<< hermes-proxy <<<";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

// Created on first use, not at import: piped stdin reaches EOF and closes the
// interface before an eagerly created one is ever asked anything.
let rl = null;
const ask = (q) => {
  if (!rl) rl = createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((res) => rl.question(q + " ", res));
};
const closeInput = () => {
  if (rl) rl.close();
};

const die = (msg) => {
  console.error(`\nLỗi: ${msg}`);
  closeInput();
  process.exit(1);
};

const readJsonFile = async (...parts) => {
  try {
    return JSON.parse(await fs.readFile(path.resolve(...parts), "utf8"));
  } catch {
    return null;
  }
};

// Caddyfile strings are placeholder-expanded, so a literal { must be doubled.
export const caddyQuote = (s) =>
  s.replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\{/g, "{{");

// Build the proxy block mapping the Gateway's /v1/* onto the Provider, attaching
// the key on the way out. The add-in always calls /v1/chat/completions;
// handle_path strips that /v1, and the rewrite re-adds whatever path prefix the
// Provider actually uses (/v1 for most, none for some, deeper for Azure).
//
// The handle_errors clause belongs in here, not beside it in Caddyfile.example:
// setup rewrites only what is between the markers, so anything outside them
// never reaches a Caddyfile that already exists — the feature would work on
// fresh installs and silently not on every current one. Generated here, one
// `npm run setup` repairs any install.
export function proxyBlock(provider, apiKey) {
  const url = new URL(provider);
  const prefix = url.pathname.replace(/\/+$/, "");
  return [
    MARKER_START,
    "handle_path /v1/* {",
    prefix ? `    rewrite * ${prefix}{uri}` : null,
    `    reverse_proxy ${url.origin} {`,
    `        header_up Host ${url.host}`,
    `        header_up Authorization "Bearer ${caddyQuote(apiKey)}"`,
    "    }",
    "}",
    "",
    "# Mark responses the Gateway generated itself, so the Task Pane can tell an",
    "# Upstream Hop failure (Caddy never reached the Provider — what a change of",
    "# network looks like) from the same status code sent by the Provider, which",
    "# passes through unmarked. See ADR-0005 and shared/failures.js.",
    "handle_errors {",
    '    header X-Hermes-Gateway-Error "{err.status_code}"',
    '    respond "Local Gateway error" {err.status_code}',
    "}",
    MARKER_END,
  ]
    .filter((line) => line !== null)
    .join("\n");
}

// Replace only the marked region, so hand edits elsewhere in a working
// Caddyfile survive re-running setup. The block is re-indented to match the
// marker's own indentation, so regenerating an already-generated file is a
// no-op rather than a slow drift to the left margin.
export function spliceProxyBlock(template, block) {
  const start = template.indexOf(MARKER_START);
  const end = template.indexOf(MARKER_END);
  if (start === -1 || end === -1 || end < start) return null;
  const lineStart = template.lastIndexOf("\n", start) + 1;
  const indent = template.slice(lineStart, start);
  const indented = /^\s*$/.test(indent)
    ? block
        .split("\n")
        .map((line) => (line ? indent + line : line))
        .join("\n")
    : block;
  return (
    template.slice(0, lineStart) + indented + template.slice(end + MARKER_END.length)
  );
}

async function writeCaddyfile(provider, apiKey) {
  const caddyfilePath = path.join(root, "Caddyfile");
  let template = await fs.readFile(caddyfilePath, "utf8").catch(() => null);
  const created = template === null;
  if (created) {
    template = await fs
      .readFile(path.join(root, "Caddyfile.example"), "utf8")
      .catch(() => die("thiếu Caddyfile.example — repo không đầy đủ."));
  }

  const next = spliceProxyBlock(template, proxyBlock(provider, apiKey));
  if (next === null) {
    die(
      `Caddyfile hiện có không chứa cặp mốc ${MARKER_START} … ${MARKER_END}.\n` +
        "Thêm hai dòng mốc đó vào (xem Caddyfile.example) rồi chạy lại, hoặc xóa\n" +
        "Caddyfile để setup sinh lại từ đầu.",
    );
  }
  await fs.writeFile(caddyfilePath, next, { mode: 0o600 });
  await fs.chmod(caddyfilePath, 0o600).catch(() => {});
  return { caddyfilePath, created };
}

async function main() {
  const major = Number(process.versions.node.split(".")[0]);
  if (major < MIN_NODE_MAJOR) {
    console.error(
      `Lỗi: cần Node.js >= ${MIN_NODE_MAJOR}, đang chạy ${process.versions.node}.`,
    );
    process.exit(1);
  }

  console.log("=== Hermes Office — thiết lập Provider ===\n");
  console.log("Provider là API tương thích OpenAI mà add-in sẽ hỏi.");
  console.log(
    "Hỗ trợ: OpenAI, OpenRouter, Azure OpenAI, router tự host — bất kỳ endpoint /chat/completions nào.",
  );
  console.log("Khóa API ghi vào Caddyfile (git-ignored), KHÔNG nằm trong config.json.\n");

  const example = (await readJsonFile(root, "config.example.json")) ?? {};
  const defaultProvider = "https://api.openai.com/v1";
  const defaultModel = example.model ?? "gpt-4o-mini";

  // Non-interactive path (reinstall scripts, CI). The key comes from the
  // environment, never from a flag — argv is visible to every process on the box
  // and lands in shell history.
  const flag = (name) => {
    const hit = process.argv.find((a) => a.startsWith(`--${name}=`));
    return hit ? hit.slice(name.length + 3).trim() : "";
  };
  const envKey = (process.env.HERMES_API_KEY ?? "").trim();

  let provider =
    flag("provider") || (await ask(`URL Provider [${defaultProvider}]:`)).trim() || defaultProvider;
  provider = provider.replace(/\/+$/, "");
  if (!/^https:\/\//i.test(provider)) {
    die(
      `URL Provider phải bắt đầu bằng https://. Nhận được "${provider}".\n` +
        "(Provider chạy local như Ollama/LM Studio: bật TLS trong cài đặt của chúng.)",
    );
  }
  try {
    new URL(provider);
  } catch {
    die(`URL Provider không hợp lệ: "${provider}".`);
  }

  const apiKey = envKey || (await ask("Khóa API:")).trim();
  if (!apiKey) die("khóa API là bắt buộc.");

  // Taken verbatim. An earlier version stripped a `vendor/` prefix for
  // non-OpenRouter providers, but self-hosted routers use that form too
  // (`ds/deepseek-v4-flash`), so the guess silently broke working model ids.
  // Only the Provider knows what it calls its models.
  const model =
    flag("model") ||
    (await ask(`Tên model, đúng như Provider gọi [mặc định ${defaultModel}]:`)).trim() ||
    defaultModel;

  const { caddyfilePath, created } = await writeCaddyfile(provider, apiKey);
  const configPath = path.join(root, "config.json");
  // Re-running setup must not silently re-arm the Custom Function gate: whether
  // `=HERMES.*` may call the Provider is a security choice the user made, not a
  // Provider setting. Carry it over; default off.
  const previous = (await readJsonFile(configPath)) ?? {};
  await fs.writeFile(
    configPath,
    JSON.stringify(
      {
        name: new URL(provider).host,
        model,
        customFunctions: previous.customFunctions === true,
      },
      null,
      2,
    ) + "\n",
  );

  console.log(
    `\n${created ? "Đã tạo" : "Đã cập nhật khối proxy trong"} ${caddyfilePath}`,
  );
  console.log(`Đã ghi ${configPath} — model "${model}", không chứa khóa.`);

  // A running Gateway holds the Caddyfile in memory, so without this the new
  // Provider or key silently has no effect until the process is restarted.
  if (hasLaunchAgent()) {
    const result = await restartLaunchAgent();
    if (result.ok) {
      console.log("Đã khởi động lại Local Gateway (LaunchAgent) để nạp cấu hình mới.");
    } else {
      console.log(
        `\n⚠ Không khởi động lại được LaunchAgent (${result.reason || "không rõ"}).\n` +
          `  Cấu hình mới CHƯA có hiệu lực. Chạy tay:  ${kickstartCommand()}`,
      );
    }
  }

  console.log(
    "\nTiếp theo:\n" +
      "  1. Build add-in:        npm run build\n" +
      (hasLaunchAgent()
        ? "  2. Gateway do LaunchAgent chạy sẵn — không cần npm run serve\n"
        : "  2. Chạy Local Gateway:  npm run serve\n") +
      "  3. Sideload manifest trong Word/Excel (Home → Manage Add-ins)\n",
  );
  closeInput();
}

// Importable for tests; only prompts when run as a script.
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  main();
}
