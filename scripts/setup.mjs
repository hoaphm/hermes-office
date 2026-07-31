#!/usr/bin/env node
// scripts/setup.mjs — Interactive provider configuration.
// Reads config.example.json (or existing config.json), prompts user for values,
// writes config.json into both dist/ directories and the project root.
"use strict";

import { createInterface } from "readline";
import fs from "fs/promises";
import path from "path";
import { fileURLToPath } from "url";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const rl = createInterface({ input: process.stdin, output: process.stdout });
const ask = (q) => new Promise((res) => rl.question(q + " ", res));

const resolveJsonFile = async (...parts) => {
  const p = path.resolve(...parts);
  try { await fs.access(p); return JSON.parse(await fs.readFile(p, "utf8")); } catch { return null; }
};

(async () => {
  console.log("=== Hermes Office — Provider Setup ===\n");
  console.log("This tool will configure an OpenAI-compatible LLM API provider.");
  console.log("Supported: OpenAI, OpenRouter, Azure OpenAI, Ollama, LM Studio, any /v1/chat/completions endpoint.\n");

  const src = await resolveJsonFile(scriptDir, "..", "config.example.json") ?? { baseUrl: "https://api.openai.com/v1", model: "gpt-4o-mini" };
  let baseUrl = (await ask(`Provider base URL [${src.baseUrl}]: `)) || src.baseUrl;
  if (!baseUrl.trim()) baseUrl = src.baseUrl;
  const apiKey = (await ask("API key: ")).trim();
  if (!apiKey) { console.error("Error: API key is required."); process.exit(1); }
  let model = (await ask("Model name [OpenRouter format, e.g. anthropic/claude-sonnet-4]: "))?.trim() || src.model;
  // For non-OpenRouter, strip the slash prefix (anthropic/model -> model)
  if (!baseUrl.includes("openrouter")) model = model.replace(/^[a-zA-Z]+\/(?=.)/i, "");

  baseUrl = baseUrl.replace(/\/+$/, "");

  // Validate HTTPS
  if (!/^https:\/\//i.test(baseUrl)) {
    console.error(`\nError: baseUrl must be https://. Got "${baseUrl}".`);
    console.error("(For local providers like Ollama/LM Studio, enable TLS in their settings.)");
    process.exit(1);
  }

  const config = { name: "Custom", baseUrl, apiKey, model };
  const configPath = path.join(scriptDir, "..", "config.json");
  const dir = path.dirname(configPath);
  try { await fs.access(dir); } catch { await fs.mkdir(dir, { recursive: true }); }
  await fs.writeFile(configPath, JSON.stringify(config, null, 2));
  console.log("\nConfig written to: " + configPath);
  console.log("\nNext steps:\n" +
    '  1. Build add-ins:        npm run build\n' +
    '  2. Start server:         npm run serve\n' +
    '  3. Sideload manifests via Office (Home → Manage Add-ins)\n');
  rl.close();
})();
