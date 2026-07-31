import test from "node:test";
import assert from "node:assert/strict";
import { caddyQuote, proxyBlock, spliceProxyBlock } from "./setup.mjs";

test("proxyBlock re-adds the Provider's own path prefix", () => {
  const block = proxyBlock("https://router.example.test/v1", "k");
  assert.match(block, /rewrite \* \/v1\{uri\}/);
  assert.match(block, /reverse_proxy https:\/\/router\.example\.test\b/);
  assert.match(block, /header_up Host router\.example\.test/);
});

test("proxyBlock omits the rewrite when the Provider has no path prefix", () => {
  const block = proxyBlock("https://router.example.test", "k");
  assert.equal(block.includes("rewrite"), false);
});

test("proxyBlock keeps a deep Provider path (Azure-style)", () => {
  const block = proxyBlock("https://x.openai.azure.com/openai/deployments/gpt4", "k");
  assert.match(block, /rewrite \* \/openai\/deployments\/gpt4\{uri\}/);
});

test("proxyBlock attaches the key as a bearer token", () => {
  assert.match(proxyBlock("https://p.test/v1", "sk-abc"), /Authorization "Bearer sk-abc"/);
});

test("caddyQuote escapes quotes, backslashes and placeholder braces", () => {
  assert.equal(caddyQuote('a"b'), 'a\\"b');
  assert.equal(caddyQuote("a\\b"), "a\\\\b");
  assert.equal(caddyQuote("a{env.HOME}"), "a{{env.HOME}");
});

test("spliceProxyBlock replaces only the marked region", () => {
  const template = [
    "before",
    "# >>> hermes-proxy >>>",
    "old contents",
    "# <<< hermes-proxy <<<",
    "after",
  ].join("\n");
  const out = spliceProxyBlock(template, "# >>> hermes-proxy >>>\nNEW\n# <<< hermes-proxy <<<");
  assert.equal(out, "before\n# >>> hermes-proxy >>>\nNEW\n# <<< hermes-proxy <<<\nafter");
  assert.equal(out.includes("old contents"), false);
});

test("spliceProxyBlock refuses a template without markers", () => {
  assert.equal(spliceProxyBlock("no markers here", "block"), null);
});

test("a generated Caddyfile can be regenerated without drift", () => {
  const template = [
    "https://localhost:8643 {",
    "    # hand edit that must survive",
    "    # >>> hermes-proxy >>>",
    "    # (run `npm run setup`)",
    "    # <<< hermes-proxy <<<",
    "}",
  ].join("\n");
  const once = spliceProxyBlock(template, proxyBlock("https://p.test/v1", "k1"));
  const twice = spliceProxyBlock(once, proxyBlock("https://p.test/v1", "k1"));
  assert.equal(once, twice);
  assert.match(twice, /hand edit that must survive/);
  // Both markers keep the template's indentation, so the block cannot creep
  // leftwards each time setup runs.
  assert.match(twice, /\n {4}# >>> hermes-proxy >>>/);
  assert.match(twice, /\n {4}# <<< hermes-proxy <<</);
});

test("regenerating with a new key leaves no trace of the old one", () => {
  const template = "# >>> hermes-proxy >>>\nx\n# <<< hermes-proxy <<<";
  const first = spliceProxyBlock(template, proxyBlock("https://p.test/v1", "OLD-KEY"));
  const second = spliceProxyBlock(first, proxyBlock("https://p.test/v1", "NEW-KEY"));
  assert.equal(second.includes("OLD-KEY"), false);
  assert.match(second, /NEW-KEY/);
});
