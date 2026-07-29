import test from "node:test";
import assert from "node:assert/strict";
import { buildProxyBlock, applyProxyBlock, LOCAL_PROXY_BASE } from "./setup.mjs";

const TEMPLATE = `https://localhost:8643 {
    tls internal

    handle /config.json {
        root * .
        file_server
    }

    # >>> hermes-proxy >>>
    # (disabled — add-ins call the provider directly)
    # <<< hermes-proxy <<<
}
`;

test("buildProxyBlock: provider on /v1 re-adds its own base path", () => {
  const block = buildProxyBlock("https://ninerouter.duckdns.org/v1");
  assert.match(block, /handle_path \/v1\/\*/);
  assert.match(block, /rewrite \* \/v1\{uri\}/);
  assert.match(block, /reverse_proxy https:\/\/ninerouter\.duckdns\.org\b/);
  assert.match(block, /header_up Host ninerouter\.duckdns\.org/);
});

test("buildProxyBlock: provider on a non-/v1 base path", () => {
  // A plain reverse_proxy would send this provider /v1/chat/completions and
  // 404. handle_path strips the local prefix, the rewrite adds the real one.
  const block = buildProxyBlock("https://example.test/api/v3");
  assert.match(block, /rewrite \* \/api\/v3\{uri\}/);
  assert.match(block, /reverse_proxy https:\/\/example\.test\b/);
});

test("buildProxyBlock: provider at the domain root emits no rewrite", () => {
  const block = buildProxyBlock("https://example.test");
  assert.doesNotMatch(block, /rewrite/);
  assert.match(block, /reverse_proxy https:\/\/example\.test\b/);
});

test("buildProxyBlock: non-default port is preserved in host and upstream", () => {
  const block = buildProxyBlock("https://example.test:8443/v1");
  assert.match(block, /reverse_proxy https:\/\/example\.test:8443\b/);
  assert.match(block, /header_up Host example\.test:8443/);
});

test("applyProxyBlock: writes the block between the markers", () => {
  const out = applyProxyBlock(TEMPLATE, "https://ninerouter.duckdns.org/v1");
  assert.match(out, /handle_path \/v1\/\*/);
  assert.doesNotMatch(out, /\(disabled/);
  // Everything outside the managed region is untouched.
  assert.match(out, /handle \/config\.json/);
  assert.match(out, /tls internal/);
});

test("applyProxyBlock: is idempotent across re-runs", () => {
  const once = applyProxyBlock(TEMPLATE, "https://a.test/v1");
  const twice = applyProxyBlock(once, "https://a.test/v1");
  assert.equal(twice, once);
  // Exactly one managed region, not a stack of them.
  assert.equal(once.match(/# >>> hermes-proxy >>>/g).length, 1);
  assert.equal(once.match(/handle_path/g).length, 1);
});

test("applyProxyBlock: switching providers replaces, never appends", () => {
  const first = applyProxyBlock(TEMPLATE, "https://a.test/v1");
  const second = applyProxyBlock(first, "https://b.test/v1");
  assert.match(second, /reverse_proxy https:\/\/b\.test\b/);
  assert.doesNotMatch(second, /a\.test/);
  assert.equal(second.match(/handle_path/g).length, 1);
});

test("applyProxyBlock: null upstream disables the block again", () => {
  const on = applyProxyBlock(TEMPLATE, "https://a.test/v1");
  const off = applyProxyBlock(on, null);
  assert.doesNotMatch(off, /handle_path/);
  assert.doesNotMatch(off, /a\.test/);
  assert.match(off, /\(disabled/);
});

test("applyProxyBlock: returns null when the markers are missing", () => {
  assert.equal(applyProxyBlock("https://localhost:8643 {\n  tls internal\n}\n", "https://a.test/v1"), null);
});

test("LOCAL_PROXY_BASE matches the port the Caddyfile serves", () => {
  assert.equal(LOCAL_PROXY_BASE, "https://localhost:8643/v1");
  assert.match(TEMPLATE, /https:\/\/localhost:8643 \{/);
});
