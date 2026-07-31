import test from "node:test";
import assert from "node:assert/strict";
import { plistContent, xmlEscape } from "./launchagent.mjs";

const base = { caddyPath: "/opt/homebrew/bin/caddy", repoRoot: "/Users/x/hermes-office", logDir: "/Users/x/Library/Logs/hermes-office" };

test("plistContent points caddy at the repo's own Caddyfile", () => {
  const out = plistContent(base);
  assert.match(out, /<string>\/opt\/homebrew\/bin\/caddy<\/string>/);
  assert.match(out, /<string>\/Users\/x\/hermes-office\/Caddyfile<\/string>/);
  assert.match(out, /<key>WorkingDirectory<\/key>\s*<string>\/Users\/x\/hermes-office<\/string>/);
});

test("plistContent keeps the label the other scripts look for", () => {
  // gateway.mjs finds, kickstarts and boots out this exact label; a mismatch
  // would leave setup.mjs unable to reload the Gateway.
  assert.match(plistContent(base), /<string>com\.hermes\.caddy<\/string>/);
});

test("plistContent requests restart-on-exit and start-at-login", () => {
  const out = plistContent(base);
  assert.match(out, /<key>KeepAlive<\/key>\s*<true\/>/);
  assert.match(out, /<key>RunAtLoad<\/key>\s*<true\/>/);
});

test("plistContent sends both streams to the log directory", () => {
  const out = plistContent(base);
  assert.match(out, /<string>\/Users\/x\/Library\/Logs\/hermes-office\/caddy\.log<\/string>/);
  assert.match(out, /<string>\/Users\/x\/Library\/Logs\/hermes-office\/caddy\.error\.log<\/string>/);
});

test("a path containing XML metacharacters cannot break the plist", () => {
  const out = plistContent({ ...base, repoRoot: "/Users/x/a&b<c" });
  assert.match(out, /\/Users\/x\/a&amp;b&lt;c\/Caddyfile/);
  assert.equal(out.includes("a&b<c"), false);
});

test("xmlEscape handles the three characters that matter", () => {
  assert.equal(xmlEscape("a&b<c>d"), "a&amp;b&lt;c&gt;d");
});
