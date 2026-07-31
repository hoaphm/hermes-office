import test from "node:test";
import assert from "node:assert/strict";
import net from "net";
import { gatewayPort, isPortBusy, DEFAULT_PORT } from "./gateway.mjs";

test("gatewayPort reads the port the Caddyfile actually binds", () => {
  assert.equal(gatewayPort("https://localhost:8643 {\n tls internal\n}"), 8643);
  assert.equal(gatewayPort("https://localhost:9000 {"), 9000);
});

test("gatewayPort falls back to the default when no port is declared", () => {
  assert.equal(gatewayPort("localhost {"), DEFAULT_PORT);
  assert.equal(gatewayPort(""), DEFAULT_PORT);
  assert.equal(gatewayPort(undefined), DEFAULT_PORT);
});

test("gatewayPort ignores a port mentioned in a redirect further down", () => {
  // The first localhost:PORT is the site address; later ones are prose or
  // redirects and must not win.
  const text = "https://localhost:8643 {\n  redir / https://localhost:1234/word/taskpane.html\n}";
  assert.equal(gatewayPort(text), 8643);
});

test("isPortBusy is true for a listening socket and false once it closes", async () => {
  const server = net.createServer();
  await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
  const port = server.address().port;
  assert.equal(await isPortBusy(port), true);
  await new Promise((resolve) => server.close(resolve));
  assert.equal(await isPortBusy(port), false);
});
