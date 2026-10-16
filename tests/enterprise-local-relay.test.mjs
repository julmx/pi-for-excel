import assert from "node:assert/strict";
import http from "node:http";
import { once } from "node:events";
import { test } from "node:test";
import { createEnterpriseLocalRelay } from "../scripts/enterprise-local-relay.mjs";

async function start(server) {
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  return server.address().port;
}

test("local relay streams to the fixed API and rejects other destinations and browser origins", async (t) => {
  const seen = [];
  let redirect = false;
  const upstream = http.createServer(async (req, res) => {
    let body = "";
    for await (const chunk of req) body += chunk;
    seen.push({ path: req.url, method: req.method, headers: req.headers, body });
    if (redirect) { res.writeHead(302, { location: "https://attacker.invalid/" }); res.end(); return; }
    res.writeHead(200, { "content-type": "text/event-stream", "set-cookie": "unwanted=value" });
    res.write("data: first\n\n");
    setTimeout(() => res.end("data: [DONE]\n\n"), 10);
  });
  const upstreamPort = await start(upstream);
  t.after(() => { upstream.closeAllConnections(); upstream.close(); });
  let relay;
  const server = http.createServer(async (req, res) => {
    if (!await relay(req, res)) { res.writeHead(404); res.end(); }
  });
  const port = await start(server);
  t.after(() => { server.closeAllConnections(); server.close(); });
  const origin = `https://127.0.0.1:${port}`;
  relay = createEnterpriseLocalRelay({
    apiBaseUrl: "https://approved.example.invalid/llm/v1",
    localOrigin: origin,
    request: (url, options, callback) => {
      assert.equal(url.origin, "https://approved.example.invalid");
      // Transport boundary: exercise HTTP against a local mock of the external API.
      return http.request(`http://127.0.0.1:${upstreamPort}${url.pathname}`, options, callback);
    },
  });
  const call = (path, { method = "GET", headers = {}, body = "" } = {}) => new Promise((resolve, reject) => {
    const req = http.request(`http://127.0.0.1:${port}${path}`, { method, headers }, (res) => {
      let result = "";
      res.on("data", (chunk) => { result += chunk; });
      res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body: result }));
      res.on("error", reject);
    });
    req.on("error", reject);
    req.end(body);
  });
  const payload = JSON.stringify({ model: "synthetic-model", messages: [], stream: true });
  const response = await call("/__enterprise-api/chat/completions", {
    method: "POST",
    headers: { origin, "content-type": "application/json", authorization: "Bearer synthetic-test", cookie: "do-not-forward=1" },
    body: payload,
  });
  assert.equal(response.status, 200);
  assert.equal(response.body, "data: first\n\ndata: [DONE]\n\n");
  assert.equal(response.headers["set-cookie"], undefined);
  assert.equal(response.headers["cache-control"], "no-store");
  assert.equal(seen[0].path, "/llm/v1/chat/completions");
  assert.equal(seen[0].headers.authorization, "Bearer synthetic-test");
  assert.equal(seen[0].headers.cookie, undefined);
  assert.equal(seen[0].headers.origin, undefined);
  assert.equal(seen[0].body, payload);
  assert.equal((await call("/__enterprise-api/models")).status, 200);
  const allowedCalls = seen.length;
  for (const [path, options, status] of [
    ["/__enterprise-api/models?target=https://attacker.invalid", {}, 404],
    ["/__enterprise-api/../models", {}, 404],
    ["/__enterprise-api/other", {}, 404],
    ["/__enterprise-api/models", { method: "DELETE" }, 404],
    ["/__enterprise-api/models", { headers: { origin: "https://attacker.invalid" } }, 403],
    ["/__enterprise-api/models", { headers: { host: "attacker.invalid" } }, 403],
    ["/__enterprise-api/models", { headers: { "sec-fetch-site": "cross-site" } }, 403],
    ["/__enterprise-api/chat/completions", { method: "POST", headers: { "content-type": "application/json" }, body: payload }, 403],
    ["/__enterprise-api/chat/completions", { method: "POST", headers: { origin, "content-type": "text/plain" }, body: payload }, 403],
  ]) {
    assert.equal((await call(path, options)).status, status, path);
  }
  assert.equal(seen.length, allowedCalls);
  assert.equal((await call("/__enterprise-api/chat/completions", {
    method: "POST",
    headers: { origin, "content-type": "application/json" },
    body: "x".repeat(8 * 1024 * 1024 + 1),
  })).status, 413);
  assert.equal(seen.length, allowedCalls);
  redirect = true;
  assert.equal((await call("/__enterprise-api/models")).status, 502);
  upstream.closeAllConnections();
  await new Promise((resolve) => upstream.close(resolve));
  assert.equal((await call("/__enterprise-api/models")).status, 502);
});

test("local relay rejects unsafe configuration", () => {
  for (const apiBaseUrl of ["http://approved.invalid/v1", "https://user:secret@approved.invalid/v1", "https://approved.invalid/v1?q=1", "https://approved.invalid/v1#fragment"]) {
    assert.throws(() => createEnterpriseLocalRelay({ apiBaseUrl, localOrigin: "https://127.0.0.1:3141" }));
  }
  assert.throws(() => createEnterpriseLocalRelay({ apiBaseUrl: "https://approved.invalid/v1", localOrigin: "https://0.0.0.0:3141" }));
});
