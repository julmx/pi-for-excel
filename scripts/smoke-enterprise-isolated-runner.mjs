#!/usr/bin/env node
/* Synthetic-only isolation test of the production-built enterprise runner. */
import assert from "node:assert/strict";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { chromium } from "playwright";
import { buildEnterpriseContentSecurityPolicy, buildEnterprisePythonRunnerCsp } from "./enterprise-csp.mjs";

const port = 3141;
const parent = "http://127.0.0.1:" + port;
const runner = "http://localhost:" + (port + 1);
const canaryName = "__pi_synthetic_worker_isolation_probe";
const parentCsp = buildEnterpriseContentSecurityPolicy("https://api.example.invalid/v1", runner);
const runnerCsp = buildEnterprisePythonRunnerCsp().replace("https://127.0.0.1:3141", parent);
const workerCsp = buildEnterprisePythonRunnerCsp(true).replace("https://127.0.0.1:3141", parent);
const dist = path.resolve("dist");

const serve = async (req, res) => {
  const isolated = req.headers.host === "localhost:" + (port + 1)
    && req.socket.localPort === port + 1;
  const pathname = new URL(req.url ?? "/", parent).pathname;
  if (isolated && !(pathname === "/enterprise-runner.html"
      || pathname.startsWith("/assets/") || pathname.startsWith("/enterprise-pyodide/"))) {
    res.writeHead(403);
    res.end("runner resource denied");
    return;
  }
  if (pathname === "/__smoke.html") {
    res.writeHead(200, { "content-type": "text/html; charset=utf-8",
      "content-security-policy": parentCsp });
    res.end("<!doctype html><html><body></body></html>");
    return;
  }
  try {
    const file = path.resolve(dist, pathname.replace(/^\/+/u, ""));
    if (!file.startsWith(dist + path.sep)) throw new Error("invalid path");
    const data = await readFile(file);
    const type = {
      ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8",
      ".mjs": "text/javascript; charset=utf-8", ".wasm": "application/wasm",
      ".json": "application/json; charset=utf-8", ".zip": "application/zip",
      ".whl": "application/octet-stream",
    }[path.extname(file)] ?? "application/octet-stream";
    res.writeHead(200, {
      "content-type": type,
      "content-security-policy": isolated
        ? /^\/assets\/enterprise-python-worker-[\w-]+\.js$/u.test(pathname)
          ? workerCsp : runnerCsp : parentCsp,
    });
    res.end(data);
  } catch {
    res.writeHead(404);
    res.end("not found");
  }
};
const server = createServer(serve);
const runnerServer = createServer(serve);
await Promise.all([
  new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  }),
  new Promise((resolve, reject) => {
    runnerServer.once("error", reject);
    runnerServer.listen(port + 1, "127.0.0.1", resolve);
  }),
]);
let browser;
try {
  browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  await context.route("**/*", (route) => {
    const origin = new URL(route.request().url()).origin;
    if (origin !== parent && origin !== runner) {
      void route.abort("blockedbyclient");
    } else {
      void route.continue();
    }
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(parent + "/__smoke.html");
  const test = await page.evaluate(async ({ runner, canaryName }) => {
    const opened = indexedDB.open(canaryName, 1);
    opened.onupgradeneeded = () => opened.result.createObjectStore("synthetic");
    const db = await new Promise((resolve, reject) => {
      opened.onsuccess = () => resolve(opened.result);
      opened.onerror = () => reject(opened.error);
    });
    try {
      await new Promise((resolve, reject) => {
        const transaction = db.transaction("synthetic", "readwrite");
        transaction.objectStore("synthetic").put("SYNTHETIC_NOT_SECRET", "value");
        transaction.oncomplete = resolve;
        transaction.onerror = () => reject(transaction.error);
      });
    } finally {
      db.close();
    }

    const python = [
      "import pandas as pd",
      "from js import indexedDB, location",
      "dbs = await indexedDB.databases()",
      "result = {'runner_origin': str(location.origin), 'sees_parent_database': any(str(db.name) == '__pi_synthetic_worker_isolation_probe' for db in dbs)}",
      "rows = await excel.read('Source!A1:B2')",
      "df = pd.DataFrame(rows[1:], columns=rows[0])",
      "await excel.create_sheet('Result')",
      "await excel.write('Result!A1', [list(df.columns)] + df.values.tolist())",
      "result['rows'] = len(df)",
    ].join("\n");

    const requests = [];
    try {
      const response = await new Promise((resolve, reject) => {
        const frame = document.createElement("iframe");
        frame.setAttribute("sandbox", "allow-scripts allow-same-origin");
        frame.style.display = "none";
        frame.src = runner + "/enterprise-runner.html";
        const id = crypto.randomUUID();
        const nonce = crypto.randomUUID();
        const channel = new MessageChannel();
        const port = channel.port1;
        const cleanup = () => {
          clearTimeout(timeout);
          port.close();
          frame.remove();
        };
        const timeout = setTimeout(() => {
          cleanup();
          reject(new Error("Isolated runner timed out"));
        }, 120_000);
        port.addEventListener("message", (event) => {
          const m = event.data;
          if (m?.id !== id || m.nonce !== nonce) return;
          if (m.type === "ready") {
            port.postMessage({ type: "run", id, nonce, payload: { type: "run", id, code: python } });
          } else if (m.type === "worker-error") {
            cleanup();
            reject(new Error(m.error));
          } else if (m.type === "forward" && m.payload.type === "excel-call") {
            const payload = m.payload;
            requests.push(payload.request.tool);
            const tool = payload.request.tool;
            const result = tool === "read_range"
              ? { details: { kind: "read_range_csv", values: [["Name", "Amount"], ["Example", 42]] } }
              : { content: [{ type: "text", text: "Synthetic Excel result" }] };
            const permitted = ["read_range", "modify_structure", "write_cells"].includes(tool);
            port.postMessage({
              type: "excel-result", id, nonce,
              payload: { type: "excel-result", callId: payload.callId,
                response: JSON.stringify(permitted
                  ? { id: payload.callId, ok: true, result }
                  : { id: payload.callId, ok: false, error: "Not permitted" }) },
            });
          } else if (m.type === "forward" && m.payload.type === "done") {
            cleanup();
            resolve(m.payload);
          }
        });
        port.start();
        frame.addEventListener("load", () => {
          frame.contentWindow.postMessage(
            { type: "enterprise-python-connect", id, nonce }, runner, [channel.port2],
          );
        });
        document.body.append(frame);
      });
      return { response, requests };
    } finally {
      await new Promise((resolve, reject) => {
        const deletion = indexedDB.deleteDatabase(canaryName);
        deletion.onsuccess = resolve;
        deletion.onerror = () => reject(deletion.error);
      });
    }
  }, { runner, canaryName });

  assert.equal(test.response.ok, true, test.response.error);
  assert.deepEqual(JSON.parse(test.response.resultJson), {
    runner_origin: runner, sees_parent_database: false, rows: 1,
  });
  assert.deepEqual(test.requests, ["read_range", "modify_structure", "write_cells"]);
  assert.deepEqual(errors, []);
  console.log("Isolated enterprise runner: taskpane IndexedDB denied, pandas + Excel RPC passed.");
} finally {
  await browser?.close();
  await Promise.all([
    new Promise((resolve) => server.close(resolve)),
    new Promise((resolve) => runnerServer.close(resolve)),
  ]);
}
