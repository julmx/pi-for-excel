#!/usr/bin/env node
/**
 * Production-bundle browser smoke test for the enterprise Python/Excel Worker.
 * It drives the REAL Vite-built module worker, imports self-hosted Pyodide over
 * HTTP, and exchanges read/create/write RPC messages with a mocked Excel host.
 *
 * This is a runtime integration test, not an Office.js or private API test.
 */
import assert from "node:assert/strict";
import { createServer } from "node:http";
import { readFile, readdir } from "node:fs/promises";
import path from "node:path";
import { chromium } from "playwright";
import { buildEnterpriseContentSecurityPolicy } from "./enterprise-csp.mjs";

const dist = path.resolve("dist");
const assets = await readdir(path.join(dist, "assets"));
const workerFile = assets.find((name) => /^enterprise-python-worker-[\w-]+\.js$/u.test(name));
assert.ok(workerFile, "Vite build did not emit the enterprise Python Worker.");

// Use the deployed enterprise CSP, including the more restrictive policy
// applied to the Worker script, rather than a permissive test-only policy.
const enterpriseCsp = buildEnterpriseContentSecurityPolicy("https://api.example.invalid/llm/v1");
const workerCsp = enterpriseCsp
  .replace(/connect-src[^;]*/u, "connect-src 'self'")
  .replace(/worker-src[^;]*/u, "worker-src 'none'");

const mime = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".wasm": "application/wasm",
  ".json": "application/json; charset=utf-8",
  ".zip": "application/zip",
  ".whl": "application/octet-stream",
};

const server = createServer(async (req, res) => {
  try {
    const uri = new URL(req.url ?? "/", "http://127.0.0.1");
    if (uri.pathname === "/__smoke.html") {
      res.writeHead(200, {
        "content-type": mime[".html"],
        "content-security-policy": enterpriseCsp,
      });
      res.end("<!doctype html><title>Enterprise Python integration test</title>");
      return;
    }
    const relative = decodeURIComponent(uri.pathname).replace(/^\/+/, "");
    const target = path.resolve(dist, relative);
    if (!target.startsWith(dist + path.sep)) throw new Error("Invalid path.");
    const data = await readFile(target);
    res.writeHead(200, {
      "content-type": mime[path.extname(target)] ?? "application/octet-stream",
      "content-security-policy":
        /^\/assets\/enterprise-python-worker-[\w-]+\.js$/u.test(uri.pathname)
          ? workerCsp
          : enterpriseCsp,
    });
    res.end(data);
  } catch {
    res.writeHead(404);
    res.end("not found");
  }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const address = server.address();
assert.ok(address && typeof address === "object");
const baseUrl = `http://127.0.0.1:${address.port}`;

let browser;
try {
  browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  await context.route("**/*", async (route) => {
    const url = new URL(route.request().url());
    if (url.origin !== baseUrl) {
      await route.abort("blockedbyclient");
      return;
    }
    await route.continue();
  });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto(baseUrl + "/__smoke.html");

  const outcome = await page.evaluate((file) => new Promise((resolve, reject) => {
    const worker = new Worker("/assets/" + file, { type: "module" });
    const requests = [];
    const timer = setTimeout(() => {
      worker.terminate();
      reject(new Error("Python Worker integration test timed out."));
    }, 120_000);
    const finish = (value) => {
      clearTimeout(timer);
      worker.terminate();
      resolve({ ...value, requests });
    };
    worker.addEventListener("error", (error) => {
      clearTimeout(timer);
      worker.terminate();
      reject(new Error(`Python Worker startup failed: ${error.message}`));
    });
    worker.addEventListener("message", (event) => {
      const message = event.data;
      if (message.type === "excel-call") {
        requests.push({ tool: message.request.tool, params: message.request.params });
        let response;
        switch (message.request.tool) {
          case "read_range":
            response = { id: message.callId, ok: true, result: {
              details: { kind: "read_range_csv", values: [
                ["Group A", "", "Group B"],
                ["Name", "City", "Status"],
              ] },
            } };
            break;
          case "modify_structure":
          case "write_cells":
            response = { id: message.callId, ok: true, result: {
              content: [{ type: "text", text: "Synthetic Excel operation completed." }],
            } };
            break;
          default:
            response = { id: message.callId, ok: false, error: "Unexpected Excel operation." };
        }
        worker.postMessage({
          type: "excel-result",
          callId: message.callId,
          response: JSON.stringify(response),
        });
        return;
      }
      if (message.type === "done") finish(message);
    });
    worker.postMessage({
      type: "run",
      id: "enterprise-python-smoke",
      code: `import pandas as pd
import sys
print("Python stdout visible")
print("Python stderr visible", file=sys.stderr)
rows = await excel.read("Source!A1:C2")
groups = pd.Series(rows[0]).replace("", None).ffill().tolist()
headers = [f"{group} - {name}" for group, name in zip(groups, rows[1])]
await excel.create_sheet("PythonResult")
await excel.write("PythonResult!A1", [headers])
result = headers`,
    });
  }), workerFile);

  assert.equal(outcome.ok, true, outcome.error ?? "Python Worker returned an error");
  assert.deepEqual(JSON.parse(outcome.resultJson), [
    "Group A - Name",
    "Group A - City",
    "Group B - Status",
  ]);
  assert.equal(outcome.stdout?.trim(), "Python stdout visible");
  assert.equal(outcome.stderr?.trim(), "Python stderr visible");
  assert.deepEqual(outcome.requests.map((entry) => entry.tool), [
    "read_range",
    "modify_structure",
    "write_cells",
  ]);
  assert.deepEqual(outcome.requests[2].params.values, [[
    "Group A - Name",
    "Group A - City",
    "Group B - Status",
  ]]);
  // A second, independent Worker must also propagate Python exceptions and
  // the print output that preceded them, without using the Excel bridge.
  const failed = await page.evaluate((file) => new Promise((resolve, reject) => {
    const worker = new Worker("/assets/" + file, { type: "module" });
    const timeout = setTimeout(() => {
      worker.terminate();
      reject(new Error("Python error propagation test timed out."));
    }, 120_000);
    worker.addEventListener("error", (error) => {
      clearTimeout(timeout);
      worker.terminate();
      reject(new Error(`Worker error: ${error.message}`));
    });
    worker.addEventListener("message", (event) => {
      if (event.data.type !== "done") return;
      clearTimeout(timeout);
      worker.terminate();
      resolve(event.data);
    });
    worker.postMessage({
      type: "run",
      id: "enterprise-python-error-smoke",
      code: 'print("before failure")\nraise ValueError("synthetic exception")',
    });
  }), workerFile);
  assert.equal(failed.ok, false);
  assert.match(failed.error ?? "", /ValueError.*synthetic exception/u);
  assert.equal(failed.stdout?.trim(), "before failure");
  assert.deepEqual(errors, []);
  console.log("Browser smoke passed: Python stdout/stderr, errors, Excel RPC and results.");
} finally {
  await browser?.close();
  await new Promise((resolve) => server.close(resolve));
}
