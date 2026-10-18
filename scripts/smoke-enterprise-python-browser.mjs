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
  assert.match(failed.error ?? "", /ValueError.*synthetic exception/u, JSON.stringify(failed));
  assert.equal(failed.stdout?.trim(), "before failure");

  // Regression for PR #13: a first-party Excel write can be refused without
  // throwing in JavaScript. The bridge turns that into ok:false, and the real
  // enterprise Python Worker must raise RuntimeError to Python callers.
  // Unlike the ordinary successful-write mock above, exercise try/except.
  const blocked = await page.evaluate((file) => new Promise((resolve, reject) => {
    const worker = new Worker("/assets/" + file, { type: "module" });
    const requests = [];
    const timeout = setTimeout(() => {
      worker.terminate();
      reject(new Error("Blocked-write Python integration test timed out."));
    }, 120_000);
    worker.addEventListener("error", (error) => {
      clearTimeout(timeout);
      worker.terminate();
      reject(new Error(`Blocked-write Worker failed: ${error.message}`));
    });
    worker.addEventListener("message", (event) => {
      const message = event.data;
      if (message.type === "excel-call") {
        requests.push(message.request);
        worker.postMessage({
          type: "excel-result",
          callId: message.callId,
          response: JSON.stringify({
            id: message.callId,
            ok: false,
            error: "⛔ Write blocked — Source!A1 contains 1 non-empty cell(s).",
          }),
        });
        return;
      }
      if (message.type !== "done") return;
      clearTimeout(timeout);
      worker.terminate();
      resolve({ ...message, requests });
    });
    worker.postMessage({
      type: "run",
      id: "enterprise-python-blocked-write-smoke",
      code: `result = {"status": None, "error": None}
try:
    await excel.write("Source!A1", [["SHOULD_NOT_OVERWRITE"]])
    result["status"] = "NO_EXCEPTION"
except RuntimeError as error:
    result["status"] = "EXCEPTION_LEVEE"
    result["error"] = str(error)`,
    });
  }), workerFile);
  assert.equal(blocked.ok, true, JSON.stringify(blocked));
  assert.deepEqual(JSON.parse(blocked.resultJson), {
    status: "EXCEPTION_LEVEE",
    error: "⛔ Write blocked — Source!A1 contains 1 non-empty cell(s).",
  });
  assert.equal(blocked.requests.length, 1);
  assert.equal(blocked.requests[0].tool, "write_cells");
  assert.equal(blocked.requests[0].params.start_cell, "Source!A1");
  assert.deepEqual(blocked.requests[0].params.values, [["SHOULD_NOT_OVERWRITE"]]);
  // Characterization of the CURRENT security boundary, not an endorsement:
  // model-generated Python should not be assumed confined to the Excel RPC.
  // Create ONLY a synthetic origin-scoped database; never inspect user or
  // enterprise storage, credentials, sessions or workbook contents.
  const isolationProbe = await page.evaluate(async (file) => {
    const databaseName = "__pi_synthetic_worker_isolation_probe";
    const marker = "SYNTHETIC_CANARY_NOT_A_SECRET";
    const opened = indexedDB.open(databaseName, 1);
    opened.onupgradeneeded = () => {
      if (!opened.result.objectStoreNames.contains("synthetic")) {
        opened.result.createObjectStore("synthetic");
      }
    };
    const db = await new Promise((resolve, reject) => {
      opened.onsuccess = () => resolve(opened.result);
      opened.onerror = () => reject(opened.error ?? new Error("Synthetic DB open failed"));
    });
    try {
      await new Promise((resolve, reject) => {
        const transaction = db.transaction("synthetic", "readwrite");
        transaction.objectStore("synthetic").put(marker, "marker");
        transaction.oncomplete = () => resolve();
        transaction.onerror = () => reject(transaction.error ?? new Error("Synthetic write failed"));
      });
    } finally {
      db.close();
    }

    const python = [
      "import asyncio",
      "from js import indexedDB",
      "from pyodide.ffi import create_proxy",
      "async def _read_idb_request(request):",
      "    future = asyncio.get_running_loop().create_future()",
      "    def on_success(_event):",
      "        if not future.done(): future.set_result(request.result)",
      "    def on_error(_event):",
      "        if not future.done(): future.set_exception(RuntimeError('Synthetic IndexedDB request failed'))",
      "    success = create_proxy(on_success)",
      "    error = create_proxy(on_error)",
      "    request.addEventListener('success', success)",
      "    request.addEventListener('error', error)",
      "    try:",
      "        return await future",
      "    finally:",
      "        request.removeEventListener('success', success)",
      "        request.removeEventListener('error', error)",
      "        success.destroy()",
      "        error.destroy()",
      "_database = await _read_idb_request(indexedDB.open('__pi_synthetic_worker_isolation_probe'))",
      "try:",
      "    _transaction = _database.transaction('synthetic', 'readonly')",
      "    _value = await _read_idb_request(_transaction.objectStore('synthetic').get('marker'))",
      "    result = {'can_access_worker_indexeddb': True, 'synthetic_canary': str(_value)}",
      "finally:",
      "    _database.close()",
    ].join("\n");

    try {
      return await new Promise((resolve, reject) => {
        const worker = new Worker("/assets/" + file, { type: "module" });
        const timeout = setTimeout(() => {
          worker.terminate();
          reject(new Error("Synthetic worker isolation probe timed out."));
        }, 120_000);
        worker.addEventListener("error", (error) => {
          clearTimeout(timeout);
          worker.terminate();
          reject(new Error("Synthetic worker isolation probe crashed: " + error.message));
        });
        worker.addEventListener("message", (event) => {
          if (event.data.type !== "done") return;
          clearTimeout(timeout);
          worker.terminate();
          resolve(event.data);
        });
        worker.postMessage({ type: "run", id: "synthetic-worker-isolation-probe", code: python });
      });
    } finally {
      // Remove the *only* database created by this diagnostic. Deleting this
      // synthetic name never touches the application's "pi-for-excel" DB.
      await new Promise((resolve, reject) => {
        const request = indexedDB.deleteDatabase(databaseName);
        request.onsuccess = () => resolve();
        request.onerror = () => reject(request.error ?? new Error("Synthetic DB cleanup failed"));
        request.onblocked = () => reject(new Error("Synthetic DB cleanup was blocked"));
      });
    }
  }, workerFile);
  assert.equal(isolationProbe.ok, true, isolationProbe.error ?? "Synthetic isolation probe failed");
  // The current same-origin design exposes this browser storage capability.
  // Replace this characterization with a DENIAL assertion when isolation is
  // implemented; do not mistake this passing test for a security guarantee.
  assert.deepEqual(JSON.parse(isolationProbe.resultJson), {
    can_access_worker_indexeddb: true,
    synthetic_canary: "SYNTHETIC_CANARY_NOT_A_SECRET",
  });
  console.log("Security characterization: Python can read synthetic same-origin IndexedDB.");

  assert.deepEqual(errors, []);
  console.log("Browser smoke passed: Python stdout/stderr, errors, Excel RPC and results.");
} finally {
  await browser?.close();
  await new Promise((resolve) => server.close(resolve));
}
