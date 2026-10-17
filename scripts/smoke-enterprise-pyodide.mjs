#!/usr/bin/env node
/**
 * Exercise the exact self-hosted Pyodide distribution shipped with the
 * enterprise Vite build. Unlike a build-only check this catches missing
 * pyodide.asm.js, WASM, standard-library and pandas/numpy wheel assets.
 */
import assert from "node:assert/strict";
import { access, readFile, stat } from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const root = path.resolve("dist/enterprise-pyodide");
const required = [
  "pyodide.mjs",
  "pyodide.asm.js",
  "pyodide.asm.wasm",
  "python_stdlib.zip",
  "pyodide-lock.json",
];
for (const name of required) {
  const file = path.join(root, name);
  await access(file);
  assert.ok((await stat(file)).size > 0, `Empty Pyodide asset: ${name}`);
}

const lock = JSON.parse(await readFile(path.join(root, "pyodide-lock.json"), "utf8"));
const wheels = new Set();
function visit(name) {
  if (wheels.has(name)) return;
  const item = lock.packages[name];
  assert.ok(item, `Missing pinned Pyodide package: ${name}`);
  wheels.add(name);
  for (const dep of item.depends ?? []) visit(dep);
}
visit("numpy");
visit("pandas");
for (const name of wheels) {
  const filename = lock.packages[name].file_name;
  if (filename) await access(path.join(root, filename));
}

const { loadPyodide } = await import(pathToFileURL(path.join(root, "pyodide.mjs")).href);
// Under Node.js Pyodide expects a filesystem directory here, whereas the
// browser uses the HTTPS index URL. A file:// index URL is treated as a
// relative module specifier by the Emscripten loader in Node.
const py = await loadPyodide({
  indexURL: root + path.sep,
});
await py.loadPackage(["numpy", "pandas"]);
const answer = await py.runPythonAsync(`
import pandas as pd
import numpy as np
df = pd.DataFrame({"name": ["alice", "bob"]})
result = list(df["name"].str.upper())
result == ["ALICE", "BOB"] and int(np.array([1, 2, 3]).sum()) == 6
`);
assert.equal(answer, true, "Self-hosted Python data science execution failed");
console.log("Enterprise Pyodide smoke test passed (interpreter + pandas + NumPy).");
