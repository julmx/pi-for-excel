/**
 * Download a pinned Pyodide runtime and its NumPy/pandas wheels at BUILD time.
 * The Excel taskpane serves them from its own HTTPS origin at runtime.
 * Not used in the regular upstream build.
 */
import { mkdir, readFile, writeFile } from "node:fs/promises";
import path from "node:path";

const VERSION = "0.27.7";
const SOURCE = `https://cdn.jsdelivr.net/pyodide/v${VERSION}/full/`;
const TARGET = path.resolve("public/enterprise-pyodide");
// The Emscripten loader imports pyodide.asm.js at runtime. Omitting it
// passes a Vite build but fails when Python first starts in Excel.
const CORE = [
  "pyodide.mjs",
  "pyodide.asm.js",
  "pyodide.asm.wasm",
  "python_stdlib.zip",
  "pyodide-lock.json",
];

async function download(name) {
  const file = path.join(TARGET, name);
  try {
    if ((await readFile(file)).length > 0) return;
  } catch { /* Fetch missing file. */ }
  const response = await fetch(new URL(name, SOURCE), { redirect: "error" });
  if (!response.ok) throw new Error(`Pyodide download failed (${response.status}): ${name}`);
  await writeFile(file, new Uint8Array(await response.arrayBuffer()));
  console.log(`Pyodide asset: ${name}`);
}

if (process.env.VITE_PI_ENTERPRISE_MODE === "true") {
  await mkdir(TARGET, { recursive: true });
  for (const file of CORE) await download(file);
  const lock = JSON.parse(await readFile(path.join(TARGET, "pyodide-lock.json"), "utf8"));
  const selected = new Set();
  function include(name) {
    if (selected.has(name)) return;
    const info = lock.packages[name];
    if (!info) throw new Error(`Missing Pyodide package ${name}`);
    selected.add(name);
    for (const dependency of info.depends ?? []) include(dependency);
  }
  include("numpy");
  include("pandas");
  for (const name of selected) {
    const info = lock.packages[name];
    if (info.file_name) await download(info.file_name);
  }
  console.log(`Self-hosted Pyodide ${VERSION}: ${selected.size} packages ready`);
}
