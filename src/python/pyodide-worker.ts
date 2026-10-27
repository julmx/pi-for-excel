/**
 * Pyodide Web Worker.
 *
 * Runs in a dedicated worker thread. Loads Pyodide from CDN, executes
 * Python code, and posts results back to the main thread.
 *
 * Message protocol:
 * - Main → Worker: PyodideWorkerRequest
 * - Worker → Main: PyodideWorkerResponse
 */

import { executePyodideCode } from "./pyodide-execution.js";

// Worker globals — declared locally to avoid tsconfig lib conflicts
declare const self: {
  addEventListener(type: "message", listener: (event: MessageEvent) => void): void;
  postMessage(message: unknown): void;
};

export interface PyodideWorkerRequest {
  id: string;
  type: "run";
  code: string;
  inputJson?: string;
  packages?: string[];
}

export interface PyodideWorkerResponse {
  id: string;
  ok: boolean;
  stdout?: string;
  stderr?: string;
  resultJson?: string;
  error?: string;
  loadTimeMs?: number;
  runTimeMs?: number;
}

// Pyodide types (minimal subset for the worker)
interface PyodideInterface {
  runPythonAsync(code: string): Promise<unknown>;
  loadPackage(names: string | string[]): Promise<void>;
  /** Load the Pyodide-distribution packages the code imports (numpy, pandas, scipy, …). */
  loadPackagesFromImports(code: string): Promise<void>;
  globals: {
    set(name: string, value: unknown): void;
    get(name: string): unknown;
    delete(name: string): void;
  };
  toPy(value: unknown): unknown;
  version: string;
  FS: {
    writeFile(path: string, data: string | Uint8Array): void;
    readFile(path: string, options?: { encoding: string }): string;
  };
}

declare function _loadPyodide(options?: {
  indexURL?: string;
  stdout?: (text: string) => void;
  stderr?: (text: string) => void;
}): Promise<PyodideInterface>;

const PYODIDE_CDN_URL = "https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";

let pyodide: PyodideInterface | null = null;
let loadingPromise: Promise<PyodideInterface> | null = null;

async function ensurePyodide(): Promise<PyodideInterface> {
  if (pyodide) return pyodide;
  if (loadingPromise) return loadingPromise;

  loadingPromise = (async () => {
    try {
      // Dynamic import from CDN — works in module workers (no importScripts)
      const mod = await import(/* @vite-ignore */ `${PYODIDE_CDN_URL}pyodide.mjs`) as {
        loadPyodide: typeof _loadPyodide;
      };

      const instance = await mod.loadPyodide({
        indexURL: PYODIDE_CDN_URL,
      });

      pyodide = instance;
      return instance;
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      throw new Error(
        "Failed to load Pyodide from jsDelivr. " +
        "Check internet access and Content-Security-Policy allowlists for https://cdn.jsdelivr.net. " +
        `Original error: ${reason}`,
      );
    } finally {
      if (!pyodide) {
        loadingPromise = null;
      }
    }
  })();

  return loadingPromise;
}

/**
 * Detect import statements and return package names that may need
 * micropip.install(). Only handles top-level imports.
 */
function detectImports(code: string): string[] {
  const imports = new Set<string>();
  const lines = code.split("\n");

  for (const line of lines) {
    const trimmed = line.trim();
    // import foo / import foo.bar
    const importMatch = /^import\s+([\w.]+)/.exec(trimmed);
    if (importMatch) {
      const importName = importMatch[1];
      if (importName) {
        const topLevel = importName.split(".")[0];
        if (topLevel) imports.add(topLevel);
      }
    }
    // from foo import bar / from foo.bar import baz
    const fromMatch = /^from\s+([\w.]+)\s+import/.exec(trimmed);
    if (fromMatch) {
      const importName = fromMatch[1];
      if (importName) {
        const topLevel = importName.split(".")[0];
        if (topLevel) imports.add(topLevel);
      }
    }
  }

  // Remove stdlib and Pyodide-bundled packages
  const SKIP = new Set([
    // Python stdlib (common ones)
    "sys", "os", "io", "re", "json", "math", "csv", "datetime",
    "collections", "itertools", "functools", "operator", "string",
    "textwrap", "unicodedata", "struct", "copy", "pprint",
    "typing", "types", "abc", "enum", "dataclasses",
    "pathlib", "tempfile", "shutil", "glob",
    "hashlib", "hmac", "secrets", "base64", "binascii",
    "html", "xml", "email", "urllib", "http",
    "logging", "warnings", "traceback",
    "time", "calendar", "random", "statistics",
    "decimal", "fractions",
    "threading", "multiprocessing", "subprocess", "signal",
    "socket", "ssl", "select",
    "sqlite3", "dbm",
    "gzip", "bz2", "zipfile", "tarfile", "lzma", "zlib",
    "unittest", "doctest",
    "argparse", "configparser", "getopt",
    "ctypes", "array", "weakref", "gc",
    "inspect", "dis", "ast", "code", "codeop",
    "pickle", "shelve", "marshal",
    "platform", "sysconfig", "site",
    // Pyodide distribution packages (pre-compiled WASM): not preloaded, but
    // loadPackagesFromImports loads them, so micropip must not try PyPI for them.
    "numpy", "np",
    "pandas", "pd",
    "scipy",
    "matplotlib", "mpl", "plt",
    "seaborn", "sns",
    "sklearn", "scikit_learn",
    "PIL", "Pillow",
    "lxml",
    "bs4", "beautifulsoup4",
    "pyodide", "micropip",
    "js",
  ]);

  return [...imports].filter((pkg) => !SKIP.has(pkg));
}

async function handleRun(request: PyodideWorkerRequest): Promise<PyodideWorkerResponse> {
  const loadStart = performance.now();
  const py = await ensurePyodide();
  const loadTimeMs = Math.round(performance.now() - loadStart);

  // Load explicitly requested packages
  if (request.packages && request.packages.length > 0) {
    await py.loadPackage(request.packages);
  }

  // Load distribution packages the code imports (numpy, pandas, scipy, …).
  // If this fails the import itself raises a clear ModuleNotFoundError.
  try {
    await py.loadPackagesFromImports(request.code);
  } catch {
    // ignore: reported by the code's own import
  }

  // Auto-detect and install other missing packages via micropip
  const detectedPackages = detectImports(request.code);
  if (detectedPackages.length > 0) {
    try {
      await py.loadPackage("micropip");
      const packageList = detectedPackages.map((p) => `"${p}"`).join(", ");
      await py.runPythonAsync(`
import micropip
try:
    await micropip.install([${packageList}])
except Exception:
    pass
`);
    } catch {
      // Package install failed — continue anyway, the code itself will
      // produce a clear ImportError if the package is truly needed.
    }
  }

  // Set up or clear input_data between requests to prevent data leakage
  if (request.inputJson) {
    py.globals.set("__input_json__", request.inputJson);
    await py.runPythonAsync(`
import json as __json__
input_data = __json__.loads(__input_json__)
del __input_json__
`);
  } else {
    await py.runPythonAsync("input_data = None");
  }

  const runStart = performance.now();
  const output = await executePyodideCode(py, request.code);
  const response: PyodideWorkerResponse = {
    id: request.id,
    ok: output.error === undefined,
    loadTimeMs,
    runTimeMs: Math.round(performance.now() - runStart),
  };
  if (output.error !== undefined) response.error = output.error;
  if (output.stdout !== undefined) response.stdout = output.stdout;
  if (output.stderr !== undefined) response.stderr = output.stderr;
  if (output.resultJson !== undefined) response.resultJson = output.resultJson;
  return response;
}

self.addEventListener("message", (event: MessageEvent<PyodideWorkerRequest>) => {
  const request = event.data;

  void handleRun(request).then(
    (response) => self.postMessage(response),
    (error: unknown) => {
      const message = error instanceof Error ? error.message : String(error);
      const response: PyodideWorkerResponse = {
        id: request.id,
        ok: false,
        error: `Worker error: ${message}`,
      };
      self.postMessage(response);
    },
  );
});
