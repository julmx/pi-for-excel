/**
 * Shared Python execution semantics from the original Pi Pyodide Worker.
 *
 * Runtime loading and transport remain the responsibility of each Worker:
 * ordinary Pi uses its existing provider, while the enterprise Worker loads
 * only bundled Pyodide and exposes allowlisted Excel RPCs.
 */
export interface PyodideExecutor<T> {
  runPythonAsync(code: string): Promise<T>;
  globals: { delete(name: string): void };
}

export interface PythonExecutionOutput {
  stdout?: string;
  stderr?: string;
  resultJson?: string;
  error?: string;
}

export async function executePyodideCode<T>(
  py: PyodideExecutor<T>,
  code: string,
  stringifyUnknownResult = false,
): Promise<PythonExecutionOutput> {
  // Keep the upstream worker's familiar stdout/stderr and result behavior.
  await py.runPythonAsync(`
import sys as __sys__
import io as __io__
__stdout_capture__ = __io__.StringIO()
__stderr_capture__ = __io__.StringIO()
__sys__.stdout = __stdout_capture__
__sys__.stderr = __stderr_capture__
result = None
`);

  const output: PythonExecutionOutput = {};
  try {
    try {
      await py.runPythonAsync(code);
    } catch (error) {
      output.error = error instanceof Error ? error.message : String(error);
    }

    try {
      const stdout = await py.runPythonAsync("__stdout_capture__.getvalue()");
      if (typeof stdout === "string" && stdout.length > 0) output.stdout = stdout;
    } catch { /* Keep the original execution error. */ }

    try {
      const stderr = await py.runPythonAsync("__stderr_capture__.getvalue()");
      if (typeof stderr === "string" && stderr.length > 0) output.stderr = stderr;
    } catch { /* Keep the original execution error. */ }

    if (output.error === undefined) {
      try {
        const result = await py.runPythonAsync(stringifyUnknownResult
          ? "import json as __json__; __json__.dumps(result, default=str) if result is not None else None"
          : "import json as __json__; __json__.dumps(result) if result is not None else None");
        if (typeof result === "string") output.resultJson = result;
      } catch {
        // Non-JSON Python values are allowed; stdout remains available.
      }
    }
  } finally {
    await py.runPythonAsync(`
__sys__.stdout = __sys__.__stdout__
__sys__.stderr = __sys__.__stderr__
`).catch(() => { /* Avoid masking the execution error. */ });

    try {
      py.globals.delete("__stdout_capture__");
      py.globals.delete("__stderr_capture__");
    } catch { /* Runtime may already be shutting down. */ }
  }
  return output;
}
