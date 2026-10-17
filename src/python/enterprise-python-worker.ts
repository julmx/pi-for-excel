/**
 * Enterprise-only Pyodide worker.
 * Python communicates with Excel exclusively through the typed RPC messages.
 * The Pyodide runtime and wheels are served from the taskpane origin.
 */
interface StartMessage { type: "run"; id: string; code: string }
interface ExcelResultMessage { type: "excel-result"; callId: string; response: string }
type Incoming = StartMessage | ExcelResultMessage;
type ExcelCallback = (response: string) => void;

declare const self: {
  addEventListener(type: "message", listener: (event: MessageEvent<Incoming>) => void): void;
  addEventListener(type: "error", listener: (event: ErrorEvent) => void): void;
  addEventListener(type: "unhandledrejection", listener: (event: PromiseRejectionEvent) => void): void;
  postMessage(message: object): void;
};

interface Runtime {
  loadPackage(names: string[]): Promise<void>;
  runPythonAsync(code: string): Promise<string | number | boolean | object | null | undefined>;
  globals: { set(key: string, value: object): void };
}

const pending = new Map<string, ExcelCallback>();
let activeRunId: string | null = null;

function announceStage(id: string, stage: string): void {
  self.postMessage({ type: "stage", id, stage });
}

let counter = 0;

async function run(request: StartMessage): Promise<void> {
  activeRunId = request.id;
  announceStage(request.id, "loading the Pyodide module");
  const localLoaderPath = "/enterprise-pyodide/pyodide.mjs";
  const module = await import(/* @vite-ignore */ localLoaderPath) as {
    loadPyodide(options: { indexURL: string }): Promise<Runtime>;
  };
  announceStage(request.id, "initializing WebAssembly");
  const py = await module.loadPyodide({
    indexURL: "/enterprise-pyodide/",
  });
  announceStage(request.id, "loading pandas and NumPy");
  await py.loadPackage(["numpy", "pandas"]);
  announceStage(request.id, "registering the Excel bridge");

  py.globals.set("_excel_rpc", async (tool: string, paramsJson: string): Promise<string> =>
    new Promise<string>((resolve) => {
      const callId = `${request.id}-${++counter}`;
      pending.set(callId, resolve);
      const params: unknown = JSON.parse(paramsJson);
      if (!params || typeof params !== "object" || Array.isArray(params)) {
        throw new Error("Python Excel request must be a JSON object.");
      }
      self.postMessage({
        type: "excel-call",
        id: request.id,
        callId,
        request: { id: callId, tool, params },
      });
    }));

  await py.runPythonAsync(`
import json as _excel_json
import io as _excel_io
import sys as _excel_sys

class _ExcelApi:
    async def call(self, tool, **params):
        response = _excel_json.loads(await _excel_rpc(tool, _excel_json.dumps(params, default=str)))
        if not response["ok"]:
            raise RuntimeError(response["error"])
        return response["result"]

    async def read(self, range):
        response = await self.call("read_range", range=range, mode="csv")
        details = response.get("details") or {}
        return details.get("values", [])

    async def write(self, start_cell, values, **options):
        return await self.call("write_cells", start_cell=start_cell, values=values, **options)

    async def create_sheet(self, name):
        return await self.call("modify_structure", action="add_sheet", new_name=name)

    async def format(self, range, **options):
        return await self.call("format_cells", range=range, **options)

    async def chart(self, **options):
        return await self.call("charts", **options)

    async def formula(self, **options):
        return await self.call("fill_formula", **options)

excel = _ExcelApi()
result = None
`);

  try {
    announceStage(request.id, "running the Python script");
    // Pyodide supports top-level await via runPythonAsync.
    await py.runPythonAsync(request.code);
    const output = await py.runPythonAsync(`
import json as _excel_json
_excel_result = result if result is not None else None
_excel_json.dumps(_excel_result, default=str)
`);
    if (typeof output !== "string") throw new Error("Python did not return a JSON string.");
    announceStage(request.id, "returning Python results");
    self.postMessage({ type: "done", id: request.id, ok: true, resultJson: output });
  } catch (error) {
    self.postMessage({
      type: "done",
      id: request.id,
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    });
  }
}

// Surface errors from asynchronously loaded runtime components in the same
// request protocol. No Python code, spreadsheet values or credentials logged.
self.addEventListener("unhandledrejection", (event) => {
  if (!activeRunId) return;
  const reason: unknown = event.reason;
  self.postMessage({
    type: "done",
    id: activeRunId,
    ok: false,
    error: `Unhandled Python Worker rejection: ${reason instanceof Error ? reason.message : String(reason)}`,
  });
});

self.addEventListener("error", (event) => {
  if (!activeRunId) return;
  self.postMessage({
    type: "done",
    id: activeRunId,
    ok: false,
    error: `Python Worker runtime error: ${event.message || "no details"}`,
  });
});

self.addEventListener("message", (event) => {
  const message = event.data;
  if (message.type === "excel-result") {
    const resolve = pending.get(message.callId);
    pending.delete(message.callId);
    resolve?.(message.response);
  } else if (message.type === "run") {
    void run(message).catch((error: Error) => {
      self.postMessage({ type: "done", id: message.id, ok: false, error: error.message });
    });
  }
});
