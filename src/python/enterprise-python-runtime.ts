/**
 * One-shot enterprise Python session. The main thread owns all Excel calls;
 * the worker receives only serialized results. No native Python bridge.
 */
import type { EnterpriseExcelOperation, EnterpriseExcelOperationResponse } from "./enterprise-excel-bridge.js";

interface RunMessage { type: "run"; id: string; code: string }
interface ExcelCallMessage { type: "excel-call"; id: string; callId: string; request: EnterpriseExcelOperation }
interface FinishedMessage { type: "done"; id: string; ok: boolean; resultJson?: string; error?: string }
type WorkerMessage = ExcelCallMessage | FinishedMessage;

export interface EnterprisePythonResult {
  ok: boolean;
  resultJson?: string;
  error?: string;
}

export function runEnterprisePython(
  code: string,
  handleExcel: (request: EnterpriseExcelOperation, signal?: AbortSignal) => Promise<EnterpriseExcelOperationResponse>,
  signal?: AbortSignal,
): Promise<EnterprisePythonResult> {
  if (signal?.aborted) return Promise.resolve({ ok: false, error: "Aborted" });

  return new Promise((resolve) => {
    const worker = new Worker(new URL("./enterprise-python-worker.ts", import.meta.url), { type: "module" });
    const run: RunMessage = { type: "run", id: crypto.randomUUID(), code };
    let done = false;
    const timeout = setTimeout(() => finish({ ok: false, error: "Python timed out." }), 120_000);
    const abort = () => finish({ ok: false, error: "Aborted" });

    function finish(response: EnterprisePythonResult): void {
      if (done) return;
      done = true;
      clearTimeout(timeout);
      signal?.removeEventListener("abort", abort);
      worker.terminate();
      resolve(response);
    }

    worker.addEventListener("error", (event) => {
      finish({ ok: false, error: `Python worker failed: ${event.message}` });
    });
    worker.addEventListener("message", (event: MessageEvent<WorkerMessage>) => {
      const message = event.data;
      if (message.id !== run.id || done) return;
      if (message.type === "done") {
        finish({
          ok: message.ok,
          ...(message.resultJson !== undefined ? { resultJson: message.resultJson } : {}),
          ...(message.error !== undefined ? { error: message.error } : {}),
        });
        return;
      }
      void handleExcel(message.request, signal).then(
        (response) => {
          if (!done) worker.postMessage({
            type: "excel-result",
            callId: message.callId,
            response: JSON.stringify(response),
          });
        },
        (error: Error) => {
          if (!done) worker.postMessage({
            type: "excel-result",
            callId: message.callId,
            response: JSON.stringify({
              id: message.callId,
              ok: false,
              error: error.message,
            }),
          });
        },
      );
    });
    signal?.addEventListener("abort", abort, { once: true });
    worker.postMessage(run);
  });
}
