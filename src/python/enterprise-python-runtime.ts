/**
 * One-shot enterprise Python session. The main thread owns all Excel calls;
 * the worker receives only serialized results. No native Python bridge.
 */
import type { EnterpriseExcelOperation, EnterpriseExcelOperationResponse } from "./enterprise-excel-bridge.js";

interface RunMessage { type: "run"; id: string; code: string }
interface ExcelCallMessage { type: "excel-call"; id: string; callId: string; request: EnterpriseExcelOperation }
interface FinishedMessage { type: "done"; id: string; ok: boolean; resultJson?: string; stdout?: string; stderr?: string; error?: string }
interface StageMessage { type: "stage"; id: string; stage: string }
type WorkerMessage = ExcelCallMessage | FinishedMessage | StageMessage;

export interface EnterprisePythonResult {
  ok: boolean;
  resultJson?: string;
  stdout?: string;
  stderr?: string;
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
    let lastStage = "worker startup";
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
      const description = event instanceof ErrorEvent
        ? event.message || "No error details supplied by Excel WebView2"
        : "No error details supplied by Excel WebView2";
      const source = event instanceof ErrorEvent && event.filename
        ? ` in ${event.filename}:${event.lineno}`
        : "";
      finish({
        ok: false,
        error: `Python Worker failed during ${lastStage}: ${description}${source}. Check the local server logs for any failed asset requests.`,
      });
    });
    worker.addEventListener("messageerror", () => {
      finish({ ok: false, error: `Python Worker message could not be decoded during ${lastStage}.` });
    });
    worker.addEventListener("message", (event: MessageEvent<WorkerMessage>) => {
      const message = event.data;
      if (message.id !== run.id || done) return;
      if (message.type === "stage") {
        lastStage = message.stage;
        return;
      }
      if (message.type === "done") {
        finish({
          ok: message.ok,
          ...(message.resultJson !== undefined ? { resultJson: message.resultJson } : {}),
          ...(message.stdout !== undefined ? { stdout: message.stdout } : {}),
          ...(message.stderr !== undefined ? { stderr: message.stderr } : {}),
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
