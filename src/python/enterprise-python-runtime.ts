/**
 * Run enterprise Python through an iframe on an isolated HTTPS origin.
 *
 * A normal Web Worker shares its creator's origin and IndexedDB. The runner
 * page lives on localhost while Excel and its credentials live on 127.0.0.1.
 * No same-origin fallback: a missing runner must fail closed.
 */
import type { EnterpriseExcelOperation, EnterpriseExcelOperationResponse } from "./enterprise-excel-bridge.js";

const TASKPANE_ORIGIN = import.meta.env.VITE_PI_ENTERPRISE_TASKPANE_ORIGIN ?? "https://127.0.0.1:3141";
const RUNNER_ORIGIN = import.meta.env.VITE_PI_ENTERPRISE_RUNNER_ORIGIN ?? "https://localhost:3141";

interface RunMessage { type: "run"; id: string; code: string }
interface ExcelCallMessage { type: "excel-call"; id: string; callId: string; request: EnterpriseExcelOperation }
interface FinishedMessage { type: "done"; id: string; ok: boolean; resultJson?: string; stdout?: string; stderr?: string; error?: string }
interface StageMessage { type: "stage"; id: string; stage: string }
type WorkerMessage = ExcelCallMessage | FinishedMessage | StageMessage;

interface RunnerPortMessage {
  type: "ready" | "forward" | "worker-error";
  id: string;
  nonce: string;
  payload?: WorkerMessage;
  error?: string;
}

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
  if (window.location.origin !== TASKPANE_ORIGIN) {
    return Promise.resolve({ ok: false, error: "Enterprise Python requires an isolated, configured runner origin." });
  }

  return new Promise((resolve) => {
    const frame = document.createElement("iframe");
    frame.setAttribute("sandbox", "allow-scripts allow-same-origin");
    frame.setAttribute("aria-hidden", "true");
    frame.style.display = "none";
    frame.src = `${RUNNER_ORIGIN}/enterprise-runner.html`;

    const run: RunMessage = { type: "run", id: crypto.randomUUID(), code };
    const nonce = crypto.randomUUID();
    const channel = new MessageChannel();
    const port = channel.port1;
    let done = false;
    let ready = false;
    let lastStage = "isolated runner startup";
    const timeout = setTimeout(() => finish({ ok: false, error: "Python timed out." }), 120_000);
    const handshakeTimeout = setTimeout(() => finish({
      ok: false,
      error: "Isolated Python runner did not connect. Verify HTTPS localhost and the local server.",
    }), 15_000);
    const abort = (): void => finish({ ok: false, error: "Aborted" });

    function finish(response: EnterprisePythonResult): void {
      if (done) return;
      done = true;
      clearTimeout(timeout);
      clearTimeout(handshakeTimeout);
      signal?.removeEventListener("abort", abort);
      port.close();
      channel.port2.close();
      frame.remove();
      resolve(response);
    }

    port.addEventListener("message", (event: MessageEvent<RunnerPortMessage>) => {
      const message = event.data;
      if (!message || done || message.id !== run.id || message.nonce !== nonce) return;
      if (message.type === "worker-error") {
        finish({ ok: false, error: `Python Worker failed during ${lastStage}: ${message.error ?? "Unknown error"}` });
        return;
      }
      if (message.type === "ready" && !ready) {
        ready = true;
        clearTimeout(handshakeTimeout);
        port.postMessage({ type: "run", id: run.id, nonce, payload: run });
        return;
      }
      if (!ready || message.type !== "forward" || !message.payload) return;
      const payload = message.payload;
      if (payload.id !== run.id) return;
      if (payload.type === "stage") {
        lastStage = payload.stage;
        return;
      }
      if (payload.type === "done") {
        finish({
          ok: payload.ok,
          ...(payload.resultJson !== undefined ? { resultJson: payload.resultJson } : {}),
          ...(payload.stdout !== undefined ? { stdout: payload.stdout } : {}),
          ...(payload.stderr !== undefined ? { stderr: payload.stderr } : {}),
          ...(payload.error !== undefined ? { error: payload.error } : {}),
        });
        return;
      }

      void handleExcel(payload.request, signal).then(
        (response) => {
          if (!done) port.postMessage({
            type: "excel-result", id: run.id, nonce,
            payload: { type: "excel-result", callId: payload.callId, response: JSON.stringify(response) },
          });
        },
        (error: Error) => {
          if (!done) port.postMessage({
            type: "excel-result", id: run.id, nonce,
            payload: { type: "excel-result", callId: payload.callId,
              response: JSON.stringify({ id: payload.callId, ok: false, error: error.message }) },
          });
        },
      );
    });
    port.addEventListener("messageerror", () => {
      finish({ ok: false, error: "Isolated Python message could not be decoded." });
    });
    port.start();

    frame.addEventListener("load", () => {
      if (done) return;
      const receiver = frame.contentWindow;
      if (!receiver) {
        finish({ ok: false, error: "Isolated Python runner frame unavailable." });
        return;
      }
      receiver.postMessage({ type: "enterprise-python-connect", id: run.id, nonce },
        RUNNER_ORIGIN, [channel.port2]);
    });
    frame.addEventListener("error", () => {
      finish({ ok: false, error: "Isolated Python runner frame failed to load." });
    });
    signal?.addEventListener("abort", abort, { once: true });
    document.body.append(frame);
  });
}
