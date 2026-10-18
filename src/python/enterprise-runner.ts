/**
 * One-run, cross-origin relay for the existing enterprise Pyodide Worker.
 *
 * This page is ONLY served from https://localhost:3141 and framed by
 * https://127.0.0.1:3141. It has no provider credentials or Excel API.
 * A per-run MessagePort is the sole communication capability; only the
 * parent taskpane performs the allowlisted Excel tool calls.
 */
const TASKPANE_ORIGIN = import.meta.env.VITE_PI_ENTERPRISE_TASKPANE_ORIGIN ?? "https://127.0.0.1:3141";
const RUNNER_ORIGIN = import.meta.env.VITE_PI_ENTERPRISE_RUNNER_ORIGIN ?? "https://localhost:3141";

interface ConnectMessage { type: "enterprise-python-connect"; id: string; nonce: string }
interface PortMessage { type: "run" | "excel-result"; id: string; nonce: string; payload: unknown }
interface WorkerMessage { type: string; id?: string }

let connected = false;

window.addEventListener("message", (event: MessageEvent<ConnectMessage>) => {
  if (connected || window.location.origin !== RUNNER_ORIGIN
      || event.origin !== TASKPANE_ORIGIN || event.source !== window.parent) return;
  const message = event.data;
  if (message?.type !== "enterprise-python-connect"
      || typeof message.id !== "string" || !message.id
      || typeof message.nonce !== "string" || !message.nonce
      || event.ports.length !== 1) return;

  connected = true;
  const port = event.ports[0];
  if (!port) return;
  let worker: Worker | null = null;
  let started = false;

  const close = (): void => {
    worker?.terminate();
    worker = null;
    port.close();
  };
  port.addEventListener("message", (reply: MessageEvent<PortMessage>) => {
    const inbound = reply.data;
    if (inbound?.id !== message.id || inbound.nonce !== message.nonce) return;
    if (inbound.type === "run" && !started) {
      const request = inbound.payload;
      if (!request || typeof request !== "object" || !("type" in request)
          || request.type !== "run" || !("id" in request) || request.id !== message.id
          || !("code" in request) || typeof request.code !== "string") return;
      started = true;
      worker = new Worker(new URL("./enterprise-python-worker.ts", import.meta.url), { type: "module" });
      worker.addEventListener("message", (workerEvent: MessageEvent<WorkerMessage>) => {
        const response = workerEvent.data;
        if (response?.id !== message.id
            || !["stage", "excel-call", "done"].includes(response.type)) return;
        port.postMessage({ type: "forward", id: message.id, nonce: message.nonce, payload: response });
        if (response.type === "done") close();
      });
      worker.addEventListener("error", (error: ErrorEvent) => {
        port.postMessage({ type: "worker-error", id: message.id, nonce: message.nonce,
          error: error.message || "Python worker failed" });
        close();
      });
      worker.addEventListener("messageerror", () => {
        port.postMessage({ type: "worker-error", id: message.id, nonce: message.nonce,
          error: "Python worker message could not be decoded" });
        close();
      });
      worker.postMessage(request);
    } else if (inbound.type === "excel-result" && started && worker) {
      const response = inbound.payload;
      if (!response || typeof response !== "object" || !("type" in response)
          || response.type !== "excel-result") return;
      worker.postMessage(response);
    }
  });
  port.start();
  port.postMessage({ type: "ready", id: message.id, nonce: message.nonce });
});
