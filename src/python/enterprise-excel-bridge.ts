/**
 * Explicit Excel RPC boundary for a future enterprise Pyodide worker.
 *
 * The worker must never receive Office.js objects or callable JS references.
 * It can submit only a structured operation naming an approved first-party
 * Excel tool. Callers MUST supply the workbook mutation gate, including the
 * user's Auto / Confirm preference; this bridge cannot bypass it.
 *
 * This module does not start Python or install a worker.
 */
import type { AnyCoreTool } from "../tools/registry.js";
import { filterEnterpriseBuiltInTools } from "../tools/enterprise-policy.js";
import { getToolExecutionMode } from "../tools/execution-policy.js";

export interface EnterpriseExcelOperation {
  id: string;
  tool: string;
  params: Record<string, unknown>;
}

export type EnterpriseExcelOperationResponse =
  | { id: string; ok: true; result: Awaited<ReturnType<AnyCoreTool["execute"]>> }
  | { id: string; ok: false; error: string };

export interface EnterpriseExcelBridgeOptions {
  tools: readonly AnyCoreTool[];
  /** Called exactly once before each workbook-mutating operation. */
  approveMutation: (tool: string, params: Record<string, unknown>) => Promise<void>;
  /** Return a stable operation identifier for existing workbook audit logs. */
  newToolCallId: () => string;
}

type RpcInput = object | string | number | boolean | null | undefined;

function isRecord(value: RpcInput): value is Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  const proto: unknown = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

function parseRequest(raw: RpcInput): EnterpriseExcelOperation {
  if (!isRecord(raw) || typeof raw.id !== "string" || raw.id.length === 0
    || typeof raw.tool !== "string" || raw.tool.length === 0
    || (typeof raw.params !== "object" && raw.params !== null)
    || !isRecord(raw.params)) {
    throw new Error("Invalid Excel bridge request.");
  }

  return { id: raw.id, tool: raw.tool, params: raw.params };
}

/**
 * A request handler for the host side of the Worker RPC.
 *
 * Existing Excel tools own their parameter validation, overwrite protection,
 * audit records, and recovery behavior. The forthcoming integration must
 * attach this handler to the already coordinated workbook runtime.
 */
export function createEnterpriseExcelBridge(options: EnterpriseExcelBridgeOptions): {
  handle: (raw: RpcInput, signal?: AbortSignal) => Promise<EnterpriseExcelOperationResponse>;
} {
  const permitted = filterEnterpriseBuiltInTools(options.tools);
  const byName = new Map(permitted.map((tool) => [tool.name, tool]));

  return {
    async handle(raw, signal) {
      let request: EnterpriseExcelOperation;
      try {
        request = parseRequest(raw);
      } catch (error) {
        return {
          id: isRecord(raw) && typeof raw.id === "string" ? raw.id : "",
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        };
      }

      const tool = byName.get(request.tool);
      if (!tool) {
        return { id: request.id, ok: false, error: "Excel operation not permitted." };
      }

      try {
        if (signal?.aborted) throw new Error("Aborted");
        if (getToolExecutionMode(tool.name, request.params) === "mutate") {
          await options.approveMutation(tool.name, request.params);
        }
        if (signal?.aborted) throw new Error("Aborted");
        const result = await tool.execute(options.newToolCallId(), request.params, signal);
        return { id: request.id, ok: true, result };
      } catch (error) {
        return {
          id: request.id,
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        };
      }
    },
  };
}
