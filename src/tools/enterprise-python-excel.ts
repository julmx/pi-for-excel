/**
 * Enterprise-only Python tool.
 *
 * The agent sees one coordinated workbook operation. All inner calls dispatch
 * to the first-party Excel tools and never bypass their overwrite checks and
 * recovery hooks. The outer coordinator handles one Auto/Confirm decision
 * for the whole script (including read-only scripts in this first MVP).
 */
import type { AgentTool } from "@earendil-works/pi-agent-core";
import { Type, type Static } from "typebox";
import type { SpreadsheetHostKind } from "../host/index.js";
import { createCoreTools } from "./registry.js";
import { filterEnterpriseBuiltInTools } from "./enterprise-policy.js";
import { createEnterpriseExcelBridge } from "../python/enterprise-excel-bridge.js";
import { runEnterprisePython } from "../python/enterprise-python-runtime.js";

const schema = Type.Object({
  code: Type.String({
    minLength: 1,
    description: "Python code to run in isolated Pyodide. Use await excel.read(range), await excel.write(start_cell, values), await excel.create_sheet(name), await excel.format(range, **options), await excel.chart(**options), or await excel.call(tool, **params). Assign JSON-serializable output to result.",
  }),
});
type Params = Static<typeof schema>;

export function createEnterprisePythonExcelTool(hostKind: SpreadsheetHostKind): AgentTool<typeof schema> {
  return {
    name: "python_excel",
    label: "Python + Excel",
    description: "Execute Python/pandas/numpy and read/write Excel in one script using the structured excel API. No native Python bridge or arbitrary Office.js.",
    parameters: schema,
    execute: async (_toolCallId, params: Params, signal) => {
      const tools = filterEnterpriseBuiltInTools(createCoreTools({ hostKind }));
      // This tool is always classified 'mutate' by execution-policy.ts and
      // runs inside withWorkbookCoordinator's batch write authorization.
      // The worker cannot bypass that outer authorization.
      const bridge = createEnterpriseExcelBridge({
        tools,
        approveMutation: () => Promise.resolve(),
        newToolCallId: () => crypto.randomUUID(),
      });
      const response = await runEnterprisePython(params.code, bridge.handle, signal);
      if (!response.ok) {
        return {
          content: [{ type: "text", text: `Python execution failed: ${response.error ?? "Unknown error"}` }],
        };
      }
      return {
        content: [{
          type: "text",
          text: `Python completed. Result: ${response.resultJson ?? "null"}`,
        }],
      };
    },
  };
}
