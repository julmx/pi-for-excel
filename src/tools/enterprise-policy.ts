/**
 * Opt-in enterprise build: expose only reviewed first-party Excel tools to the agent.
 *
 * This is a tool-surface reduction, NOT a complete data-egress or security
 * boundary. See docs/enterprise-hardening.md before using confidential data.
 */
const ENTERPRISE_EXCEL_TOOL_NAMES = new Set<string>([
  "get_workbook_overview",
  "read_range",
  "write_cells",
  "fill_formula",
  "search_workbook",
  "modify_structure",
  "format_cells",
  "conditional_format",
  "charts",
  "trace_dependencies",
  "explain_formula",
  "view_settings",
  "comments",
  "workbook_history",
  "python_excel",
]);

/** Preserve tool instances and ordering for the *Python RPC* whitelist. */
export function filterEnterpriseBuiltInTools<T extends { name: string }>(
  tools: readonly T[],
): T[] {
  return tools.filter((tool) => ENTERPRISE_EXCEL_TOOL_NAMES.has(tool.name));
}

/** Only the agent can use the Office.js fallback, never isolated Python. */
export function filterEnterpriseAgentTools<T extends { name: string }>(tools: readonly T[]): T[] {
  return tools.filter((tool) => ENTERPRISE_EXCEL_TOOL_NAMES.has(tool.name)
    || tool.name === "execute_office_js");
}
