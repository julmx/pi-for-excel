import assert from "node:assert/strict";
import { test } from "node:test";

import { filterEnterpriseAgentTools, filterEnterpriseBuiltInTools } from "../src/tools/enterprise-policy.ts";

void test("enterprise tools include the supported Excel reading and editing surface", () => {
  const tools = [
    { name: "read_range", id: 1 },
    { name: "write_cells", id: 2 },
    { name: "modify_structure", id: 3 },
    { name: "charts", id: 4 },
  ];

  const result = filterEnterpriseBuiltInTools(tools);
  assert.deepEqual(result, tools);
  assert.equal(result[0], tools[0]);
});

void test("enterprise tools exclude code execution, bridges, skills and arbitrary integrations", () => {
  const tools = [
    { name: "read_range" },
    { name: "execute_office_js" },
    { name: "python_run" },
    { name: "python_transform_range" },
    { name: "tmux" },
    { name: "mcp" },
    { name: "web_search" },
    { name: "fetch_page" },
    { name: "files" },
    { name: "skills" },
    { name: "instructions" },
    { name: "extensions_manager" },
    { name: "integration_custom" },
    { name: "write_cells" },
  ];

  assert.deepEqual(filterEnterpriseBuiltInTools(tools).map((tool) => tool.name), [
    "read_range",
    "write_cells",
  ]);
});

void test("only the agent receives the Office.js fallback, never Python RPC", () => {
  const tools = [
    { name: "read_range" },
    { name: "execute_office_js" },
    { name: "python_excel" },
    { name: "python_run" },
    { name: "files" },
  ];
  assert.deepEqual(filterEnterpriseAgentTools(tools).map((tool) => tool.name), [
    "read_range", "execute_office_js", "python_excel",
  ]);
  assert.deepEqual(filterEnterpriseBuiltInTools(tools).map((tool) => tool.name), [
    "read_range", "python_excel",
  ]);
});

void test("enterprise tools fail closed for names not explicitly permitted", () => {
  assert.deepEqual(filterEnterpriseBuiltInTools([{ name: "unknown_future_tool" }]), []);
});
