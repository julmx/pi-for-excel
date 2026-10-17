import assert from "node:assert/strict";
import { test } from "node:test";
import { Type } from "typebox";
import type { AnyCoreTool } from "../src/tools/registry.ts";
import { createEnterpriseExcelBridge } from "../src/python/enterprise-excel-bridge.ts";

function fakeTool(name: string, events: string[]): AnyCoreTool {
  return {
    name,
    label: name,
    description: name,
    parameters: Type.Object({}),
    execute: (_id, params) => {
      events.push("execute:" + name);
      return Promise.resolve({
        content: [{ type: "text", text: JSON.stringify(params) }],
        details: { kind: name },
      });
    },
  };
}

void test("reads run without write approvals; Excel mutations pass through the gate", async () => {
  const events: string[] = [];
  const bridge = createEnterpriseExcelBridge({
    tools: [fakeTool("read_range", events), fakeTool("write_cells", events)],
    newToolCallId: () => "tool-1",
    approveMutation: (name) => {
      events.push("approve:" + name);
      return Promise.resolve();
    },
  });

  const read = await bridge.handle({ id: "a", tool: "read_range", params: { range: "A1:B2" } });
  assert.equal(read.ok, true);
  assert.deepEqual(events, ["execute:read_range"]);

  const write = await bridge.handle({ id: "b", tool: "write_cells", params: { start_cell: "A1", values: [[1]] } });
  assert.equal(write.ok, true);
  assert.deepEqual(events, ["execute:read_range", "approve:write_cells", "execute:write_cells"]);
});

void test("all permitted structured Excel tools use their existing implementation", async () => {
  const names = [
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
  ];
  const events: string[] = [];
  const bridge = createEnterpriseExcelBridge({
    tools: names.map((name) => fakeTool(name, events)),
    newToolCallId: () => "call",
    approveMutation: () => Promise.resolve(),
  });
  for (const tool of names) {
    const result = await bridge.handle({ id: tool, tool, params: { action: "create" } });
    assert.equal(result.ok, true, tool);
  }
  assert.equal(events.length, names.length);
});

void test("unknown tools and invalid RPC payloads fail closed", async () => {
  const events: string[] = [];
  const bridge = createEnterpriseExcelBridge({
    tools: [
      fakeTool("write_cells", events),
      fakeTool("execute_office_js", events),
      fakeTool("python_run", events),
      fakeTool("tmux", events),
      fakeTool("mcp", events),
      fakeTool("extensions_manager", events),
    ],
    newToolCallId: () => "call",
    approveMutation: () => Promise.resolve(),
  });
  for (const tool of ["execute_office_js", "python_run", "tmux", "mcp", "extensions_manager", "constructor"]) {
    const result = await bridge.handle({ id: "request", tool, params: {} });
    assert.deepEqual(result, { id: "request", ok: false, error: "Excel operation not permitted." });
  }
  for (const payload of [null, [], { id: 123, tool: "write_cells", params: {} },
    { id: "a", tool: "write_cells", params: [] }]) {
    const result = await bridge.handle(payload);
    assert.equal(result.ok, false);
  }
  assert.deepEqual(events, []);
});

void test("confirmation denial and cancellation prevent execution", async () => {
  const events: string[] = [];
  const bridge = createEnterpriseExcelBridge({
    tools: [fakeTool("write_cells", events)],
    newToolCallId: () => "call",
    approveMutation: () => Promise.reject(new Error("Not approved")),
  });
  const denied = await bridge.handle({ id: "1", tool: "write_cells", params: {} });
  assert.deepEqual(denied, { id: "1", ok: false, error: "Not approved" });
  const cancelled = await bridge.handle(
    { id: "2", tool: "write_cells", params: {} }, AbortSignal.abort(),
  );
  assert.deepEqual(cancelled, { id: "2", ok: false, error: "Aborted" });
  assert.deepEqual(events, []);
});

function responseTool(name: string, text: string, details: Record<string, unknown>): AnyCoreTool {
  return {
    name,
    label: name,
    description: name,
    parameters: Type.Object({}),
    execute: () => Promise.resolve({
      content: [{ type: "text", text }],
      details,
    }),
  };
}

void test("blocked writes are reported as failures to Python", async () => {
  const bridge = createEnterpriseExcelBridge({
    tools: [responseTool("write_cells", "Write blocked — destination already contains values.", {
      kind: "write_cells",
      blocked: true,
    })],
    newToolCallId: () => "write",
    approveMutation: () => Promise.resolve(),
  });

  const result = await bridge.handle({
    id: "1",
    tool: "write_cells",
    params: { start_cell: "Sheet1!A1", values: [[42]] },
  });
  assert.deepEqual(result, {
    id: "1",
    ok: false,
    error: "Write blocked — destination already contains values.",
  });
});

void test("Excel tools that return error text do not silently succeed", async () => {
  const tools = [
    responseTool("read_range", "Error reading \"Missing!A1\": worksheet not found.", {}),
    responseTool("write_cells", "Error writing cells: worksheet protected.", {
      kind: "write_cells",
      blocked: false,
    }),
    responseTool("modify_structure", "Error: A worksheet with this name already exists.", {
      kind: "modify_structure",
      action: "add_sheet",
    }),
  ];
  const bridge = createEnterpriseExcelBridge({
    tools,
    newToolCallId: () => "call",
    approveMutation: () => Promise.resolve(),
  });

  for (const tool of tools) {
    const result = await bridge.handle({ id: tool.name, tool: tool.name, params: {} });
    assert.equal(result.ok, false, tool.name);
    if (!result.ok) assert.match(result.error, /^Error/u);
  }
});

void test("successful Excel tool results still pass through unchanged", async () => {
  const bridge = createEnterpriseExcelBridge({
    tools: [responseTool("read_range", "**Sheet1!A1**: errors found in source data.", {
      kind: "read_range_csv",
      values: [["Error count", 2]],
    })],
    newToolCallId: () => "read",
    approveMutation: () => Promise.resolve(),
  });
  const result = await bridge.handle({ id: "1", tool: "read_range", params: { range: "A1" } });
  assert.equal(result.ok, true);
  if (result.ok) assert.deepEqual(result.result.details, {
    kind: "read_range_csv",
    values: [["Error count", 2]],
  });
});
