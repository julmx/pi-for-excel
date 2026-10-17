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
