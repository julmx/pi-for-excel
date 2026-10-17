import assert from "node:assert/strict";
import { test } from "node:test";

import { buildSystemPrompt } from "../src/prompt/system-prompt.ts";
import { filterEnterpriseBuiltInTools } from "../src/tools/enterprise-policy.ts";
import { CORE_TOOL_CAPABILITIES } from "../src/tools/capabilities.ts";

void test("enterprise agent prompt matches the approved Excel tool registry", () => {
  const prompt = buildSystemPrompt({
    enterpriseMode: true,
    executionMode: "yolo",
  });

  const names = filterEnterpriseBuiltInTools(CORE_TOOL_CAPABILITIES).map((tool) => tool.name);
  assert.ok(names.length > 0);
  for (const name of names) {
    assert.ok(prompt.includes(`**${name}**`), `missing tool ${name}`);
  }

  assert.match(prompt, /only the following Excel tools are available/i);
  assert.match(prompt, /python_excel.*Pyodide|Pyodide[\s\S]*python_excel/iu);
  assert.match(prompt, /await excel.read/iu);
  assert.match(prompt, /read_range.*csv/i);
  assert.match(prompt, /fill_formula/u);
  assert.match(prompt, /merged\s+group headings/u);
  assert.doesNotMatch(prompt, /workbook_history.*python_transform_range/u);
  assert.match(prompt.replace(/\s+/gu, " "), /use python_excel instead/iu);
  assert.doesNotMatch(prompt, /Two Python tools are always available/iu);
  assert.doesNotMatch(prompt, /run a Python snippet and inspect/iu);
  assert.doesNotMatch(prompt, /in-browser via Pyodide.*no setup required/iu);
  assert.doesNotMatch(prompt, /You have a persistent file workspace/u);
  assert.doesNotMatch(prompt, /use \*\*extensions_manager\*\* so it is installed/u);
});

void test("enterprise instructions supersede stale Python guidance without losing preferences", () => {
  const prompt = buildSystemPrompt({
    enterpriseMode: true,
    executionMode: "safe",
    userInstructions: "Always use Python for this workbook.",
  });
  assert.match(prompt, /Always use Python for this workbook/u);
  assert.match(prompt, /restricted tool\s+availability below takes precedence/iu);
  assert.match(prompt, /Current mode: \*\*Confirm\*\*/u);
  assert.match(prompt, /Ask for explicit user confirmation before mutating workbook tools/iu);
});

void test("ordinary builds retain their existing richer tool documentation", () => {
  const prompt = buildSystemPrompt({ enterpriseMode: false });
  assert.match(prompt, /Two Python tools are always available/u);
  assert.match(prompt, /You have a persistent file workspace/u);
  assert.match(prompt, /\*\*extensions_manager\*\*/u);
});
