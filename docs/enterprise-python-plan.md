# Enterprise Python + Excel: implementation plan

This document records the agreed V1.1 scope for the enterprise fork.
It is a development plan, **not an approval to process confidential workbooks**.

## Scope

Let an enterprise Pyodide script combine computation and Excel operations in
one agent tool call. The script should be able to use every approved
first-party Excel operation available to the agent, including reading ranges,
writing values and formulas, managing sheets, formatting, conditional
formatting, comments, chart creation/updates, and workbook inspection.

Do not give generated Python a general Office.js object, JavaScript evaluator,
native Windows interpreter, shell, network client, arbitrary package installer,
or access to local OneDrive files.

**Structured Excel tables:** The current enterprise allowlist can work with
cell ranges, but it does not include a first-party tool for creating native
Excel Table objects. Add a dedicated, parameter-validated table operation
before advertising `excel.table(...)`. Do not fall back to arbitrary
Office.js or silently claim Table object support. PivotTables are also not
part of the initial structured tool surface.

## Architecture

1. The agent calls a new `python_excel` tool, passing Python code and context.
2. A **self-hosted** Pyodide Web Worker receives code, with local pinned
   dependencies. No native Python bridge or remote CDN fallback.
3. Python receives an `excel` module with a narrow RPC API, including
   `excel.read(...)`, `excel.write(...)`, `excel.create_sheet(...)`,
   `excel.format(...)`, `excel.chart(...)` and a generic
   `excel.call(tool_name, params)` escape hatch **only** for existing
   approved structured Excel tool names. There is no `exec_js` operation.
4. A TypeScript host dispatcher validates each RPC operation against the
   enterprise tool allowlist and routes it to the existing Excel tool
   implementation. It does not grant Python a JavaScript/Office.js handle.
5. Modifications follow the existing workbook coordinator, audit, backup,
   overwrite protection, and user-configurable Auto / Confirm behavior.
   Aim to aggregate confirmation for a Python transformation rather than
   prompting for every underlying Excel operation. Never silently overwrite
   occupied cells just because Auto mode was selected.
6. Results and errors return to the agent as structured data, with visible
   summaries of changes.

## Delivery sequence

### PR 1: Excel RPC contract and host dispatch

`src/python/enterprise-excel-bridge.ts` provides an explicit allowlisted
host operation dispatcher, invoked only with approved, first-party tool
instances. It accepts an **obligatory** mutation authorization callback,
and routes calls to existing tools instead of recreating Excel writes.
This foundational PR has no running Python tool and no effect on released
enterprise behavior.

**Integration caveat:** In later PRs the dispatcher must be attached inside
the same workbook-coordinator operation, not recursively inside a second
nested coordinator lock. Confirmation must be aggregated at the outer
`python_excel` operation, with a real host-provided mutation gate for any
inner operations. Do not pass an unconditional no-op gate to circumvent
Confirm mode.

### PR 2: local Pyodide worker and Python/Excel interaction

Bundle a pinned Pyodide distribution and the chosen packages locally. Add a
Worker-host message protocol and Python-side `excel` library. Support
read/write/formulas/sheet operations first, then formatting, charts and other
structured tools through the same RPC interface. Add a dedicated first-party
Table object tool if native Excel Tables are required. Adapt the enterprise CSP
for the Worker and its local WebAssembly resources while preserving
restrictions on external destinations.

### PR 3: integrate agent tool and Excel approval/recovery

Create `python_excel` for enterprise builds only and expose it in the
system prompt. Run Python inside one top-level tool invocation. Ensure
Auto/Confirm applies to the overall transformation with an understandable
write summary and make failures/recovery visible. Reuse existing structured
tool implementations and workbook history; test wide merged-header
transformations and representative multi-sheet tasks.

## Deferred by design

- A **comprehensive confinement and security audit**, including adversarial
  prompts and exhaustive network capture, will be a separate final project
  phase rather than PR 4 in the initial implementation sequence.
- **New configurable resource ceilings** (cell count, memory budgets, payload
  volume, execution concurrency) are not a V1.1 priority. Retain existing
  cancellation/timeout behavior and do not deliberately weaken existing
  protections while implementing the worker.
- Complete all standard CI checks and basic functional smoke tests on each PR.
  Deferring the comprehensive audit is **not** permission to ignore known
  security failures or existing company deployment controls.

The goal is useful Python/Excel functionality first, then a full security
qualification before confidential production deployment.
