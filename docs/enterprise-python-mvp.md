# Enterprise Python + Excel: quick functional test (MVP)

**Status:** experimental feature-branch build. Use disposable, synthetic data
only. Comprehensive confinement review and new quotas are intentionally deferred.

## Windows setup

Checkout `feature/enterprise-python-functional-mvp` from
`julmx/pi-for-excel` and use the existing local launcher:

```powershell
git fetch origin
git switch --track origin/feature/enterprise-python-functional-mvp
.\scripts\start-enterprise-local.ps1 -ApiBaseUrl "https://your-internal-api.example/v1"
```

If already on this branch, use `git pull --ff-only` instead.
The first enterprise build downloads a pinned Pyodide runtime plus
NumPy/pandas wheels **during the build** into the gitignored
`public/enterprise-pyodide` folder. The running Excel add-in serves those
assets from `https://127.0.0.1:3141`, not a third-party CDN.
Normal upstream builds do not need Pyodide.

Keep the PowerShell server running, reopen Excel (or refresh the taskpane),
and **start a new Pi conversation** so the agent gets the updated tool list.
The first Python invocation loads Pyodide and can take longer.

## First useful test

Create a disposable workbook with a sheet named `Source`, values in A1:C3,
then ask the agent to run Python to create a `PythonResult` sheet and copy
the values with uppercase text. The agent should use the new `python_excel`
tool. A direct tool script looks like this:

```python
rows = await excel.read("Source!A1:C3")
cleaned = [[str(x).upper() for x in row] for row in rows]
await excel.create_sheet("PythonResult")
write = await excel.write("PythonResult!A1", cleaned)
result = {"rows": len(cleaned), "write": write["content"]}
```

Use only synthetic values. Other supported operations:
`await excel.format("PythonResult!A1:C1", bold=True)`,
`await excel.chart(action="create", source_range="PythonResult!A1:B3", chart_type="column")`,
`await excel.formula(range="PythonResult!D2:D5", formula="=A2", allow_overwrite=False)`,
or `await excel.call("modify_structure", action="rename_sheet", sheet="PythonResult", new_name="Output")`.

The `excel.call(name, **params)` entrypoint accepts the **approved existing
first-party Excel tool names only**. Native Excel Table objects are not yet
exposed by a dedicated approved tool, although ordinary range data and
formatting are supported.

## MVP tradeoffs

- All Excel RPCs run through existing first-party tool implementations.
- One Auto/Confirm decision applies to the **entire Python call**, including
  a read-only Python script. Auto requires no per-operation approval.
- This MVP uses a fresh Worker per call. Large repeated tasks may be slower.
- No Python-native filesystem bridge, arbitrary browser JS/Office.js tool,
  MCP, external Python package installation or other tool integrations.
- The previous restricted CSP is amended to allow only a local worker and
  WebAssembly execution. Runtime assets remain same-origin.
- The original automatic backup / overwrite checks are reused where supported.
  Do not assume multi-step scripts are atomic transactions.
- A comprehensive security audit and new memory/cell limits are deferred,
  but testing must not use confidential workbooks until they are completed.

If the build fails while preparing Pyodide, inspect the network/proxy setup
of the **build machine**. This download is separate from the company LLM API
and occurs before the Excel taskpane starts.
