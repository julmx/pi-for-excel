# Enterprise Python + Excel: quick functional test (MVP)

**Status:** experimental enterprise build on `main`. Use disposable, synthetic data
only. Comprehensive confinement review and new quotas are intentionally deferred.

## Windows setup

Use the current `main` branch of `julmx/pi-for-excel` (including PR #12 and
PR #13) with the existing local launcher:

```powershell
git fetch origin
git switch main
git pull --ff-only
.\scripts\start-enterprise-local.ps1 -ApiBaseUrl "https://your-internal-api.example/v1"
```

If already on `main`, just run `git pull --ff-only` after fetching.
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

To verify that PR #13 does not report blocked writes as successful, run this
second script while `Source!A1` still contains data:

```python
await excel.write("Source!A1", [["SHOULD_NOT_OVERWRITE"]])
```

Python should report an error indicating that the write was blocked, and
`Source!A1` must remain unchanged. Do not pass `allow_overwrite=True` for this
test. Run it only on a disposable workbook.

Use only synthetic values. Other supported operations:
`await excel.format("PythonResult!A1:C1", bold=True)`,
`await excel.chart(action="create", source_range="PythonResult!A1:B3", chart_type="column")`,
`await excel.formula(range="PythonResult!D2:D5", formula="=A2", allow_overwrite=False)`,
or `await excel.call("modify_structure", action="rename_sheet", sheet="PythonResult", new_name="Output")`.

The `excel.call(name, **params)` entrypoint accepts the **approved existing
first-party Excel tool names only**. Native Excel Table objects are not yet
exposed by a dedicated approved tool, although ordinary range data and
formatting are supported.

## Troubleshooting: enterprise model blocked / proxy warning

If the taskpane says `Enterprise policy blocked an unapproved API endpoint or
provider`, it has not yet executed Python. This comes from the **selected
LLM model**, not the Python Worker or port 3003.

1. Save the workbook, stop the old PowerShell server, and close all Excel
   windows before restarting. A stale Excel WebView can keep an earlier build.
2. Confirm your checkout with `git branch --show-current`. It must be
   `main`, and `git pull --ff-only` must have completed.
3. Run `start-enterprise-local.ps1` with the **same exact approved HTTPS
   base URL** already configured under **Settings → Custom OpenAI-compatible
   gateway**, including the URL path (typically `/v1`). Do not paste a key
   or token into `-ApiBaseUrl`.
4. Confirm the installed manifest is `manifest.enterprise.local.xml`,
   reload the taskpane and select the actual gateway model. Start a new
   chat to exclude previously saved provider selections and error messages.
5. If the yellow `Proxy not running` banner persists, the current taskpane
   is likely using a non-enterprise or cached build. It is intentionally
   absent from the enterprise application. The optional localhost:3003
   proxy does not fix the corporate API model policy.

The branch also prefers a registered approved gateway over its temporary
`enterprise-bootstrap` model and ignores invalid providers restored from
old sessions. It does not relax endpoint validation.

If random-looking text you pasted into a chat was an API key, revoke and
rotate the key; a chat message is not the secure credential input.

## If LLM requests show "Connection error" (Windows WebView2)

First verify the configured gateway matches the `-ApiBaseUrl` used to build.
A generic connection failure can be caused by CORS, DNS/TLS or gateway network
policy. The `GET /src/taskpane.html → 200` lines only verify local hosting;
they do **not** prove the enterprise API received any traffic.

If the gateway rejects cross-origin browser requests, enable the optional
**same-origin local API relay**. The HTTPS server on 127.0.0.1:3141 forwards
only `GET /models` and `POST /chat/completions` to the one configured
approved HTTPS endpoint, with normal TLS certificate validation. It does not
start the legacy port 3003 proxy or offer arbitrary URL forwarding:

```powershell
.\scripts\start-enterprise-local.ps1 -ApiBaseUrl "https://your-approved-api.example/v1" -SkipInstall -UseLocalApiRelay
```

Restart Excel completely and open a fresh chat. This **rebuilds the frontend**
with same-origin request mapping; an earlier build will not be changed by
toggling the server alone. The local Node.js process temporarily receives
request bodies and runtime API credentials in memory so it can forward them
to the configured gateway; it does not persist them or log them.

If the result is `Enterprise API unreachable from local HTTPS server`, check
the Node.js machine's outbound TLS reachability or its required enterprise
network proxy. If the gateway responds with 401/403, examine its authentication
or access policy. Do not disable certificate verification or paste API keys
into logs/chat. If the relay works, that is useful evidence of a browser
CORS issue, though not an exhaustive diagnosis.

To keep using direct browser requests, launch without
`-UseLocalApiRelay`.

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
