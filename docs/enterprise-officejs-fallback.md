# Enterprise Office.js fallback — limited restoration (prototype)

## Decision and security boundary

This change restores the **upstream** `execute_office_js` agent tool instead of recreating Excel features in bespoke tools. The intended user risk model is to prevent access to arbitrary files and data **outside the active workbook** on the Windows machine, while permitting broad Office.js operations **inside** the active Excel workbook.

Office.js does not grant general native Windows filesystem access. However, the **tool executes arbitrary model-generated JavaScript in the taskpane's privileged browser origin**, not in an isolated Office.js-only realm. That code can access browser APIs and **same-origin IndexedDB**, where Pi stores provider keys, sessions, and settings, including possibly older workbook content. This matters even though Windows files remain inaccessible directly. **No claim is made that the code can access only the active workbook or that the approach is safe for confidential workbooks.**

The user approval dialog is a risk-reduction measure, not a code sandbox. The upstream source-code risk regex is heuristic, not a security boundary.

## Minimal changes

- Reuse the existing `execute_office_js` implementation and approval dialog (no new general-purpose Excel command engine).
- Register it in the Enterprise *agent* tool list only; `python_excel` and its first-party structured RPC whitelist remain unchanged. `await excel.call("execute_office_js", ...)` must fail.
- Require **explicit confirmation before every `execute_office_js` call**, including Auto mode. The existing confirmation must fail closed if no UI is available. Structured tools remain governed by Auto/Confirm as before.
- Allow `blob:` for **taskpane script-src** so the existing dynamic module import can run. Keep the current `connect-src` restriction to the approved model gateway and Microsoft Office host, and **do not** enable `unsafe-eval`, open proxy access, native Python bridges, MCP, or unrestricted file access.
- Instruct the agent to prefer structured Excel tools and use Office.js only when needed for native Excel features such as filters, sorting, advanced conditional formatting.

## Limitations

1. JavaScript still runs on the **same origin** as taskpane secrets/session storage. There is no safe automatic guarantee that it will restrict its read access to the active workbook. This is an **open security concern** and must be addressed separately before claiming meaningful isolation from non-workbook application data.
2. The allowed gateway is an outbound channel. The CSP blocks arbitrary domains, but does not prohibit sending unexpected workbook data to an already-approved destination. The taskpane's network policy must be reviewed in actual Windows Excel WebView2.
3. Office.js operations can alter the workbook outside the stated range. The generic tool has no equivalent of the structured tools' per-cell overwrite checks and **no guaranteed automatic rollback**.
4. Some Excel operations (links, refresh, queries or other host-managed actions) may generate traffic from Excel itself, outside taskpane CSP scope.
5. The original `execute_office_js` imports a module using `blob:`. The revised CSP allows it in principle; **host compatibility has not been validated in real Excel Windows**.
6. The earlier dedicated Python isolation and its remaining runtime/payload validation limits are unchanged by this branch.

## Minimal Windows smoke test (synthetic only)

Use a disposable workbook and the existing local Enterprise Windows build instructions; switch to the branch `feature/enterprise-officejs-fallback-20261020` **without discarding local changes**. Launch the Enterprise local server with the approved enterprise model gateway as before.

1. Create a test sheet with values `A1:C4`: headers `Item, City, Score`; rows `a, X, 10`, `b, Y, 20`, `c, X, 30`.
2. Ask the assistant to apply a **native AutoFilter** to the range and verify that the filter arrows appear in Excel.
3. Ask for a **three-color conditional-formatting scale** on `C2:C4`, then visually verify the result.
4. In **Auto** mode, verify that Office.js still requests confirmation for **each** call; decline once and confirm no change, then approve a harmless operation. In **Confirm**, expect one dedicated Office.js request rather than a second generic mutation confirmation.
5. Ask Python to invoke `await excel.call("execute_office_js", code="return 1", explanation="Test")` and verify the bridge rejects the tool name.
6. With synthetic data, inspect DevTools Network/CSP violations in Excel WebView2 and verify no unexpected external destinations. Confirm `blob:` loading works without relaxing `connect-src`.
7. Do not use confidential data until this behavior, IndexedDB exposure and the corporate deployment policy have been reviewed.

## Acceptance

A passing test only means that the **upstream functional fallback** was restored without obvious regressions to Python or outbound CSP restrictions. It is **not** a certification, proof of no out-of-workbook access, or approval to process confidential data.

No changes are made to upstream core Excel tool implementations.
