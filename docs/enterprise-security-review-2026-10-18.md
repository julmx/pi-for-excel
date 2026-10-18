# Enterprise security and upstream divergence review — 2026-10-18

## Scope and decision

Static source-code review of `julmx/pi-for-excel` at `09ff3d94` compared with `tmustier/pi-for-excel` `main` at `9a2aef73`. Reviewed the enterprise-specific Python Worker, Excel RPC, allowlists, model/network policy, local HTTPS relay, CSP, credential storage, dependency download, CI and upstream differences.

**Decision: working functional MVP, not approved for confidential enterprise workbooks.** Real Excel/Windows testing successfully exercised blocked-write exceptions, pandas transforms and a 500-row × 150-column pipeline. That evidence does not establish a security boundary.

This is a focused **static audit**, not a penetration test or formal certification. No secret extraction, adversarial worker execution, confidential workbook, production deployment, or packet capture was performed. Separate confirmed code properties from plausible impact and verification still required.

## Findings (prioritized)

### H1 — Model-generated Pyodide code shares the taskpane storage origin (high, architectural)

**Evidence:**
- `src/python/enterprise-python-runtime.ts` starts the Worker with `new Worker(new URL("./enterprise-python-worker.ts", import.meta.url))`, on the taskpane origin.
- `src/python/enterprise-python-worker.ts` executes agent-provided Python through unrestricted `runPythonAsync`; only Excel *RPC* operations are allowlisted.
- Pyodide intentionally exposes Worker JavaScript globals through the Python `js` module (Pyodide documentation: https://github.com/pyodide/pyodide/blob/main/docs/usage/type-conversions.md).
- `WorkerGlobalScope.indexedDB` is available in browser workers (MDN: https://developer.mozilla.org/en-US/docs/Web/API/WorkerGlobalScope/indexedDB).
- `src/storage/init-app-storage.ts` uses IndexedDB database `pi-for-excel`; `src/storage/local/provider-keys-store.ts` keeps API keys in the `provider-keys` store; sessions and settings also reside there.

**Impact:** A worker isolates execution from the main UI thread, not from same-origin storage. The Excel RPC allowlist and Worker `connect-src 'self'` do **not** make arbitrary Pyodide scripts a secure sandbox. A model-generated script may have access to same-origin browser capabilities beyond `excel.read/write`; the actual reachable APIs should be validated with synthetic data. An unrestricted external exfiltration channel is **not** established: CSP and the approved gateway still constrain egress.

**Follow-up — synthetic proof (2026-10-18):** The existing compiled enterprise Python Worker successfully read `SYNTHETIC_CANARY_NOT_A_SECRET` from `__pi_synthetic_worker_isolation_probe` through `from js import indexedDB` under the restricted CSP, in Chromium on Windows and Linux. The browser smoke deletes this test-only database. It never reads the real `pi-for-excel` storage and does not prove identical behavior in the Excel WebView2 host. See `scripts/smoke-enterprise-python-browser.mjs` (isolation characterization). This directly confirms same-origin IndexedDB access, not external exfiltration.

**Disposition:** Block authorization of confidential workbooks. The smallest credible containment candidate is a Python runner on a **different HTTPS origin**, communicating with the existing host Excel RPC through an origin-verified, per-run authenticated message channel. The runner must not receive or persist API credentials; it should load only local Python assets with its own restrictive CSP. Test a denied synthetic same-origin storage read from the runner, while retaining pandas/Excel functionality in Windows Excel. Avoid changing the Excel tool implementations, adding Python import blacklists, or claiming an ordinary same-origin Worker is a sandbox.

### H2 — Incomplete resource and payload limits (high for availability)

`src/python/enterprise-python-runtime.ts` has a 120-second Worker timeout, but `src/python/enterprise-excel-bridge.ts` contains no explicit RPC-count, cell-count or serialized-byte ceiling; the Worker has no application-enforced Python memory ceiling. A 75,000-cell successful smoke test proves useful capacity, not a safe maximum.

**Disposition:** Specify a limited number of Excel operations, input/output bytes, cells per range, and a bounded memory/concurrency budget. Add rejection tests without allocating genuinely excessive memory. Preserve existing timeout and cancellation semantics. Inspect Office operations still in flight when a Worker times out; multi-step edits are not transactional.

### M1 — Build-time supply-chain integrity is not content-pinned

`scripts/prepare-enterprise-pyodide.mjs` downloads the fixed Pyodide version 0.27.7 and its selected wheel assets over HTTPS from jsDelivr, but accepts already cached non-empty files and does not verify a local cryptographic digest or artifact manifest. Fixed version plus TLS does not establish reproducible bytes.

**Disposition:** Pin and verify integrity of downloaded and cached assets; continue self-hosting them at runtime. Reuse an existing verified distribution/lock workflow rather than adding a parallel package system.

### M2 — Python RPC argument trust is broader than the declared TypeBox schemas

`src/python/enterprise-excel-bridge.ts` verifies the RPC request envelope and allowed tool name, but directly delegates `request.params` to `tool.execute`. `vite.config.ts` explicitly substitutes the Ajv validator with a stub, with a skip-validation fallback in Office. Individual tools have some checks but are not a substitute for one validated RPC boundary.

**Disposition:** Validate arguments using the *existing tool TypeBox schema* with a CSP-compatible interpreter (if proven compatible), and reject oversized/deeply nested payloads. Do not introduce a second independent Excel operation schema.

### M3 — Optional local API relay needs hostile-client review

`scripts/enterprise-api-relay.mjs` pins upstream path/method and retains TLS verification; `scripts/serve-dist-https.mjs` binds to loopback. Good controls. However the relay handler does not explicitly check request `Origin`/Fetch Metadata, enforce a maximum request body size, or authenticate the local caller. The relay is explicitly a development workaround, not an enterprise production gateway.

**Disposition:** Before broader deployment, test cross-origin requests and enforce appropriate origin/size/timeout bounds. Prefer a company-managed, authenticated gateway for production.

### M4 — Credentials, sessions and workbook data persist in browser IndexedDB

Existing upstream storage persists provider keys and sessions. This becomes especially relevant when arbitrary agent-generated Python is run on the **same origin** (H1). Scope retention, local profile access, session deletion, gateway authentication and whether the browser needs a reusable API key at all.

**Disposition:** For sensitive data, evaluate an enterprise-authenticated gateway with no long-lived client API secret, and apply retention/clearing policy. This is not fixed by merely hiding the API key in the UI.

### M5 — Multi-step Python transformations are not atomic

The outer workbook coordinator serializes the `python_excel` tool and obtains one Auto/Confirm decision; inner operations reuse first-party checks/history, with `approveMutation: () => Promise.resolve()` because the outer gate has already authorized the script. Individual Excel writes can succeed before a later Python failure/timeout. No all-or-nothing rollback is guaranteed.

**Disposition:** Keep the behavior explicit in UI/docs, test partial failure/recovery, and do not promise one-click rollback for arbitrary scripts. Avoid a new transaction engine unless a concrete user requirement arises.

### M6 — Enterprise-specific CI is branch-scoped

`.github/workflows/enterprise-checks.yml` triggers on `feature/enterprise-*`, not every push/PR to `main`. The ordinary CI catches many tests, but the full self-hosted Pyodide/Chromium Linux+Windows enterprise smoke may be skipped on unrelated `main` PRs.

**Disposition:** Run the existing enterprise smoke for changes to its runtime/packaging/security paths on `main`, preferably through targeted workflow triggers, not another test framework.

## Controls that are already present and should be retained

- Model selection checks exact approved OpenAI-compatible gateway base URL/protocol; public-provider fallback is rejected in the restricted runtime (`src/auth/enterprise-model-policy.ts`).
- Browser fetch policy restricts path and destination, while enterprise CSP constrains scripts and outbound traffic; these are defense in depth, **not** a standalone egress proof.
- Allowed first-party Excel tools and typed RPC reuse existing write-protection, readback, recovery and workbook operation code.
- `python_excel` is treated as a mutating operation and receives the outer user-configurable Auto/Confirm gate.
- Self-hosted Pyodide/NumPy/pandas avoids runtime third-party CDN loads; the local HTTPS server is loopback-only for the smoke setup.
- Linux/Windows automated Worker integration tests and actual Windows Excel functional tests have passed, including blocked-cell error propagation. They do not cover worker storage isolation or adversarial network behavior.

## Upstream divergence snapshot

GitHub compare `tmustier/main...julmx/main`: **189 fork-side commits** and **5 upstream-side commits** not shared at the reviewed refs; 53 changed files returned. This does *not* mean the fork reimplemented the original project: the majority of upstream Excel tools, UI and original Python implementation remain intact. The fork reports `0.10.0-pre` with Pi AI/Agent Core `0.87.0`; upstream `main` reports `0.11.1-pre` with Pi AI/Agent Core `1.3.0`. These major runtime differences warrant compatibility tests before selective upstream updates.

| Difference | Justification / assessment |
| --- | --- |
| Restricted model/provider selection, network/CSP and local HTTPS smoke deployment | Necessary for a private approved gateway; preserve while validating egress. |
| Self-hosted enterprise Pyodide Worker and structured Excel RPC | Justified by no-runtime-CDN and multi-operation Excel script requirements. Shared Python execution code from upstream reduces duplication; security isolation remains unresolved (H1). |
| Enterprise-specific tool list and runtime shutdown of extensions/bridges | Appropriate capability reduction, but cannot replace a worker security boundary. |
| Original Python worker, existing Excel tool behavior, UI/agent orchestration | Mostly reused; keep parity wherever possible. |
| Upstream dependency/model/MCP/runtime changes not yet integrated | Review individually; no unconditional wholesale merge. Upstream has newer dependency security baselines, including a temporary **unmerged node-forge patch** described in `docs/security-dependencies.md`. Assess the advisory and its scope before selectively updating the fork; do not blindly copy an unreviewed archive patch. |

Upstream references: https://github.com/tmustier/pi-for-excel and https://github.com/tmustier/pi-for-excel/blob/main/docs/security-dependencies.md .

## Recommended next actions — minimum change

1. **Now:** keep enterprise build limited to synthetic/non-confidential workbooks. Record H1 as the release blocker and perform a synthetic-only capability probe of JS/IndexedDB access in the real Office WebView2.
2. **Before confidential use:** resolve H1 with a genuinely separate trust/storage boundary; confirm with an adversarial test. Do not add a fake Python import blacklist.
3. **Then:** add compact caps/validation at the existing Excel RPC boundary; test cancellation and partial commits; tighten optional relay only if retained.
4. **Maintenance:** verify Pyodide asset digests; selectively review upstream's newer dependency security updates and compare five outstanding commits; maintain the existing smoke on relevant merges.
5. **After these controls:** perform real outbound traffic capture, credential retention review and enterprise IT approval on the intended Windows deployment.

**Acceptance threshold:** functional success is established; **security approval is not**. No security change should weaken the current restrictive model allowlist, network policy, or Auto/Confirm behavior. Avoid unrelated refactors and new infrastructure unless the trust-boundary design specifically requires it.
