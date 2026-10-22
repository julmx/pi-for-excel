# Windows smoke test: self-hosted enterprise add-in

> **Synthetic data only.** This setup is for a local functional and network
> smoke test, not a production deployment or approval for confidential data.

This procedure creates a locally hosted, restricted **production** build in
Excel for Windows. It does **not** install or call Pi for Excel's upstream
Vercel site. A separate Office add-in ID prevents replacement of an existing
stock Pi for Excel installation.

## Prerequisites

- Excel for Windows with Office taskpane add-in sideloading permitted by your
  Microsoft 365 tenant / organization.
- Windows PowerShell (5.1 or later) or PowerShell 7.
- Node.js 22.19+ (or Node.js 24+) and npm.
- [mkcert](https://github.com/FiloSottile/mkcert) for a local HTTPS
  certificate. Trusting its development CA modifies Windows certificate
  trust settings. **Obtain IT approval before doing this on a managed PC.**
- An approved HTTPS OpenAI-compatible company API that supports CORS from
  `https://127.0.0.1:3141` and, for workbook-editing agents, tool calling.
- No initial model ID is necessary. The add-in discovers the available models
  from the approved `<baseUrl>/models` endpoint.

Do **not** write an API key to a build variable, manifest, shell script,
public repository, or issue. Use the custom gateway's runtime credential
field after starting the add-in. The endpoint URL itself is bundled into
the local built JavaScript, so consider it deployment configuration.

## 1. Clone or update `main`

Open **PowerShell** on the Windows computer that runs Excel:

```powershell
git clone https://github.com/julmx/pi-for-excel.git
cd pi-for-excel
git fetch origin
git switch main
git pull --ff-only
```

If you already cloned the fork, run `git fetch origin`, `git switch main`,
and `git pull --ff-only`. Use the generated enterprise-local manifest, not
the upstream production manifest.

## 2. Create local HTTPS certificates (once)

From the repository root, after approval to trust a development CA:

```powershell
mkcert -install
mkcert -cert-file cert.pem -key-file key.pem localhost 127.0.0.1
```

Keep `cert.pem` and especially `key.pem` **on the local machine**. Both
are gitignored. The certificate must be trusted by the Windows profile
running Excel. If IT policy forbids local development CAs, do not bypass
that policy; obtain an IT-issued certificate or an approved testing method.

## 3. Build, generate the manifest and start the server

Pass your approved **base** URL, not the `/chat/completions` path. The
example below is deliberately fictional:

```powershell
.\scripts\start-enterprise-local.ps1 -ApiBaseUrl "https://api.example.invalid/llm/v1"
```

To authorize **multiple** corporate gateways, supply each approved API base URL
when building (not just in the taskpane settings):

```powershell
.\scripts\start-enterprise-local.ps1 -ApiBaseUrls @(
  "https://quality.example.invalid/v1",
  "https://genx.example.invalid/api"
)
```

Then configure both gateways in Settings, leaving Model ID empty for discovery,
and open the **model selector**: it already lists models from every registered
provider with their gateway names. A gateway stored in Settings but absent
from `-ApiBaseUrls` remains visible as **Not approved** and its models are
intentionally excluded. For a gateway change, stop the server, rebuild with
the updated approved URL list, and restart Excel; editing Settings alone
must never expand the network permissions.

The existing singular `-ApiBaseUrl` option remains supported. If one of the
approved APIs requires the optional same-origin relay for WebView2 CORS,
`-UseLocalApiRelay` remains available and forwards each indexed approved API
to its own pinned upstream. Use only approved test APIs and synthetic data.

The script runs `npm ci`, compiles with
`VITE_PI_ENTERPRISE_MODE=true`, creates the independent
`manifest.enterprise.local.xml`, and serves the production files over
HTTPS at `https://127.0.0.1:3141/src/taskpane.html`.

For subsequent runs, `-SkipInstall` skips `npm ci` if
`node_modules` already exists. Keep the PowerShell session open while
using Excel. Use **Ctrl+C** to stop the server.

Notes:
- `npm ci` retrieves third-party dependencies during the **build**.
  Review their security and license status before wider deployment.
- The server is bound to IPv4 loopback only. It is not a central enterprise
  deployment server and is intentionally unavailable from other computers.
- `manifest.enterprise.local.xml` is generated and gitignored. It points
  only to the local server for executable and image assets.
- The client still loads `office.js` from Microsoft's Office CDN.
- The local server sets a strict CSP and blocks connections to other origins,
  except the explicitly approved API origins and the Office.js CDN where applicable.
  CSP compatibility must be validated in the installed Excel WebView.
- No separately installed Python, tmux, MCP or external proxy server is needed.
  Enterprise Python runs through bundled Pyodide in an isolated-origin Worker.
  The same Node process serves both `https://127.0.0.1:3141` (Excel taskpane)
  and `https://localhost:3142` (minimal Python runner). They are distinct
  browser origins with separate IndexedDB storage. Taskpane cookies for
  `127.0.0.1` are not shared with `localhost`, but older applications hosted
  under `localhost` may have left cookies shared across ports. This prototype
  remains restricted to synthetic testing until that exposure is reviewed.
  Port 3142 must be available. The runner cannot use the API relay, and the
  taskpane may frame only that runner origin. The existing certificate covering
  `localhost` and `127.0.0.1` can be reused without installing mkcert.

## 4. Sideload into Excel for Windows

Use the generated **`manifest.enterprise.local.xml`** file, not
`manifest.prod.xml`.

Depending on your Excel build, use **Home / Insert → Add-ins → More
Add-ins / My Add-ins → Upload My Add-in** and select the generated XML.
Look for **Pi for Excel (Enterprise Local)** or **Open Pi Enterprise**
on the ribbon.

The specific menus vary with Office versions. Microsoft's [Office add-in
testing instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
cover sideloading, and its [Windows shared-folder catalog instructions](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)
describe an alternative testing path. A trusted network-share catalog may
require a tenant/IT policy change; it is **not a production deployment method**.

If your organization disables sideloading or the upload option is missing,
ask the Office administrator to authorize a test. Do not try to bypass tenant
controls or import the XML as Excel's legacy XML Expansion Pack.

## 5. Configure the corporate API and test

1. Create a **new disposable workbook**, with synthetic names and numbers.
2. In the add-in, open `/settings` and configure **Custom
   OpenAI-compatible gateway** with exactly the same approved base URL.
   Leave **Model ID** blank to discover every listed model. Configure **Max
   context tokens** and **Max output tokens** to appropriate gateway limits
   (both are persisted), and supply the runtime credential if required.
3. Use the model selector. The restricted build may discover additional
   models through the approved API's `/models` endpoint. The OpenAI-compatible
   discovery format usually returns just model IDs, not whether the server
   accepts adjustable reasoning. By default these models show **Thinking: Off**
   as their only option. For a model **verified** to support `reasoning_effort`
   at the configured gateway, edit that gateway and add its exact ID to
   **Model IDs with adjustable thinking** (comma-separated for several).
   Refresh/select that model and check that the Thinking menu now offers
   additional levels. This opt-in changes metadata, not API capabilities:
   a gateway that ignores or rejects the effort parameter will still not
   provide functional reasoning controls. Do not enable this for unknown models.
4. Ask: `Which sheet is active and what is in A1:B3?`
5. Switch the **Auto / Confirm** mode from the status bar. Check both modes
   with a trivial disposable-cell edit. User confirmation remains optional.
6. Check workbook history/recovery, then repeat a request after disabling
   the API. The add-in must not fall back to a public model.
7. In `python_excel`, run the following safe origin check:

   ```python
   from js import location
   result = {"python_origin": str(location.origin)}
   ```

   The reported value must be `https://localhost:3142`, **not** the
   taskpane origin `https://127.0.0.1:3141`. Then repeat the synthetic
   pandas read → new-sheet write test. A missing iframe or Worker error is
   a failed test; do not enable a same-origin fallback.

The model names may be visible without a working API if a saved catalogue
exists; that alone does not prove the endpoint or CORS integration works.

## 6. Network and security validation

Before using real data, verify on the **actual Windows Excel WebView**:

- The app's HTML and assets come from `127.0.0.1:3141`, not Vercel.
- LLM requests and `/models` discovery go only to the company-approved API.
- No prompts or cell contents are sent to public providers, web search,
  external bridges, or arbitrary proxies.
- Only expected Office.js / Microsoft 365 infrastructure connections occur.
- The API accepts CORS from `https://127.0.0.1:3141`. Missing CORS
  support can block direct model requests even if the endpoint is reachable.
- The CSP is enforced, not merely specified. Report any CSP errors.
- Auto and Confirm modes both work as intended.
- Stale provider credentials, stored sessions and error logs are reviewed.

If the strict CSP breaks the host integration, collect the rejected
origin/path from the WebView's developer tools and review it before permitting
it. Do **not** disable CSP or add wildcard `connect-src` rules to make a
confidential-data test appear to work.

### Windows Excel WebView2 acceptance — 18 October 2026

User-observed tests in the installed Excel add-in, using **synthetic data**:

- `python_excel` reported `location.origin = https://localhost:3142`;
  the Excel taskpane remained on `https://127.0.0.1:3141`.
- pandas created a DataFrame, wrote to a new worksheet and read it back;
  an attempted overwrite of a filled cell raised `RuntimeError` without
  changing the original values.
- An IndexedDB database named `__pi_isolation_manual_test` was created
  in the **taskpane** DevTools context. The parent verified
  `parent_has_canary: true`; a Python probe executed by the production
  isolated runner reported `canary_visible: false`.
  Only a synthetic marker was involved, never real application credentials.
  The test script requested deletion of the synthetic database in `finally`.
- A WebView2 Network HAR captured taskpane loopback API calls and local
  runner/Pyodide asset loads, with no other destinations **observed in that
  capture**. Separately, a 60-second Windows TCP sampling of the local
  Node.js server reported `51.159.8.229:443`, matching the contemporaneous
  DNS resolution of the configured enterprise gateway. Sampling and HAR
  capture cannot prove the absence of every possible outbound request.
- Linux and Windows Chromium tests additionally check the isolated storage
  boundary and that a synthetic disallowed network request does not reach
  its local test sink (PR #21).

**Remaining security and release limits:** These results confirm the tested
WebView2 isolation and network behavior, not a comprehensive sandbox audit
or an authorization to use confidential workbooks. Review existing
`localhost` cookies, browser-adjacent transports, provider-key/session
retention, relay authorization and organizational approval. The resource
limits, cancellation and partial-write scenarios were consciously deferred
during this validation; no claim is made that they passed.
