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
- An initial model ID exposed by the API (the add-in can discover other
  models later from `<baseUrl>/models`).

Do **not** write an API key to a build variable, manifest, shell script,
public repository, or issue. Use the custom gateway's runtime credential
field after starting the add-in. The endpoint URL itself is bundled into
the local built JavaScript, so consider it deployment configuration.

## 1. Clone the feature branch

Open **PowerShell** on the Windows computer that runs Excel:

```powershell
git clone https://github.com/julmx/pi-for-excel.git
cd pi-for-excel
git fetch origin
git switch --track origin/feature/enterprise-windows-local-smoke
```

If you already cloned the fork, run `git fetch origin` and switch to the
feature branch instead. Do not use `main` or the upstream production
manifest for this test.

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
  except the approved API origin and the Office.js CDN where applicable.
  CSP compatibility must be validated in the installed Excel WebView.
- No Python, tmux, MCP or external proxy server is needed for this test.

## 4. Sideload into Excel for Windows

### Optional fixed-destination local relay for APIs without CORS

For a synthetic local test, use `-UseLocalRelay` with the launch script if
the approved API does not accept browser CORS requests. This changes the
transport: the loopback HTTPS server forwards only `GET /models` and
`POST /chat/completions` to the configured HTTPS API. It does not follow
redirects, accept user-supplied targets, forward cookies, or log credentials.
POST requires the exact local Origin and JSON content type. The relay binds
to IPv4 loopback, applies size/time limits and preserves streamed responses.

```powershell
.\scripts\start-enterprise-local.ps1 -ApiBaseUrl "https://api.example.invalid/llm/v1" -SkipInstall -UseLocalRelay
```

In the custom gateway settings, use the base URL
`https://127.0.0.1:3141/__enterprise-api`, the approved model ID, and the
runtime API credential. The restricted build is pinned to that local base
URL; the server is independently pinned to the upstream API. Both must be
started with the matching launch option. Existing direct-API configurations
will no longer be selectable in this build. Stop and rebuild without
`-UseLocalRelay` to return to the direct transport. No upstream API change or
browser security exception is needed. This optional transport requires its
own real Excel smoke test and does not establish confidential-data approval.

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
   Supply one model ID and the runtime credential (if required).
3. Use the model selector. The restricted build may discover additional
   models through the approved API's `/models` endpoint.
4. Ask: `Which sheet is active and what is in A1:B3?`
5. Switch the **Auto / Confirm** mode from the status bar. Check both modes
   with a trivial disposable-cell edit. User confirmation remains optional.
6. Check workbook history/recovery, then repeat a request after disabling
   the API. The add-in must not fall back to a public model.

The model names may be visible without a working API if a saved catalogue
exists; that alone does not prove the endpoint or CORS integration works.

If the thinking selector offers only Off for a custom model, edit the gateway
and set **Reasoning support** to **Enabled**, then save and select the model
again. This opts into the OpenAI `reasoning_effort` request parameter; it does
not establish that the gateway implements that parameter. Choose High for an
initial synthetic test. Off sends `none` for an explicitly enabled unknown
model. Automatic retains registered-model detection and defaults unknown
models to reasoning disabled. Validate actual reasoning output with your API.

The gateway form also exposes **Max context tokens** and **Max output tokens**.
Set these to the limits of your actual API deployment (for example, 262144
context and 16000 output tokens for a deployment configured with those limits).
The output limit must be a positive integer no greater than the context limit.
Leaving output blank uses matched registry metadata, or 4096 for an unknown
model, clamped to the context limit. Saving an explicit value overrides that
default and preserves it across discovery and restart.

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

**Known gaps:** There has not yet been a Windows Excel end-to-end smoke test.
The full set of browser transports and extension-adjacent paths is not a
certified egress boundary. Review dependencies, API-key persistence and
organizational controls before authorizing confidential workbooks.
