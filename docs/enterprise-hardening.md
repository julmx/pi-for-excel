# Enterprise deployment hardening (work in progress)

For a Windows Excel local proof of concept, follow the
[Windows local HTTPS smoke-test guide](./enterprise-windows-smoke.md).

> **Do not use confidential workbooks yet.** This is an opt-in technical
> prototype, not a reviewed or certified enterprise security control.

## Intended architecture

Self-host the Excel taskpane under a company-controlled HTTPS origin.
Use exactly one approved HTTPS OpenAI-compatible API base URL. **All models
served by that API are permitted**, without a separate model-ID allowlist.
The corporate API key must never be committed or included in a `VITE_*`
environment variable, since those variables are embedded in client bundles.

Build the restricted client with a non-secret example endpoint:

```bash
npm ci
VITE_PI_ENTERPRISE_MODE=true \
VITE_PI_ENTERPRISE_API_BASE_URL=https://api.example.invalid/llm/v1 \
npm run build
npm run test:enterprise
```

The build fails during startup if the endpoint is absent or malformed.
The default build (without `VITE_PI_ENTERPRISE_MODE=true`) retains upstream
behavior.

## What the prototype currently restricts

- **Agent tools:** Only reviewed, first-party Excel operations and the
  enterprise-only `python_excel` tool are exposed to the agent. The Python
  Worker reuses these Excel tools through allowlisted RPC; terminal, native
  Python bridges, arbitrary Office.js, MCP, web search, external integrations,
  skills and extension tools remain excluded from the restricted agent surface.
  Python itself is **not a security sandbox**: same-origin Worker JavaScript
  and IndexedDB access need isolation before confidential workbook use.
  See [the focused 2026-10-18 security review](./enterprise-security-review-2026-10-18.md).
- **Extensions and local bridges:** Enterprise startup disables persisted
  extension loading and installation/activation through the extension manager.
  The extension, add-on, tools, skills, files, experimental and debug slash
  commands are excluded. Local bridge discovery and the development background
  verification bridge are not started in enterprise mode.
- **Workbook confirmation remains configurable:** The existing Auto and
  Confirm execution modes are both supported. The chosen mode is persisted;
  neither per-edit approval nor automatic edits are forcibly selected by
  the enterprise build.
- **Agent model transport:** The main agent stream accepts only custom
  `openai-completions` gateways with the **exact approved base URL**.
  The request cannot silently fall back to a public model or user-configured
  external proxy.
- **Startup catalogue:** Public built-in providers are not registered in
  restricted mode. An unauthenticated, non-operational bootstrap entry lets
  the taskpane start before a corporate gateway is configured. Public provider
  login options are hidden on the initial welcome screen.
- **Gateway discovery:** After adding a corporate custom gateway (including
  one initial model identifier), model refresh calls
  `<approved base URL>/models` directly, using the gateway credential if
  configured. The reported model IDs are merged into the selector and cached.
  Existing normal-build custom gateway behavior is unchanged.
- **Fetch defense in depth:** Browser `window.fetch` calls in restricted
  builds are limited to same-origin resources and the approved gateway's
  `/models` and `/chat/completions` paths. Other destinations fail closed.
  This is not a complete network boundary.

Model catalogue parsing is limited (size, count and ID length). Discovered
models initially inherit the manually configured model's metadata such as
context window, so different model limits should be reviewed separately.

If the API lacks `/models`, the manually configured model can still be
used, but automatic discovery will not work.

## Required HTTPS response headers

`scripts/enterprise-csp.mjs` produces a restrictive Content Security
Policy that only permits the hosted taskpane's origin, the approved API
origin and the Office.js CDN where needed.

For a **loopback-only smoke test of the production build**, configure the
server's policy independently of the Vite build variables:

```bash
ENTERPRISE_MODE=1 \
ENTERPRISE_API_BASE_URL=https://api.example.invalid/llm/v1 \
npm run serve:dist
```

The serving script rejects missing/invalid endpoint configuration and sends
CSP, no-referrer and nosniff headers. In an actual corporate deployment,
**the enterprise reverse proxy or web server must apply an equivalent CSP
header**. Do not deploy the upstream Vercel manifest or upstream broad CSP.

This first CSP may block optional Office.js behavior, extensions and remote
assets; test with Excel for Windows/Web and revise only after network review.
Office.js itself is still loaded from Microsoft's CDN; this solution does not
make Microsoft 365 fully offline.

## Verification

```bash
npm run check
npm run build
npm run test:enterprise
npm run test:models
npm run test:context
npm run test:security
```

Use only synthetic data for an Excel smoke test. Capture and review outbound
network traffic during initialization, model refresh, direct calls, tool
execution, and API failure. Confirm that there is **no fallback** to a public
LLM or proxy if the corporate API or credential fails.

## Known gaps before production

A focused [code-level review dated 2026-10-18](./enterprise-security-review-2026-10-18.md)
documents the current Python Worker storage boundary, resource and payload
limits, build integrity, local relay and upstream maintenance gaps. It is
not a penetration test or production approval.

1. A `fetch` interceptor is not a security boundary. Other transports
   (XHR, WebSockets, resources loaded by the browser, captured references,
   code extensions) need separate review; CSP plus infrastructure egress
   controls and an approved internal hosting origin remain mandatory.
2. Runtime entry points beyond the main agent stream may still execute
   model calls or arbitrary code. Review extensions, commands, Python and
   local bridges, and disable unwanted subsystems completely.
3. Credentials and session history are currently stored in IndexedDB.
   Review token handling, XSS boundaries, logging, retention and data
   minimization, potentially replacing direct browser API keys with an
   enterprise-authenticated internal gateway.
4. Preserve both configurable execution modes. Validate that Confirm asks
   before workbook mutations and Auto makes changes without per-edit prompts,
   while maintaining workbook history/recovery and clear user control.
5. Validate CORS and Office.js compatibility, the exact gateway protocol,
   capabilities of individual discovered models (especially tool calling),
   full CI checks, and network traces before allowing confidential data.
