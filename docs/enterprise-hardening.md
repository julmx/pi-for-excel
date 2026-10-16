# Enterprise hardening: first step (not production approval)

This fork can build a **restricted agent tool list** with `VITE_PI_ENTERPRISE_MODE=true`.

## What this initial change does

When enabled at build time, the agent receives only explicitly allowlisted
first-party Excel tools from `src/tools/enterprise-policy.ts`. Terminal/tmux,
Python, arbitrary Office.js/WPS JS, file workspace, skills, integration tools,
MCP, web search, external fetch, and extension-provided tools are **not included
in the agent's advertised executable tool list**. The standard build is unchanged.

For a local proof of concept with **synthetic data only**:

```bash
npm ci
VITE_PI_ENTERPRISE_MODE=true npm run build
npm run test:context
```

## Still required before confidential use

1. **Model/endpoint enforcement:** the user interface's provider filter is not a
   security boundary. Make provider selection and outbound model calls fail
   closed for everything except the company-approved endpoint (all models served by that endpoint are permitted).
2. **Network policy:** self-host the add-in using an enterprise-controlled
   origin and configure a strict Content-Security-Policy and egress controls.
   The upstream Vercel deployment and default CSP are not acceptable for this
   scenario. Check CORS behavior of the internal API or use an authenticated,
   strictly allowlisted internal proxy.
3. **Other execution surfaces:** investigate and disable extension loading,
   custom commands, code paths not dispatched as agent tools, optional bridges,
   and arbitrary third-party content. Tool filtering does not disable these
   subsystems by itself.
4. **Mutation approval:** force confirmation mode in the deployed build and
   prevent user or persisted settings from silently switching to auto.
5. **Secrets and logs:** audit IndexedDB credentials, error reporting,
   remote assets, prompts, and local session history. Do not place live keys
   or company-private URLs in a public fork.
6. **Validation:** run tests and a network capture in Excel on the actual
   deployment artifact. Check that API outage fails closed with no public
   provider fallback.

**This change is a partial defense, not a security certification.** Do not
use the fork with confidential workbooks until the remaining controls are
implemented and validated by the organization.

## Step 2: approved internal LLM endpoint (draft)

In the enterprise build, configure an exact **HTTPS OpenAI-compatible base
URL** ; all models exposed by that endpoint are allowed. The values below are
examples, not a corporate endpoint:

```bash
VITE_PI_ENTERPRISE_MODE=true \
VITE_PI_ENTERPRISE_API_BASE_URL=https://api.example.invalid/llm/v1 \
npm run build
```

**Never put an API key in a VITE_ variable** (Vite embeds it in browser JS).
Provision the credential through an approved runtime authentication mechanism.

The enterprise build now:
- refuses to initialize if its endpoint is missing;
- validates selected model's API type, provider, full base URL at
  the agent stream boundary, before forwarding the prompt;
- avoids optional user-configured CORS proxies for those LLM calls;
- filters the model selection dialog to models from the approved endpoint.

The approved model must be configured as a **custom OpenAI-compatible gateway**
with an identical base URL ; no per-model allowlist is required. A previously selected public
model will be refused, not silently substituted.

**Important limitations:** This is not yet a global outbound network policy.
Other fetch paths, model refresh/discovery, authentication, startup assets,
extensions and non-agent model calls may still reach external hosts. The model
selector is only a convenience; the stream check is the agent-call guard.
The endpoint/model values are visible in the built client and must not be
treated as secret. Corporate API/browser CORS support still needs verification.
Do not use confidential workbooks until a strict CSP, network egress policy,
full call-path audit and Excel tests are complete.

Suggested verification on synthetic data: `npm run check`,
`npm run build`, `npm run test:context`, `npm run test:security`, and
`node --test --experimental-strip-types --import ./scripts/register-test-ts-loader.mjs tests/enterprise-model-policy.test.ts`.

## Approved model transport (second phase)

Opt-in builds now require these **non-secret** build settings:

```bash
VITE_PI_ENTERPRISE_MODE=true \
VITE_PI_ENTERPRISE_API_BASE_URL=https://example.invalid/llm/v1 \
VITE_PI_ENTERPRISE_MODEL_IDS=approved-model \
npm run build
```

The configured base URL must be HTTPS and may not contain embedded credentials,
query parameters or fragments. The list of model IDs must be explicit. Missing
or malformed configuration fails during taskpane startup.

The agent stream rejects models unless they use `openai-completions`, have a
custom gateway provider name, match an approved model ID and use the exact
configured API base URL. Approved requests deliberately bypass the optional
generic Pi proxy and go directly to the corporate API. Therefore **CORS must
allow the hosted taskpane's origin**; a separate trusted internal gateway may be
needed in the future. The model chooser hides unapproved models.

Configure the gateway and its API key through the existing custom gateway UI
with the approved URL/model. Never put API keys in `VITE_*` variables; their
values become part of the public browser bundle.

**Remaining risks:** built-in provider discovery and user-configured settings
still exist; the startup provider catalogue is not entirely restricted. Other
subsystems may make network requests outside the agent stream path. A strict
CSP, approved hosting, full fetch/egress policy, disabling remote features,
credential handling, end-to-end tests and network inspection are still required.
This is *not* safe for confidential workbooks yet.
