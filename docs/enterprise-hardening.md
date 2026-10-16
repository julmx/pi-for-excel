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
   closed for everything except the company-approved URL and model IDs.
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
