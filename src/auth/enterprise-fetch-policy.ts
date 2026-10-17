/**
 * Defense-in-depth policy for fetch requests in enterprise builds.
 *
 * Network-level CSP and infrastructure egress controls are still required:
 * this cannot intercept XMLHttpRequest, image/script loads, or captured fetch.
 */
import type { EnterpriseModelPolicy } from "./enterprise-model-policy.js";

export function isEnterpriseFetchUrlAllowed(
  policy: EnterpriseModelPolicy,
  requestedUrl: string,
  taskpaneOrigin: string,
): boolean {
  try {
    const taskpane = new URL(taskpaneOrigin);
    const request = new URL(requestedUrl, taskpane);
    if (request.username || request.password) return false;
    if (request.protocol !== "https:") return false;
    if (request.origin === taskpane.origin) return true;

    const api = new URL(policy.baseUrl);
    if (request.origin !== api.origin) return false;
    if (request.search || request.hash) return false;

    const path = api.pathname.replace(/\/$/u, "");
    return request.pathname === `${path}/models`
      || request.pathname === `${path}/chat/completions`;
  } catch {
    return false;
  }
}

/**
 * A build opt-in for the local Windows smoke test only.
 * The approved remote URL remains validated before mapping to a same-origin
 * endpoint. The local HTTPS server decides the fixed upstream destination.
 */
export function getEnterpriseLocalRelayUrl(
  policy: EnterpriseModelPolicy,
  requestedUrl: string,
  taskpaneOrigin: string,
): string | null {
  if (!isEnterpriseFetchUrlAllowed(policy, requestedUrl, taskpaneOrigin)) return null;
  const request = new URL(requestedUrl, taskpaneOrigin);
  const api = new URL(policy.baseUrl);
  if (request.origin !== api.origin) return null;

  const basePath = api.pathname.replace(/\/$/u, "");
  const suffix = request.pathname.slice(basePath.length + 1);
  if (suffix !== "models" && suffix !== "chat/completions") return null;
  return `${new URL(taskpaneOrigin).origin}/__enterprise_api/${suffix}`;
}
