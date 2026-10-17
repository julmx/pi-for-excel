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

/**
 * The Pi AI SDK may use its own HTTP transport rather than the taskpane's
 * patched window.fetch. For a local smoke test, route streaming requests
 * explicitly to the pinned same-origin server instead of relying on an
 * unrelated global fetch interceptor.
 *
 * Callers MUST validate the original model against the enterprise model
 * policy before changing its transport URL.
 */
export function getEnterpriseStreamBaseUrl(
  policy: EnterpriseModelPolicy,
  modelBaseUrl: string,
  taskpaneOrigin: string,
): string {
  const approved = new URL(policy.baseUrl);
  const supplied = new URL(modelBaseUrl);
  const normalized = (url: URL): string => url.href.replace(/\/$/u, "");
  if (normalized(supplied) !== normalized(approved)) {
    throw new Error("Enterprise relay refused a model with an unapproved API base URL.");
  }

  const origin = new URL(taskpaneOrigin);
  if (origin.protocol !== "https:" || origin.hostname !== "127.0.0.1" || origin.port !== "3141"
      || origin.pathname !== "/" || origin.search || origin.hash || origin.username || origin.password) {
    throw new Error("Local enterprise relay requires the HTTPS loopback taskpane origin.");
  }
  return `${origin.origin}/__enterprise_api`;
}
