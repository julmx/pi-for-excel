/**
 * CSP for self-hosted, restricted Excel taskpanes.
 * Host administrators must enforce this as a response HEADER, not merely a
 * meta tag. The endpoint value is the non-secret API base URL.
 */
export const ENTERPRISE_LOCAL_TASKPANE_ORIGIN = "https://127.0.0.1:3141";
export const ENTERPRISE_LOCAL_RUNNER_ORIGIN = "https://localhost:3142";

export function buildEnterpriseContentSecurityPolicy(apiBaseUrls, runnerOrigin = null) {
  // One endpoint is still accepted for existing installations.
  const urls = Array.isArray(apiBaseUrls) ? apiBaseUrls : [apiBaseUrls];
  if (urls.length === 0 || urls.length > 8) {
    throw new Error("Enterprise CSP requires between 1 and 8 approved APIs.");
  }
  const origins = [];
  for (const raw of urls) {
    if (!raw || typeof raw !== "string") {
      throw new Error("ENTERPRISE_API_BASE_URL must be configured.");
    }
    let parsed;
    try {
      parsed = new URL(raw);
    } catch {
      throw new Error("ENTERPRISE_API_BASE_URL must be a valid HTTPS URL.");
    }
    if (parsed.protocol !== "https:" || parsed.username || parsed.password ||
        parsed.search || parsed.hash || !parsed.hostname) {
      throw new Error("ENTERPRISE_API_BASE_URL must be HTTPS without credentials, query or fragment.");
    }
    if (!origins.includes(parsed.origin)) origins.push(parsed.origin);
  }

  return [
    "default-src 'none'",
    // The original execute_office_js imports a dynamically generated blob module.
    // This relaxes script creation, NOT network destinations or unsafe-eval.
    "script-src 'self' blob: 'wasm-unsafe-eval' https://appsforoffice.microsoft.com",
    "style-src 'self' 'unsafe-inline'",
    "font-src 'self' data:",
    "img-src 'self' data: blob:",
    `connect-src 'self' https://appsforoffice.microsoft.com ${origins.join(' ')}`,
    runnerOrigin ? `frame-src ${runnerOrigin}` : "frame-src 'none'",
    "worker-src 'self'",
    "object-src 'none'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join("; ");
}

/**
 * Applied ONLY on the dedicated localhost runner origin, never on the taskpane.
 * No model gateway, Office.js CDN, API relay, or navigation is permitted.
 */
export function buildEnterprisePythonRunnerCsp(worker = false) {
  return [
    "default-src 'none'",
    "script-src 'self' 'wasm-unsafe-eval'",
    "style-src 'none'",
    "img-src 'none'",
    "font-src 'none'",
    "connect-src 'self'",
    "frame-src 'none'",
    worker ? "worker-src 'none'" : "worker-src 'self'",
    `frame-ancestors ${ENTERPRISE_LOCAL_TASKPANE_ORIGIN}`,
    "object-src 'none'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join("; ");
}
