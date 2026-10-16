/**
 * CSP for self-hosted, restricted Excel taskpanes.
 * Host administrators must enforce this as a response HEADER, not merely a
 * meta tag. The endpoint value is the non-secret API base URL.
 */
export function buildEnterpriseContentSecurityPolicy(apiBaseUrl) {
  if (!apiBaseUrl || typeof apiBaseUrl !== "string") {
    throw new Error("ENTERPRISE_API_BASE_URL must be configured.");
  }
  let parsed;
  try {
    parsed = new URL(apiBaseUrl);
  } catch {
    throw new Error("ENTERPRISE_API_BASE_URL must be a valid HTTPS URL.");
  }
  if (parsed.protocol !== "https:" || parsed.username || parsed.password ||
      parsed.search || parsed.hash || !parsed.hostname) {
    throw new Error("ENTERPRISE_API_BASE_URL must be HTTPS without credentials, query or fragment.");
  }

  return [
    "default-src 'none'",
    "script-src 'self' https://appsforoffice.microsoft.com",
    "style-src 'self' 'unsafe-inline'",
    "font-src 'self' data:",
    "img-src 'self' data: blob:",
    `connect-src 'self' https://appsforoffice.microsoft.com ${parsed.origin}`,
    "frame-src 'none'",
    "worker-src 'none'",
    "object-src 'none'",
    "base-uri 'none'",
    "form-action 'none'",
  ].join("; ");
}
