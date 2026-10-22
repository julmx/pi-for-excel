/**
 * A loopback-only, pinned-upstream relay for enterprise Excel smoke tests.
 *
 * The HTTPS parent server MUST check that the client is on loopback before
 * calling this handler. This relay does not accept a user-provided URL and
 * never reads or writes credentials to disk. TLS validation is left enabled.
 */
import https from "node:https";

const RELAY_PREFIX = "/__enterprise_api/";
const ENDPOINTS = new Map([
  ["models", "GET"],
  ["chat/completions", "POST"],
]);

export function resolveEnterpriseRelayTarget(apiBaseUrls, path, method) {
  if (typeof path !== "string" || typeof method !== "string") return null;
  if (!path.startsWith(RELAY_PREFIX)) return null;

  const routes = Array.isArray(apiBaseUrls) ? apiBaseUrls : [apiBaseUrls];
  const subpath = path.slice(RELAY_PREFIX.length);
  const indexed = /^(\d+)\/(models|chat\/completions)$/u.exec(subpath);
  // Preserve the historical unindexed route for the first gateway.
  const index = indexed ? Number(indexed[1]) : 0;
  const suffix = indexed ? indexed[2] : subpath;
  if (!Number.isSafeInteger(index) || index < 0 || index >= routes.length) return null;
  if (indexed && index === 0) return null;
  if (ENDPOINTS.get(suffix) !== method.toUpperCase()) return null;

  const upstream = new URL(routes[index]);
  if (upstream.protocol !== "https:" || !upstream.hostname ||
      upstream.username || upstream.password || upstream.search || upstream.hash) {
    throw new Error("Enterprise relay requires an approved HTTPS API base URL.");
  }

  // Assign the pathname on the already validated URL. Constructing a new
  // URL from a path beginning with // could otherwise change its origin.
  upstream.pathname = `${upstream.pathname.replace(/\/$/u, "")}/${suffix}`;
  return upstream;
}

export function handleEnterpriseApiRelay(req, res, apiBaseUrls, path) {
  if (!path.startsWith(RELAY_PREFIX)) return false;

  let target;
  try {
    target = resolveEnterpriseRelayTarget(apiBaseUrls, path, req.method || "");
  } catch {
    res.writeHead(500, { "content-type": "application/json", "cache-control": "no-store" });
    res.end(JSON.stringify({ error: { message: "Enterprise relay configuration is invalid." } }));
    return true;
  }
  if (!target) {
    res.writeHead(404, { "content-type": "application/json", "cache-control": "no-store" });
    res.end(JSON.stringify({ error: { message: "Enterprise relay route not found." } }));
    return true;
  }

  const headers = { "accept-encoding": "identity" };
  for (const key of ["authorization", "x-api-key", "api-key", "accept", "content-type", "content-length"]) {
    const value = req.headers[key];
    if (typeof value === "string") headers[key] = value;
  }

  const upstream = https.request(target, {
    method: req.method,
    headers,
    timeout: 180_000,
  }, (response) => {
    console.log(`[enterprise-api-relay] ${req.method} ${path} -> ${response.statusCode || 502}`);
    if (res.destroyed) {
      response.destroy();
      return;
    }
    const out = {
      "cache-control": "no-store",
      "x-content-type-options": "nosniff",
    };
    for (const key of ["content-type", "content-encoding"]) {
      const value = response.headers[key];
      if (typeof value === "string") out[key] = value;
    }
    res.writeHead(response.statusCode || 502, out);
    response.pipe(res);
  });

  upstream.on("timeout", () => upstream.destroy(new Error("upstream_timeout")));
  upstream.on("error", (error) => {
    // Never log API request bodies, authorization headers or raw provider errors.
    console.error("[enterprise-api-relay] Upstream request failed:", error.code || "request_error");
    if (!res.headersSent) {
      res.writeHead(502, { "content-type": "application/json", "cache-control": "no-store" });
      res.end(JSON.stringify({ error: { message: "Enterprise API unreachable from local HTTPS server." } }));
    } else {
      res.destroy();
    }
  });

  req.on("aborted", () => upstream.destroy());
  res.on("close", () => {
    if (!res.writableEnded) upstream.destroy();
  });
  req.pipe(upstream);
  return true;
}
