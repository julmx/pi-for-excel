import https from "node:https";

const PREFIX = "/__enterprise-api";
const MAX_BODY_BYTES = 8 * 1024 * 1024;

/** Fixed-destination, loopback-only transport for the Windows synthetic smoke test. */
export function createEnterpriseLocalRelay({ apiBaseUrl, localOrigin, request = https.request }) {
  const api = new URL(apiBaseUrl);
  const local = new URL(localOrigin);
  if (api.protocol !== "https:" || api.username || api.password || api.search || api.hash
      || local.origin !== localOrigin || local.protocol !== "https:" || local.hostname !== "127.0.0.1") {
    throw new Error("Invalid enterprise local relay configuration.");
  }
  const base = api.href.replace(/\/$/u, "");
  return async function relay(req, res) {
    if (!req.url?.startsWith(PREFIX)) return false;
    const fail = (status, message) => {
      if (res.headersSent) { res.destroy(); return; }
      res.writeHead(status, { "Content-Type": "application/json", "Cache-Control": "no-store" });
      res.end(JSON.stringify({ error: message }));
    };
    if (req.socket.remoteAddress !== "127.0.0.1" || req.headers.host !== local.host
        || (req.headers.origin && req.headers.origin !== localOrigin)
        || (req.headers["sec-fetch-site"] && req.headers["sec-fetch-site"] !== "same-origin")) {
      fail(403, "Local relay origin rejected.");
      return true;
    }
    const suffix = req.url.slice(PREFIX.length);
    const permitted = (req.method === "GET" && suffix === "/models")
      || (req.method === "POST" && suffix === "/chat/completions");
    if (!permitted) { fail(404, "Local relay route rejected."); return true; }
    if (req.method === "POST" && (req.headers.origin !== localOrigin
        || req.headers["content-type"]?.split(";")[0]?.trim().toLowerCase() !== "application/json")) {
      fail(403, "Local relay requires a same-origin JSON request.");
      return true;
    }
    const headers = { accept: req.headers.accept || "application/json" };
    if (req.headers.authorization) headers.authorization = req.headers.authorization;
    if (req.method === "POST") headers["content-type"] = "application/json";
    try {
      const chunks = [];
      let size = 0;
      for await (const chunk of req) {
        size += chunk.length;
        if (size > MAX_BODY_BYTES) { fail(413, "Local relay request too large."); return true; }
        chunks.push(chunk);
      }
      if (res.destroyed) return true;
      await new Promise((resolve) => {
        const upstream = request(new URL(base + suffix), { method: req.method, headers }, (response) => {
          // Never follow redirects, forward cookies, or expose the upstream headers wholesale.
          if (response.statusCode >= 300 && response.statusCode < 400) {
            response.resume();
            fail(502, "Local relay refused an upstream redirect.");
            resolve();
            return;
          }
          res.writeHead(response.statusCode || 502, {
            "Content-Type": response.headers["content-type"] || "application/json",
            ...(response.headers["content-encoding"] ? { "Content-Encoding": response.headers["content-encoding"] } : {}),
            "Cache-Control": "no-store",
            "X-Content-Type-Options": "nosniff",
          });
          response.on("error", () => { fail(502, "Local relay response interrupted."); resolve(); });
          response.on("end", resolve);
          response.pipe(res);
        });
        const deadline = setTimeout(() => upstream.destroy(), 300_000);
        upstream.setTimeout(120_000, () => upstream.destroy());
        upstream.on("error", () => { fail(502, "Local relay could not reach the approved API."); resolve(); });
        const cancel = () => upstream.destroy();
        res.once("close", cancel);
        upstream.once("close", () => { clearTimeout(deadline); res.off("close", cancel); resolve(); });
        upstream.end(Buffer.concat(chunks));
      });
    } catch {
      fail(502, "Local relay request interrupted.");
    }
    return true;
  };
}
