#!/usr/bin/env node

/**
 * Serve the Vite production build (dist/) over HTTPS.
 *
 * Useful for Office add-in smoke tests where the manifest must load from HTTPS
 * but we want to exercise the *production* code path (import.meta.env.DEV === false).
 *
 * Requires mkcert-generated certs in the repo root:
 *   - cert.pem
 *   - key.pem
 */

import https from "node:https";
import fs from "node:fs";
import path from "node:path";
import {
  buildEnterpriseContentSecurityPolicy,
  buildEnterprisePythonRunnerCsp,
  ENTERPRISE_LOCAL_RUNNER_ORIGIN,
} from "./enterprise-csp.mjs";
import { handleEnterpriseApiRelay } from "./enterprise-api-relay.mjs";

const HOST = process.env.HOST || "localhost";
const PORT = Number.parseInt(process.env.PORT || "3141", 10);

// For restricted local smoke tests, require a separate serving-time policy.
// Production reverse proxies must supply an equivalent CSP response header.
const enterpriseCsp = process.env.ENTERPRISE_MODE === "1"
  ? buildEnterpriseContentSecurityPolicy(process.env.ENTERPRISE_API_BASE_URL, ENTERPRISE_LOCAL_RUNNER_ORIGIN)
  : null;

const rootDir = path.resolve(process.cwd());
const distDir = path.join(rootDir, "dist");

const keyPath = path.join(rootDir, "key.pem");
const certPath = path.join(rootDir, "cert.pem");

if (!fs.existsSync(distDir)) {
  console.error(`[serve-dist-https] dist/ not found at ${distDir}. Run: npm run build`);
  process.exit(1);
}

if (!fs.existsSync(keyPath) || !fs.existsSync(certPath)) {
  console.error("[serve-dist-https] Missing HTTPS certs (key.pem/cert.pem) in repo root.");
  console.error("Generate them with mkcert (see README Quick Start). ");
  process.exit(1);
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".wasm": "application/wasm",
  ".whl": "application/octet-stream",
  ".zip": "application/zip",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
  ".woff": "font/woff",
  ".ttf": "font/ttf",
  ".txt": "text/plain; charset=utf-8",
};

function isLoopbackAddress(addr) {
  if (!addr) return false;
  if (addr === "::1" || addr === "0:0:0:0:0:0:0:1") return true;
  if (addr.startsWith("127.")) return true;
  if (addr.startsWith("::ffff:127.")) return true;
  return false;
}

function safeJoin(base, reqPath) {
  const decoded = decodeURIComponent(reqPath);
  const cleaned = decoded.replace(/^\/+/, "");
  const full = path.resolve(base, cleaned);

  // Prevent path traversal. String-prefix checks are insufficient (e.g. /dist2).
  const rel = path.relative(base, full);
  if (rel.startsWith("..") || path.isAbsolute(rel)) {
    throw new Error("Path traversal");
  }

  return full;
}

const serveRequest = (req, res) => {
    try {
      const remote = req.socket?.remoteAddress;
      if (!isLoopbackAddress(remote)) {
        res.statusCode = 403;
        res.setHeader("content-type", "text/plain; charset=utf-8");
        console.log(`[serve-dist-https] ${req.method || "GET"} ${req.url || "/"} -> 403 (remote ${remote || "unknown"})`);
        res.end("forbidden");
        return;
      }

      const url = new URL(req.url || "/", `https://${HOST}:${PORT}`);
      let reqPath = url.pathname;
      if (reqPath === "/") reqPath = "/src/taskpane.html";

      // In the enterprise smoke build the taskpane and runner share a
      // Node process with different loopback origins, not the same browser storage.
      // Runner host serves only static runner assets, never the API relay.
      const host = req.headers.host?.toLowerCase();
      const runnerHost = enterpriseCsp && host === "localhost:3142" && req.socket.localPort === 3142;
      const taskpaneHost = enterpriseCsp && host === "127.0.0.1:3141" && req.socket.localPort === 3141;
      if (enterpriseCsp && !runnerHost && !taskpaneHost) {
        res.writeHead(403);
        res.end("Unrecognized enterprise host");
        return;
      }
      if (runnerHost) {
        const runnerAsset = reqPath === "/enterprise-runner.html"
          || reqPath.startsWith("/assets/")
          || reqPath.startsWith("/enterprise-pyodide/");
        if (!runnerAsset || (req.method !== "GET" && req.method !== "HEAD")) {
          res.writeHead(403);
          res.end("Runner cannot access taskpane resources");
          return;
        }
      }
      if (taskpaneHost && reqPath === "/enterprise-runner.html") {
        res.writeHead(403);
        res.end("Runner is only served on the isolated origin");
        return;
      }

      // Optional same-origin forwarding, restricted to the approved API
      // and permitted only after the loopback check above.
      if (taskpaneHost && process.env.ENTERPRISE_LOCAL_API_RELAY === "1") {
        if (handleEnterpriseApiRelay(req, res, process.env.ENTERPRISE_API_BASE_URL, reqPath)) {
          return;
        }
      }

      const filePath = safeJoin(distDir, reqPath);
      if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
        res.statusCode = 404;
        res.setHeader("content-type", "text/plain; charset=utf-8");
        console.log(`[serve-dist-https] ${req.method || "GET"} ${reqPath} -> 404`);
        res.end("not found");
        return;
      }

      if (enterpriseCsp) {
        // Python's worker loads local WASM/wheels but cannot contact the LLM
        // gateway, even though the taskpane is authorized to do so.
        const pythonWorker = /^\/assets\/enterprise-python-worker-[\w-]+\.js$/u.test(reqPath);
        // The taskpane is NOT allowed to host the Python runner's code.
        // In particular the runner cannot use the taskpane's API relay.
        const workerCsp = enterpriseCsp
          .replace(/connect-src[^;]*/u, "connect-src 'self'")
          .replace(/worker-src[^;]*/u, "worker-src 'none'");
        res.setHeader("Content-Security-Policy", runnerHost
          ? buildEnterprisePythonRunnerCsp(pythonWorker)
          : pythonWorker ? workerCsp : enterpriseCsp);
        res.setHeader("Referrer-Policy", "no-referrer");
        res.setHeader("X-Content-Type-Options", "nosniff");
      }

      const ext = path.extname(filePath).toLowerCase();
      res.statusCode = 200;
      res.setHeader("content-type", MIME[ext] || "application/octet-stream");
      console.log(`[serve-dist-https] ${req.method || "GET"} ${reqPath} -> 200`);

      // Cache hashed assets aggressively, keep HTML uncached
      if (reqPath.startsWith("/assets/") && /-[A-Za-z0-9]{8,}\./.test(reqPath)) {
        res.setHeader("cache-control", "public, max-age=31536000, immutable");
      } else {
        res.setHeader("cache-control", "no-cache");
      }

      fs.createReadStream(filePath).pipe(res);
    } catch (e) {
      res.statusCode = 500;
      res.setHeader("content-type", "text/plain; charset=utf-8");
      console.log(`[serve-dist-https] ${req.method || "GET"} ${req.url || "/"} -> 500`);
      res.end(`server error: ${e instanceof Error ? e.message : String(e)}`);
    }
};

const tlsCredentials = {
  key: fs.readFileSync(keyPath),
  cert: fs.readFileSync(certPath),
};
const server = https.createServer(tlsCredentials, serveRequest);
const runnerServer = enterpriseCsp ? https.createServer(tlsCredentials, serveRequest) : null;

// Enterprise smoke tests use an IPv4 loopback-only manifest and socket.
// Other development flows retain the previous listener behavior.
const listenHost = enterpriseCsp ? "127.0.0.1" : undefined;
server.listen(PORT, listenHost, () => {
  console.log(`[serve-dist-https] Serving ${distDir}`);
  const displayHost = listenHost ?? HOST;
  console.log(`[serve-dist-https] https://${displayHost}:${PORT}/src/taskpane.html`);
});

if (runnerServer) {
  // A second HTTPS socket in the SAME process: the isolated origin never
  // has a separate IndexedDB origin. Historical localhost cookies may be shared
  // across ports: this is for synthetic testing pending cookie review.
  runnerServer.listen(3142, "127.0.0.1", () => {
    console.log("[serve-dist-https] Python runner: https://localhost:3142/enterprise-runner.html");
  });
}
