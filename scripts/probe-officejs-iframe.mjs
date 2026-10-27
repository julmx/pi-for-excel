#!/usr/bin/env node
/**
 * Synthetic-only feasibility probe for Office.js inside a cross-origin iframe.
 * Separate from the Enterprise taskpane; does not modify src/ or its CSP.
 *
 * Usage: node scripts/probe-officejs-iframe.mjs
 */
import fs from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";

const root = process.cwd();
const probeDir = path.join(root, "probes", "officejs-iframe");
const parentOrigin = "https://127.0.0.1:3151";
const childOrigin = "https://localhost:3152";
const key = fs.readFileSync(path.join(root, "key.pem"));
const cert = fs.readFileSync(path.join(root, "cert.pem"));
const manifestPath = path.join(os.tmpdir(), "pi-officejs-iframe-probe.xml");

const original = fs.readFileSync(path.join(root, "manifest.xml"), "utf8");
const manifest = original
  .replace(/<Id>[^<]+<\/Id>/u, "<Id>467515a2-c49f-4fbd-893a-60dfd021512f</Id>")
  .replaceAll("https://localhost:3141", parentOrigin)
  .replaceAll("Pi for Excel", "Pi Office.js Probe")
  .replaceAll("/src/taskpane.html", "/host.html")
  .replace("  <Hosts>",
    "  <AppDomains>\n    <AppDomain>" + childOrigin
    + "</AppDomain>\n  </AppDomains>\n  <Hosts>");
if (!manifest.includes("<AppDomain>" + childOrigin + "</AppDomain>")
    || !manifest.includes(parentOrigin + "/host.html")) {
  throw new Error("Probe manifest generation failed");
}
fs.writeFileSync(manifestPath, manifest);

const parentCsp = [
  "default-src 'none'",
  "script-src 'self' https://appsforoffice.microsoft.com",
  "connect-src 'self' https://appsforoffice.microsoft.com",
  "frame-src " + childOrigin,
  "style-src 'unsafe-inline'",
  "img-src 'self'",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'"
].join("; ");
const childCsp = [
  "default-src 'none'",
  "script-src 'self' blob: https://appsforoffice.microsoft.com",
  "connect-src 'self' https://appsforoffice.microsoft.com",
  "style-src 'unsafe-inline'",
  "img-src 'none'",
  "worker-src 'none'",
  "frame-src 'none'",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
  "frame-ancestors " + parentOrigin
].join("; ");

function serve(isChild, port) {
  return https.createServer({key, cert}, (req, res) => {
    const allowedHost = isChild ? "localhost:3152" : "127.0.0.1:3151";
    if (req.headers.host !== allowedHost || !["GET", "HEAD"].includes(req.method)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }
    const pathname = new URL(req.url, "https://" + allowedHost).pathname;
    const files = isChild
      ? {"/runner.html": "runner.html", "/runner.js": "runner.js"}
      : {"/host.html": "host.html", "/host.js": "host.js"};
    let body;
    let type = "text/plain; charset=utf-8";
    if (!isChild && pathname === "/manifest.xml") {
      body = Buffer.from(manifest);
      type = "application/xml; charset=utf-8";
    } else if (Object.hasOwn(files, pathname)) {
      body = fs.readFileSync(path.join(probeDir, files[pathname]));
      type = pathname.endsWith(".html") ? "text/html; charset=utf-8" : "text/javascript; charset=utf-8";
    } else if (!isChild && /^\/assets\/icon-(16|32|80)\.png$/u.test(pathname)) {
      // The probe uses only the existing project's icon resources.
      const icon = path.join(root, "public", pathname.slice(1));
      if (fs.existsSync(icon)) {
        body = fs.readFileSync(icon);
        type = "image/png";
      }
    }
    if (!body) {
      res.writeHead(404);
      res.end("Not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": type,
      "Content-Security-Policy": isChild ? childCsp : parentCsp,
      "Cache-Control": "no-store",
      "Referrer-Policy": "no-referrer",
      "X-Content-Type-Options": "nosniff"
    });
    if (req.method === "HEAD") res.end();
    else res.end(body);
  }).listen(port, "127.0.0.1");
}
const parent = serve(false, 3151);
const child = serve(true, 3152);
process.on("SIGINT", () => {
  parent.close();
  child.close();
  process.exit(0);
});
console.log("Synthetic-only Office.js iframe probe.");
console.log("Sideload manifest: " + manifestPath);
console.log("Parent: " + parentOrigin + "/host.html");
console.log("Child : " + childOrigin + "/runner.html");
console.log("Only use a disposable workbook. No Enterprise setting has changed.");
