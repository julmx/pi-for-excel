import assert from "node:assert/strict";
import { test } from "node:test";

import { buildEnterpriseContentSecurityPolicy, buildEnterprisePythonRunnerCsp } from "../scripts/enterprise-csp.mjs";

void test("enterprise CSP permits the internal API and mandatory Office.js CDN only", () => {
  const csp = buildEnterpriseContentSecurityPolicy("https://llm.example.invalid/v1");
  assert.match(csp, /default-src 'none'/u);
  assert.match(csp, /connect-src 'self' https:\/\/appsforoffice\.microsoft\.com https:\/\/llm\.example\.invalid/u);
  assert.match(csp, /script-src 'self' blob: 'wasm-unsafe-eval' https:\/\/appsforoffice\.microsoft\.com/u);
  assert.doesNotMatch(csp, /'unsafe-eval'|script-src[^;]*https:\/\/\*/u);
  assert.match(csp, /frame-src 'none'/u);
  assert.doesNotMatch(csp, /api\.openai\.com|api\.anthropic\.com|vercel\.app|jsdelivr\.net/u);
});

void test("enterprise CSP rejects non-HTTPS URLs and embedded credentials", () => {
  for (const url of [
    "",
    "http://llm.example.invalid/v1",
    "https://person:secret@llm.example.invalid/v1",
    "https://llm.example.invalid/v1?secret=example",
    "https://llm.example.invalid/v1#fragment",
  ]) {
    assert.throws(() => buildEnterpriseContentSecurityPolicy(url));
  }
});

void test("isolated runner CSP never receives model gateway access", () => {
  const parent = buildEnterpriseContentSecurityPolicy(
    "https://private.example.invalid/v1", "https://localhost:3142",
  );
  const runner = buildEnterprisePythonRunnerCsp();
  const worker = buildEnterprisePythonRunnerCsp(true);
  assert.match(parent, /frame-src https:\/\/localhost:3142/u);
  assert.match(runner, /frame-ancestors https:\/\/127\.0\.0\.1:3141/u);
  assert.match(runner, /connect-src 'self'/u);
  assert.match(worker, /worker-src 'none'/u);
  assert.match(runner, /worker-src 'self'/u);
  assert.doesNotMatch(runner, /appsforoffice|private\.example|unsafe-inline|'unsafe-eval'|blob:/u);
});


void test("enterprise CSP allows every approved origin but no unlisted provider", () => {
  const csp = buildEnterpriseContentSecurityPolicy([
    "https://private-a.example.invalid/v1",
    "https://private-b.example.invalid/api",
  ]);
  assert.match(csp, /connect-src 'self' https:\/\/appsforoffice\.microsoft\.com https:\/\/private-a\.example\.invalid https:\/\/private-b\.example\.invalid/u);
  assert.doesNotMatch(csp, /unknown\.example\.invalid|connect-src[^;]*\*/u);
  assert.throws(() => buildEnterpriseContentSecurityPolicy([
    "https://private-a.example.invalid/v1", "http://untrusted.example.invalid/api",
  ]));
});
