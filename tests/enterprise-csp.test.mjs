import assert from "node:assert/strict";
import { test } from "node:test";

import { buildEnterpriseContentSecurityPolicy } from "../scripts/enterprise-csp.mjs";

void test("enterprise CSP permits the internal API and mandatory Office.js CDN only", () => {
  const csp = buildEnterpriseContentSecurityPolicy("https://llm.example.invalid/v1");
  assert.match(csp, /default-src 'none'/u);
  assert.match(csp, /connect-src 'self' https:\/\/appsforoffice\.microsoft\.com https:\/\/llm\.example\.invalid/u);
  assert.match(csp, /script-src 'self' 'wasm-unsafe-eval' https:\/\/appsforoffice\.microsoft\.com/u);
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
