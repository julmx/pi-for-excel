import assert from "node:assert/strict";
import { test } from "node:test";
import { getEnterpriseLocalRelayUrl } from "../src/auth/enterprise-fetch-policy.ts";
import { parseEnterpriseModelPolicy } from "../src/auth/enterprise-model-policy.ts";

const policy = parseEnterpriseModelPolicy({
  enabled: true,
  baseUrl: "https://company.example.invalid/llm/v1",
});
assert.ok(policy);

const origin = "https://127.0.0.1:3141";

void test("only the approved API destinations get same-origin relay URLs", () => {
  assert.equal(
    getEnterpriseLocalRelayUrl(policy, "https://company.example.invalid/llm/v1/models", origin),
    origin + "/__enterprise_api/models",
  );
  assert.equal(
    getEnterpriseLocalRelayUrl(policy, "https://company.example.invalid/llm/v1/chat/completions", origin),
    origin + "/__enterprise_api/chat/completions",
  );
});

void test("never relay other gateways, different paths, or credential-bearing URLs", () => {
  for (const uri of [
    "https://company.example.invalid/llm/v2/chat/completions",
    "https://attacker.invalid/llm/v1/chat/completions",
    "https://company.example.invalid/llm/v1/models?token=secret",
    "https://user:password@company.example.invalid/llm/v1/models",
    "https://127.0.0.1:3141/src/taskpane.html",
  ]) {
    assert.equal(getEnterpriseLocalRelayUrl(policy, uri, origin), null, uri);
  }
});
