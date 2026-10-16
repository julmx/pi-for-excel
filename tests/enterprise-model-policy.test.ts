import assert from "node:assert/strict";
import { test } from "node:test";

import {
  assertEnterpriseModelAllowed,
  isEnterpriseModelAllowed,
  parseEnterpriseModelPolicy,
} from "../src/auth/enterprise-model-policy.ts";

const validModel = {
  provider: "Gateway · Company",
  api: "openai-completions" as const,
  baseUrl: "https://api.example.invalid/llm/v1",
  id: "internal-model",
};

function enterprisePolicy() {
  const policy = parseEnterpriseModelPolicy({
    enabled: true,
    baseUrl: "https://api.example.invalid/llm/v1/",
  });
  assert.ok(policy);
  return policy;
}

void test("stock builds remain unaffected", () => {
  assert.equal(parseEnterpriseModelPolicy({ enabled: false }), null);
});

void test("enterprise mode refuses missing or malformed endpoint configuration", () => {
  for (const config of [
    { enabled: true },
    { enabled: true, baseUrl: "" },
    { enabled: true, baseUrl: "http://api.example.invalid/llm/v1" },
    { enabled: true, baseUrl: "https://user:pass@api.example.invalid/v1" },
    { enabled: true, baseUrl: "https://api.example.invalid/v1?token=abc" },
    { enabled: true, baseUrl: "https://api.example.invalid/v1#section" },
  ]) {
    assert.throws(() => parseEnterpriseModelPolicy(config));
  }
});

void test("any model ID is allowed when served through the approved gateway", () => {
  const policy = enterprisePolicy();
  assert.equal(isEnterpriseModelAllowed(policy, validModel), true);
  assert.equal(isEnterpriseModelAllowed(policy, { ...validModel, id: "different-model" }), true);
  assert.equal(isEnterpriseModelAllowed(policy, { ...validModel, id: "new-company-model" }), true);
});

void test("public providers, non-OpenAI APIs and alternate endpoints are blocked", () => {
  const policy = enterprisePolicy();
  for (const model of [
    { ...validModel, provider: "openai" },
    { ...validModel, api: "openai-responses" as const },
    { ...validModel, baseUrl: "https://api.public.invalid/llm/v1" },
    { ...validModel, baseUrl: "https://api.example.invalid/llm/v2" },
    { ...validModel, baseUrl: "https://api.example.invalid/llm/v1-other" },
    { ...validModel, baseUrl: "http://api.example.invalid/llm/v1" },
    { ...validModel, baseUrl: "https://api.example.invalid/llm/v1?debug=1" },
    { ...validModel, baseUrl: "https://api.example.invalid.evil.invalid/llm/v1" },
  ]) {
    assert.equal(isEnterpriseModelAllowed(policy, model), false);
    assert.throws(() => assertEnterpriseModelAllowed(policy, model));
  }
});
