import assert from "node:assert/strict";
import { test } from "node:test";
import { resolveEnterpriseRelayTarget } from "../scripts/enterprise-api-relay.mjs";

const base = "https://company.example.invalid/llm/v1";

void test("relay pins only models GET and chat completions POST under the configured API path", () => {
  const models = resolveEnterpriseRelayTarget(base, "/__enterprise_api/models", "GET");
  assert.equal(models?.href, "https://company.example.invalid/llm/v1/models");
  const chat = resolveEnterpriseRelayTarget(base + "/", "/__enterprise_api/chat/completions", "POST");
  assert.equal(chat?.href, "https://company.example.invalid/llm/v1/chat/completions");
});

void test("relay cannot be used to select arbitrary destinations or routes", () => {
  for (const path of [
    "/__enterprise_api/http://attacker.invalid",
    "/__enterprise_api/chat/completions/../models",
    "/__enterprise_api/../models",
    "/__enterprise_api/models?api_key=anything",
    "/__enterprise_api/secret",
    "/__enterprise_api/chat/completions/",
    "/other",
  ]) {
    assert.equal(resolveEnterpriseRelayTarget(base, path, "GET"), null, path);
  }
  assert.equal(resolveEnterpriseRelayTarget(base, "/__enterprise_api/chat/completions", "GET"), null);
  assert.equal(resolveEnterpriseRelayTarget(base, "/__enterprise_api/models", "POST"), null);
});

void test("unusual approved API paths cannot rewrite the upstream origin", () => {
  const target = resolveEnterpriseRelayTarget(
    "https://company.example.invalid//other-domain",
    "/__enterprise_api/models",
    "GET",
  );
  assert.equal(target?.origin, "https://company.example.invalid");
});

void test("relay requires HTTPS upstream without embedded secrets", () => {
  for (const invalid of [
    "http://company.example.invalid/v1",
    "https://user:password@company.example.invalid/v1",
    "https://company.example.invalid/v1?token=secret",
  ]) {
    assert.throws(() => resolveEnterpriseRelayTarget(invalid, "/__enterprise_api/models", "GET"));
  }
});
