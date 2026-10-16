import assert from "node:assert/strict";
import { test } from "node:test";

import type { ModelsStore, ModelsStoreEntry } from "@earendil-works/pi-ai";
import {
  parseEnterpriseModelPolicy,
} from "../src/auth/enterprise-model-policy.ts";
import { isEnterpriseFetchUrlAllowed } from "../src/auth/enterprise-fetch-policy.ts";
import { saveOpenAiGatewayConfig, type CustomProvidersStoreLike } from "../src/auth/custom-gateways.ts";
import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";
import type { CustomProvider } from "../src/storage/local/custom-providers-store.ts";

class MemoryCatalogs implements ModelsStore {
  readonly entries = new Map<string, ModelsStoreEntry>();

  read(provider: string): Promise<ModelsStoreEntry | undefined> {
    return Promise.resolve(this.entries.get(provider));
  }

  write(provider: string, entry: ModelsStoreEntry): Promise<void> {
    this.entries.set(provider, entry);
    return Promise.resolve();
  }

  delete(provider: string): Promise<void> {
    this.entries.delete(provider);
    return Promise.resolve();
  }
}

class MemoryProviders implements CustomProvidersStoreLike {
  private readonly entries = new Map<string, CustomProvider>();
  get(id: string): Promise<CustomProvider | null> {
    return Promise.resolve(this.entries.get(id) ?? null);
  }
  set(provider: CustomProvider): Promise<void> {
    this.entries.set(provider.id, provider);
    return Promise.resolve();
  }
  delete(id: string): Promise<void> {
    this.entries.delete(id);
    return Promise.resolve();
  }
  getAll(): Promise<CustomProvider[]> {
    return Promise.resolve([...this.entries.values()]);
  }
}

const policy = parseEnterpriseModelPolicy({
  enabled: true,
  baseUrl: "https://approved.example.invalid/llm/v1",
});
assert.ok(policy);

void test("enterprise fetch policy allows local assets and approved API paths only", () => {
  const origin = "https://excel.internal.invalid";
  for (const url of [
    "/assets/app.js",
    "https://excel.internal.invalid/src/taskpane.html",
    "https://approved.example.invalid/llm/v1/models",
    "https://approved.example.invalid/llm/v1/chat/completions",
  ]) {
    assert.equal(isEnterpriseFetchUrlAllowed(policy, url, origin), true, url);
  }
  for (const url of [
    "https://api.openai.com/v1/chat/completions",
    "https://approved.example.invalid/llm/v1/other",
    "https://approved.example.invalid/llm/v1/chat/completions?target=external",
    "https://approved.example.invalid.evil.invalid/llm/v1/models",
    "http://approved.example.invalid/llm/v1/models",
    "https://attacker.invalid/?copy=data",
    "data:text/plain,sensitive",
  ]) {
    assert.equal(isEnterpriseFetchUrlAllowed(policy, url, origin), false, url);
  }
});

void test("enterprise mode discovers approved gateway models without an external proxy", async () => {
  const providers = new MemoryProviders();
  const approved = await saveOpenAiGatewayConfig(providers, {
    displayName: "Approved",
    endpointUrl: policy.baseUrl,
    modelId: "initial-model",
    apiKey: "fake-test-token",
  });
  await saveOpenAiGatewayConfig(providers, {
    displayName: "External",
    endpointUrl: "https://unapproved.example.invalid/v1",
    modelId: "other-model",
    apiKey: "fake-external-token",
  });

  const requested: string[] = [];
  const catalogs = new MemoryCatalogs();
  const runtime = new BrowserModelRuntime({
    enterprisePolicy: policy,
    providerKeys: {
      get: () => Promise.resolve(null),
      set: () => Promise.resolve(),
      delete: () => Promise.resolve(),
      list: () => Promise.resolve([]),
    },
    modelCatalogs: catalogs,
    getProxyUrl: () => Promise.resolve("https://unapproved-proxy.example.invalid"),
    fetchFn: (input, init) => {
      const url = typeof input === "string"
        ? input
        : input instanceof URL ? input.toString() : input.url;
      requested.push(url);
      assert.equal(new Headers(init?.headers).get("Authorization"), "Bearer fake-test-token");
      return Promise.resolve(Response.json({
        data: [{ id: "new-model" }, { id: "initial-model" }, { id: "new-model" }],
      }));
    },
  });
  await runtime.syncCustomProviders(await providers.getAll());
  const result = await runtime.refresh({ allowNetwork: true, force: true });

  assert.equal(result.errors.size, 0);
  assert.deepEqual(requested, ["https://approved.example.invalid/llm/v1/models"]);
  assert.deepEqual(
    runtime.models.getModels(approved.providerName).map((model) => model.id),
    ["initial-model", "new-model"],
  );
  assert.equal(catalogs.entries.get(approved.providerName)?.models.length, 2);
  assert.equal(runtime.models.getProviders().some((item) => item.id === "Gateway · External"), false);
  assert.throws(() => runtime.registerExtensionProvider("extension", {
    id: "extension",
    name: "Extension",
    api: "openai-completions",
    baseUrl: policy.baseUrl,
    models: [{ id: "model" }],
    resolveApiKey: () => Promise.resolve(undefined),
  }), /disabled in enterprise mode/u);
});
