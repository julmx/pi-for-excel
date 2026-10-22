import assert from "node:assert/strict";
import { test } from "node:test";

import type { ModelsStore, ModelsStoreEntry } from "@earendil-works/pi-ai";
import {
  parseEnterpriseModelPolicy,
} from "../src/auth/enterprise-model-policy.ts";
import { isEnterpriseFetchUrlAllowed } from "../src/auth/enterprise-fetch-policy.ts";
import { saveOpenAiGatewayConfig, type CustomProvidersStoreLike } from "../src/auth/custom-gateways.ts";
import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";
import { getThinkingLevelsForModel } from "../src/models/thinking-levels.ts";
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
  assert.equal(runtime.models.getModel("openai", "gpt-5.6-sol"), undefined);
  assert.ok(runtime.models.getModel("enterprise-bootstrap", "configure-gateway"));
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


void test("enterprise discovers every gateway model without configuring a seed model ID", async () => {
  const providers = new MemoryProviders();
  const gateway = await saveOpenAiGatewayConfig(providers, {
    displayName: "Discovery only",
    endpointUrl: policy.baseUrl,
    apiKey: "fake-test-token",
    contextWindow: 131_072,
    maxTokens: 16_384,
    allowDiscoveryOnly: true,
  });
  assert.equal(gateway.modelId, "");
  assert.deepEqual((await providers.get(gateway.id))?.models, []);

  const urls: string[] = [];
  const runtime = new BrowserModelRuntime({
    enterprisePolicy: policy,
    providerKeys: {
      get: () => Promise.resolve(null),
      set: () => Promise.resolve(),
      delete: () => Promise.resolve(),
      list: () => Promise.resolve([]),
    },
    modelCatalogs: new MemoryCatalogs(),
    getProxyUrl: () => Promise.resolve("https://should-not-use.example.invalid"),
    fetchFn: (input, init) => {
      const url = typeof input === "string"
        ? input : input instanceof URL ? input.toString() : input.url;
      urls.push(url);
      assert.equal(new Headers(init?.headers).get("Authorization"), "Bearer fake-test-token");
      return Promise.resolve(Response.json({
        data: [{ id: "model-B" }, { id: "model-A" }, { id: "model-B" }],
      }));
    },
  });
  await runtime.syncCustomProviders(await providers.getAll());
  assert.ok(runtime.models.getProvider(gateway.providerName));
  const result = await runtime.refresh({ allowNetwork: true, force: true });
  assert.equal(result.errors.size, 0);
  assert.deepEqual(urls, [`${policy.baseUrl}/models`]);
  const discovered = runtime.models.getModels(gateway.providerName);
  assert.deepEqual(discovered.map((model) => model.id), ["model-A", "model-B"]);
  for (const model of discovered) {
    assert.equal(model.contextWindow, 131_072);
    assert.equal(model.maxTokens, 16_384);
  }
});

void test("enterprise never registers a discovery-only gateway pointing outside the approved API", async () => {
  const providers = new MemoryProviders();
  await saveOpenAiGatewayConfig(providers, {
    displayName: "Unapproved discovery",
    endpointUrl: "https://unapproved.example.invalid/v1",
    allowDiscoveryOnly: true,
  });
  const runtime = new BrowserModelRuntime({
    enterprisePolicy: policy,
    providerKeys: {
      get: () => Promise.resolve(null),
      set: () => Promise.resolve(),
      delete: () => Promise.resolve(),
      list: () => Promise.resolve([]),
    },
    modelCatalogs: new MemoryCatalogs(),
    getProxyUrl: () => Promise.resolve(undefined),
    fetchFn: () => Promise.reject(new Error("Must not fetch unapproved endpoint")),
  });
  await runtime.syncCustomProviders(await providers.getAll());
  assert.deepEqual(runtime.models.getProviders().filter((provider) =>
    provider.id.startsWith("Gateway · ")), []);
});


void test("each approved enterprise gateway discovers and exposes its own models", async () => {
  const multiple = parseEnterpriseModelPolicy({
    enabled: true,
    baseUrls: [
      "https://private-a.example.invalid/v1",
      "https://private-b.example.invalid/api",
    ],
  });
  assert.ok(multiple);
  const providers = new MemoryProviders();
  const firstUrl = multiple.baseUrls[0];
  const secondUrl = multiple.baseUrls[1];
  assert.ok(firstUrl);
  assert.ok(secondUrl);
  const first = await saveOpenAiGatewayConfig(providers, {
    displayName: "Quality check", endpointUrl: firstUrl,
    apiKey: "fake-key-a", allowDiscoveryOnly: true,
  });
  const second = await saveOpenAiGatewayConfig(providers, {
    displayName: "GenX", endpointUrl: secondUrl,
    apiKey: "fake-key-b", allowDiscoveryOnly: true,
  });
  const unauthorized = await saveOpenAiGatewayConfig(providers, {
    displayName: "Unauthorized", endpointUrl: "https://unknown.example.invalid/api",
    allowDiscoveryOnly: true,
  });
  const requests: string[] = [];
  const runtime = new BrowserModelRuntime({
    enterprisePolicy: multiple,
    providerKeys: {
      get: () => Promise.resolve(null), set: () => Promise.resolve(),
      delete: () => Promise.resolve(), list: () => Promise.resolve([]),
    },
    modelCatalogs: new MemoryCatalogs(),
    getProxyUrl: () => Promise.resolve("https://unapproved-proxy.example.invalid"),
    fetchFn: (input, init) => {
      const url = typeof input === "string"
        ? input : input instanceof URL ? input.toString() : input.url;
      requests.push(url);
      const key = new Headers(init?.headers).get("Authorization");
      if (url === "https://private-a.example.invalid/v1/models") {
        assert.equal(key, "Bearer fake-key-a");
        return Promise.resolve(Response.json({ data: [{ id: "quality-model" }] }));
      }
      assert.equal(url, "https://private-b.example.invalid/api/models");
      assert.equal(key, "Bearer fake-key-b");
      return Promise.resolve(Response.json({ data: [{ id: "genx-model" }] }));
    },
  });
  await runtime.syncCustomProviders(await providers.getAll());
  const result = await runtime.refresh({ allowNetwork: true, force: true });
  assert.equal(result.errors.size, 0);
  assert.deepEqual(requests.sort(), [
    "https://private-a.example.invalid/v1/models",
    "https://private-b.example.invalid/api/models",
  ]);
  assert.deepEqual(runtime.models.getModels(first.providerName).map((m) => m.id), ["quality-model"]);
  assert.deepEqual(runtime.models.getModels(second.providerName).map((m) => m.id), ["genx-model"]);
  assert.equal(runtime.models.getProvider(unauthorized.providerName), undefined);
});


void test("adjustable reasoning is opt-in per discovered gateway model", async () => {
  const providers = new MemoryProviders();
  const gateway = await saveOpenAiGatewayConfig(providers, {
    endpointUrl: policy.baseUrl,
    allowDiscoveryOnly: true,
    reasoningModelIds: ["reasoning-a"],
  });
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
    getProxyUrl: () => Promise.resolve(undefined),
    fetchFn: () => Promise.resolve(Response.json({
      data: [{ id: "ordinary-b" }, { id: "reasoning-a" }],
    })),
  });
  await runtime.syncCustomProviders(await providers.getAll());
  const result = await runtime.refresh({ allowNetwork: true, force: true });
  assert.equal(result.errors.size, 0);
  assert.equal(runtime.models.getModel(gateway.providerName, "reasoning-a")?.reasoning, true);
  assert.equal(runtime.models.getModel(gateway.providerName, "ordinary-b")?.reasoning, false);
  assert.ok(getThinkingLevelsForModel(runtime.models.getModel(gateway.providerName, "reasoning-a") ?? null).length > 1);
  assert.deepEqual(getThinkingLevelsForModel(runtime.models.getModel(gateway.providerName, "ordinary-b") ?? null), ["off"]);

  const saved = catalogs.entries.get(gateway.providerName);
  assert.equal(saved?.models.find((m) => m.id === "reasoning-a")?.reasoning, true);
  assert.equal(saved?.models.find((m) => m.id === "ordinary-b")?.reasoning, false);
});

void test("cached discovery models follow current reasoning opt-in after an edited gateway", async () => {
  const providers = new MemoryProviders();
  const gateway = await saveOpenAiGatewayConfig(providers, {
    endpointUrl: policy.baseUrl,
    allowDiscoveryOnly: true,
    reasoningModelIds: ["reasoning-a"],
  });
  const catalogs = new MemoryCatalogs();
  await catalogs.write(gateway.providerName, {
    models: [
      { id: "reasoning-a", reasoning: false, provider: gateway.providerName,
        api: "openai-completions", baseUrl: policy.baseUrl, contextWindow: 8000,
        maxTokens: 2000, name: "reasoning-a", input: ["text"],
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 } },
      { id: "ordinary-b", reasoning: true, provider: gateway.providerName,
        api: "openai-completions", baseUrl: policy.baseUrl, contextWindow: 8000,
        maxTokens: 2000, name: "ordinary-b", input: ["text"],
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 } },
    ],
  });
  const runtime = new BrowserModelRuntime({
    enterprisePolicy: policy,
    providerKeys: {
      get: () => Promise.resolve(null),
      set: () => Promise.resolve(),
      delete: () => Promise.resolve(),
      list: () => Promise.resolve([]),
    },
    modelCatalogs: catalogs,
    getProxyUrl: () => Promise.resolve(undefined),
    fetchFn: () => Promise.reject(new Error("No network necessary for cached test")),
  });
  await runtime.syncCustomProviders(await providers.getAll());
  const available = await runtime.models.getAvailable();
  assert.equal(available.find((m) => m.id === "reasoning-a")?.reasoning, true);
  assert.equal(available.find((m) => m.id === "ordinary-b")?.reasoning, false);
});
