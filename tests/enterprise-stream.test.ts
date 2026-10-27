import assert from "node:assert/strict";
import { register } from "node:module";
import { test } from "node:test";
import type { CustomProvider } from "../src/storage/local/custom-providers-store.ts";

// Node isolates test files; these build flags cannot affect ordinary-build contracts.
register(new URL("./fixtures/enterprise-build-env.mjs", import.meta.url), import.meta.url);

const { InMemoryModelsStore } = await import("@earendil-works/pi-ai");
const { BrowserModelRuntime } = await import("../src/models/browser-model-runtime.ts");
const { saveOpenAiGatewayConfig } = await import("../src/auth/custom-gateways.ts");
const { createOfficeStreamFn } = await import("../src/auth/stream-proxy.ts");
const { providerRetryDelay } = await import("../src/auth/provider-failure.ts");

void test("enterprise streaming preserves the gateway's Retry-After for queue-owned retries", async (t) => {
  const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  Object.defineProperty(globalThis, "document", { configurable: true, value: new EventTarget() });
  t.after(() => {
    if (originalDocument) Object.defineProperty(globalThis, "document", originalDocument);
    else Reflect.deleteProperty(globalThis, "document");
  });

  const providers = new Map<string, CustomProvider>();
  const gateway = await saveOpenAiGatewayConfig({
    get: (id) => Promise.resolve(providers.get(id) ?? null),
    set: (provider) => { providers.set(provider.id, provider); return Promise.resolve(); },
    delete: (id) => { providers.delete(id); return Promise.resolve(); },
    getAll: () => Promise.resolve([...providers.values()]),
  }, {
    displayName: "Approved",
    endpointUrl: "https://approved.example.invalid/v1",
    modelId: "test-model",
    apiKey: "synthetic-test-key",
    allowDiscoveryOnly: true,
  });
  const runtime = new BrowserModelRuntime({
    providerKeys: {
      get: () => Promise.resolve(null),
      set: () => Promise.resolve(),
      delete: () => Promise.resolve(),
      list: () => Promise.resolve([]),
    },
    modelCatalogs: new InMemoryModelsStore(),
    getProxyUrl: () => Promise.resolve(undefined),
  });
  await runtime.syncCustomProviders([...providers.values()]);
  const model = runtime.models.getModel(gateway.providerName, gateway.modelId);
  assert.ok(model);

  let requests = 0;
  const stream = await createOfficeStreamFn(
    () => { throw new Error("Enterprise must not look up an external proxy"); },
    runtime.models,
  )(model, { messages: [{ role: "user", content: "Synthetic request", timestamp: 0 }] }, {
    // Bound the unfixed SDK path so the regression fails promptly on the delay assertion.
    maxRetries: 0,
    fetch: (input) => {
      requests += 1;
      const url = input instanceof Request ? input.url : input.toString();
      assert.equal(url, "https://approved.example.invalid/v1/chat/completions");
      return Promise.resolve(new Response(JSON.stringify({ error: { code: "rate_limit_exceeded", message: "Slow down" } }), {
        status: 429,
        headers: { "Content-Type": "application/json", "Retry-After": "7" },
      }));
    },
  });
  const result = await stream.result();
  assert.equal(result.stopReason, "error");
  assert.equal(requests, 1);
  assert.equal(providerRetryDelay(result, 2_000), 7_000);
});
