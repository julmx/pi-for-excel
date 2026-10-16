import assert from "node:assert/strict";
import { test } from "node:test";

import { InMemoryModelsStore } from "@earendil-works/pi-ai";

import { ConnectionManager } from "../src/connections/manager.ts";
import { ExtensionRuntimeManager } from "../src/extensions/runtime-manager.ts";
import { EXTENSIONS_REGISTRY_STORAGE_KEY } from "../src/extensions/store.ts";
import { BrowserModelRuntime } from "../src/models/browser-model-runtime.ts";
import {
  getStoredExecutionMode,
  setStoredExecutionMode,
  toggleExecutionMode,
} from "../src/execution/mode.ts";

class MemorySettings {
  private readonly values = new Map<string, unknown>();
  get(key: string): Promise<unknown> {
    return Promise.resolve(this.values.get(key) ?? null);
  }
  set(key: string, value: unknown): Promise<void> {
    this.values.set(key, value);
    return Promise.resolve();
  }
}

void test("enterprise extension runtime refuses installation and persisted activation", async () => {
  const settings = new MemorySettings();
  await settings.set(EXTENSIONS_REGISTRY_STORAGE_KEY, { version: 2, items: [] });

  const connectionManager = new ConnectionManager({ settings });
  const modelRuntime = new BrowserModelRuntime({
    providerKeys: {
      get: () => Promise.resolve(null),
      set: () => Promise.resolve(),
      delete: () => Promise.resolve(),
      list: () => Promise.resolve([]),
    },
    modelCatalogs: new InMemoryModelsStore(),
    getProxyUrl: () => Promise.resolve(undefined),
  });

  let loadedSources = 0;
  const manager = new ExtensionRuntimeManager({
    settings,
    connectionManager,
    modelRuntime,
    getActiveAgent: () => null,
    refreshRuntimeTools: () => Promise.resolve(),
    refreshRuntimeModels: () => Promise.resolve(),
    reservedToolNames: new Set(),
    disabled: true,
    loadExtensionFromSource: () => {
      loadedSources += 1;
      return Promise.reject(new Error("should not load"));
    },
  });

  await manager.initialize();
  assert.deepEqual(manager.list(), []);
  assert.deepEqual(manager.getRegisteredTools(), []);
  assert.equal(loadedSources, 0);

  await assert.rejects(
    manager.installFromCode("Untrusted", "console.log('untrusted')"),
    /disabled by enterprise policy/u,
  );
  await assert.rejects(
    manager.installFromModuleSpecifier("Untrusted", "https://external.invalid/extension.js"),
    /disabled by enterprise policy/u,
  );
  await assert.rejects(
    manager.reloadExtension("any"),
    /disabled by enterprise policy/u,
  );
  await assert.rejects(
    manager.setExtensionEnabled("any", true),
    /disabled by enterprise policy/u,
  );
  assert.equal(loadedSources, 0);
});

void test("workbook mutation confirmation remains optional and user-selectable", async () => {
  const settings = new MemorySettings();
  assert.equal(await getStoredExecutionMode(settings), "yolo");

  await setStoredExecutionMode(settings, "safe");
  assert.equal(await getStoredExecutionMode(settings), "safe");
  await setStoredExecutionMode(settings, toggleExecutionMode("safe"));
  assert.equal(await getStoredExecutionMode(settings), "yolo");
});
