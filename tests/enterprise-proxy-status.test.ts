import assert from "node:assert/strict";
import { test } from "node:test";

import { startProxyPolling } from "../src/taskpane/proxy-status.ts";

void test("enterprise mode never probes the optional local helper proxy", async () => {
  let reads = 0;
  const stop = startProxyPolling({
    get: () => {
      reads += 1;
      return Promise.resolve("https://localhost:3003");
    },
  }, { enabled: false });

  // Even the first asynchronous check must not be scheduled.
  await Promise.resolve();
  stop();
  assert.equal(reads, 0);
});
