import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import {
  ENTERPRISE_LOCAL_ADDIN_ID,
  ENTERPRISE_LOCAL_ORIGIN,
  renderEnterpriseLocalManifest,
} from "../scripts/generate-enterprise-local-manifest.mjs";

const upstreamManifest = readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");

void test("enterprise local manifest has a stable, separate add-in identity", () => {
  const result = renderEnterpriseLocalManifest(upstreamManifest);
  assert.ok(result.includes(`<Id>${ENTERPRISE_LOCAL_ADDIN_ID}</Id>`));
  assert.notEqual(
    result.match(/<Id>([^<]+)<\/Id>/u)?.[1],
    upstreamManifest.match(/<Id>([^<]+)<\/Id>/u)?.[1],
  );
  assert.match(result, /DisplayName DefaultValue="Pi for Excel \(Enterprise Local\)"/u);
  assert.match(result, /Open Pi Enterprise/u);
  assert.equal(renderEnterpriseLocalManifest(upstreamManifest), result);
});

void test("enterprise URLs remain loopback-only throughout the manifest", () => {
  const result = renderEnterpriseLocalManifest(upstreamManifest);
  assert.ok(result.includes(`${ENTERPRISE_LOCAL_ORIGIN}/src/taskpane.html`));
  assert.equal(result.includes("https://localhost:3141"), false);
  assert.equal(result.includes("https://pi-for-excel.vercel.app"), false);
  assert.equal(result.includes("https://api.example.invalid"), false);

  const servedUrls = [...result.matchAll(/(?:SourceLocation|IconUrl|HighResolutionIconUrl|bt:(?:Image|Url))[^>]*DefaultValue="(https:[^"]+)"/gu)]
    .map((match) => match[1]);
  assert.ok(servedUrls.length >= 4);
  for (const url of servedUrls) {
    assert.equal(new URL(url).origin, ENTERPRISE_LOCAL_ORIGIN);
  }
});

void test("manifest generator rejects unrelated source templates", () => {
  assert.throws(() => renderEnterpriseLocalManifest("unrelated XML"));
});
