#!/usr/bin/env node

/**
 * Generate a dedicated, locally served enterprise test manifest.
 *
 * The add-in ID and branding differ from upstream so test installations do not
 * replace the public Pi for Excel add-in. The generated file is gitignored.
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export const ENTERPRISE_LOCAL_ORIGIN = "https://127.0.0.1:3141";
export const ENTERPRISE_LOCAL_ADDIN_ID = "cbcbad6b-453a-5885-8fb1-9dc3cb3ce11d";
const DEV_ORIGIN = "https://localhost:3141";

export function renderEnterpriseLocalManifest(template) {
  if (typeof template !== "string"
      || !template.includes(DEV_ORIGIN)
      || !/<Id>[0-9a-f-]{36}<\/Id>/iu.test(template)) {
    throw new Error("The source manifest must have a dev localhost origin and an Office add-in ID.");
  }
  const manifest = template
    .replace(/<Id>[0-9a-f-]{36}<\/Id>/iu, `<Id>${ENTERPRISE_LOCAL_ADDIN_ID}</Id>`)
    .replaceAll(DEV_ORIGIN, ENTERPRISE_LOCAL_ORIGIN)
    .replace('<ProviderName>Pi for Excel</ProviderName>', '<ProviderName>Pi for Excel (Enterprise Local)</ProviderName>')
    .replace('DisplayName DefaultValue="Pi for Excel"', 'DisplayName DefaultValue="Pi for Excel (Enterprise Local)"')
    .replace('id="GroupLabel" DefaultValue="Pi"', 'id="GroupLabel" DefaultValue="Pi Enterprise"')
    .replace('id="ButtonLabel" DefaultValue="Open Pi"', 'id="ButtonLabel" DefaultValue="Open Pi Enterprise"')
    .replace('id="ButtonDesc" DefaultValue="Open the Pi AI assistant sidebar"',
      'id="ButtonDesc" DefaultValue="Open the local enterprise Pi for Excel test sidebar"')
    .replace('https://github.com/tmustier/pi-for-excel/issues/new/choose',
      'https://github.com/julmx/pi-for-excel');
  if (manifest.includes(DEV_ORIGIN)) {
    throw new Error("The generated manifest still includes the non-enterprise dev origin.");
  }
  return manifest;
}

const scriptPath = fileURLToPath(import.meta.url);
if (process.argv[1] && path.resolve(process.argv[1]) === scriptPath) {
  const repoRoot = path.resolve(path.dirname(scriptPath), "..");
  const source = path.join(repoRoot, "manifest.xml");
  const output = path.join(repoRoot, "manifest.enterprise.local.xml");
  const manifest = renderEnterpriseLocalManifest(fs.readFileSync(source, "utf8"));
  fs.writeFileSync(output, manifest, { encoding: "utf8", flag: "w" });
  console.log(`Generated ${output}`);
  console.log(`Enterprise taskpane: ${ENTERPRISE_LOCAL_ORIGIN}/src/taskpane.html`);
}
