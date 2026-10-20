/**
 * Tool registry — creates all built-in tools for the agent.
 *
 * Canonical source of truth for core tools lives in `src/tools/registry.ts`.
 * Experimental/non-core tools are appended here.
 */

import type { SpreadsheetHostKind } from "../host/index.js";
import { createCoreTools, type AnyCoreTool } from "./registry.js";
import { filterEnterpriseBuiltInTools } from "./enterprise-policy.js";
import { selectOfficeCoupledToolForHost } from "./host-selection.js";
import type { SkillReadCache } from "../skills/read-cache.js";
import { createTmuxTool } from "./tmux.js";
import { createPythonRunTool } from "./python-run.js";
import { createLibreOfficeConvertTool } from "./libreoffice-convert.js";
import { createPythonTransformRangeTool } from "./python-transform-range.js";
import { createEnterprisePythonExcelTool } from "./enterprise-python-excel.js";
import { createFilesTool } from "./files.js";
import { createExecuteOfficeJsTool } from "./execute-office-js.js";
import { createExecuteWpsJsTool } from "./execute-wps-js.js";
import {
  createExtensionsManagerTool,
  type ExtensionsManagerToolRuntime,
} from "./extensions-manager.js";

export interface CreateAllToolsOptions {
  hostKind?: SpreadsheetHostKind;
  getExtensionManager?: () => ExtensionsManagerToolRuntime | null;
  getSessionId?: () => string | null;
  skillReadCache?: SkillReadCache;
}

export function createAllTools(options: CreateAllToolsOptions = {}): AnyCoreTool[] {
  const getExtensionManager = options.getExtensionManager ?? (() => null);
  const hostKind = options.hostKind ?? "office";

  const skills = {
    ...(options.getSessionId !== undefined ? { getSessionId: options.getSessionId } : {}),
    ...(options.skillReadCache !== undefined ? { readCache: options.skillReadCache } : {}),
  };

  // Enterprise retains the reviewed core tools and isolated Python. Restore the
  // original Office.js fallback for the agent only; Python RPC stays allowlisted.
  if (import.meta.env?.VITE_PI_ENTERPRISE_MODE === "true") {
    return [
      ...filterEnterpriseBuiltInTools(createCoreTools({ hostKind, skills })),
      ...(hostKind === "office" ? [createEnterprisePythonExcelTool(hostKind), createExecuteOfficeJsTool()] : []),
    ];
  }

  return [
    ...createCoreTools({
      hostKind,
      skills,
    }),
    createTmuxTool(),
    createPythonRunTool(),
    createLibreOfficeConvertTool(),
    selectOfficeCoupledToolForHost(createPythonTransformRangeTool(), hostKind),
    createFilesTool(),
    selectOfficeCoupledToolForHost(createExecuteOfficeJsTool(), hostKind),
    ...(hostKind === "wps" ? [createExecuteWpsJsTool()] : []),
    createExtensionsManagerTool({ getManager: getExtensionManager }),
  ];
}
