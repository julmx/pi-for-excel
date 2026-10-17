/**
 * Build-configured private endpoint policy for opt-in enterprise deployments.
 *
 * Guards agent model calls. It does not replace network/CSP egress controls.
 * Never put credentials in VITE_* settings (they are embedded in client JS).
 */
import type { Api, Model } from "@earendil-works/pi-ai";
import { OPENAI_GATEWAY_PROVIDER_PREFIX } from "./custom-gateways.js";

export interface EnterpriseModelPolicy {
  baseUrl: string;
}

export interface EnterpriseModelPolicyConfig {
  enabled: boolean;
  baseUrl?: string;
}

function normalizeApprovedBaseUrl(raw: string): string {
  const value = raw.trim();
  if (!value) throw new Error("Enterprise API base URL is required.");

  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error("Enterprise API base URL must be a valid HTTPS URL.");
  }

  if (
    url.protocol !== "https:"
    || !url.hostname
    || url.username
    || url.password
    || url.search
    || url.hash
  ) {
    throw new Error("Enterprise API base URL must be HTTPS without credentials, query or fragment.");
  }

  // Forbid a non-canonical user-supplied URL so a copy/paste error cannot
  // unexpectedly change the destination or path.
  return url.href.endsWith("/") ? url.href.slice(0, -1) : url.href;
}

/** Throw when a restricted build has no complete approved endpoint. */
export function parseEnterpriseModelPolicy(
  config: EnterpriseModelPolicyConfig,
): EnterpriseModelPolicy | null {
  if (!config.enabled) return null;

  const baseUrl = normalizeApprovedBaseUrl(config.baseUrl ?? "");
  return { baseUrl };
}

/** The build deliberately contains only endpoint URL, never API keys. */
export function readBuildEnterpriseModelPolicy(): EnterpriseModelPolicy | null {
  const env = import.meta.env;
  return parseEnterpriseModelPolicy({
    enabled: env?.VITE_PI_ENTERPRISE_MODE === "true",
    baseUrl: env?.VITE_PI_ENTERPRISE_API_BASE_URL ?? "",
  });
}

export function isEnterpriseModelAllowed(
  policy: EnterpriseModelPolicy,
  model: Pick<Model<Api>, "provider" | "api" | "baseUrl" | "id">,
): boolean {
  if (model.api !== "openai-completions") return false;
  if (!model.provider.startsWith(OPENAI_GATEWAY_PROVIDER_PREFIX)) return false;
  if (!model.baseUrl) return false;

  try {
    // Check the complete API base path; comparing hostnames is insufficient.
    return normalizeApprovedBaseUrl(model.baseUrl) === policy.baseUrl;
  } catch {
    return false;
  }
}

/**
 * Pick only a model registered through the approved private gateway.
 * The enterprise-bootstrap placeholder must never be chosen when a genuine
 * gateway model is available.
 */
export function findApprovedEnterpriseModel<T extends Pick<Model<Api>, "provider" | "api" | "baseUrl" | "id">>(
  policy: EnterpriseModelPolicy,
  models: readonly T[],
): T | null {
  return models.find((model) => isEnterpriseModelAllowed(policy, model)) ?? null;
}

/** No fallback to other vendors, custom gateways, or cached model choices. */
export function assertEnterpriseModelAllowed(
  policy: EnterpriseModelPolicy,
  model: Pick<Model<Api>, "provider" | "api" | "baseUrl" | "id">,
): void {
  if (!isEnterpriseModelAllowed(policy, model)) {
    if (model.provider === "enterprise-bootstrap") {
      throw new Error(
        "No enterprise gateway model is selected. Configure the approved Custom OpenAI-compatible gateway in Settings, then select its model.",
      );
    }
    throw new Error(
      "Enterprise model configuration mismatch. Select a model from the approved Custom OpenAI-compatible gateway and verify that its API base URL exactly matches the URL used to build this add-in.",
    );
  }
}
