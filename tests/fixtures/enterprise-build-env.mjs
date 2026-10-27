// Supply Vite's build-time environment at the loader boundary in this test process.
const env = {
  VITE_PI_ENTERPRISE_MODE: "true",
  VITE_PI_ENTERPRISE_API_BASE_URL: "https://approved.example.invalid/v1",
  VITE_PI_ENTERPRISE_LOCAL_API_RELAY: "false",
};

export async function load(url, context, next) {
  const result = await next(url, context);
  if (url.includes("/src/") && result.source) {
    return { ...result, source: String(result.source).replaceAll("import.meta.env", JSON.stringify(env)) };
  }
  return result;
}
