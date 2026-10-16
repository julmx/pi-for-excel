# Dependency security maintenance

## October 2026 remediation

The initial fork's CI reported 23 advisories (including three critical
advisories) in `npm audit --audit-level=high`. The dependency resolution was
out of date, and the root `overrides` section forced an affected Axios
version (`1.19.0`).

Changes:

1. Removed the stale Axios override.
2. Rebuilt `package-lock.json` using `npm install --package-lock-only`
   and non-breaking `npm audit fix` (without `--force`).
3. Removed the optional development-only `office-addin-debugging` package and
   its `npm run sideload` command. Its transitive
   `office-addin-dev-certs → mkcert → node-forge` and
   `office-addin-dev-settings → @microsoft/m365agentstoolkit-cli` chains
   remained vulnerable and npm's suggested forced downgrade would break
   tooling. The built-in documented Office sideload methods and the
   `scripts/start-enterprise-local.ps1` launcher remain available.
4. Kept the explicit toolchain and Pi dependencies at their existing declared
   major versions and did not bypass `npm audit`.

On the remediation branch, GitHub Actions' resulting npm reports were:

| Audit | Total | Moderate | High | Critical |
|---|---:|---:|---:|---:|
| All dependencies | 0 | 0 | 0 | 0 |
| Runtime dependencies only (`--omit=dev`) | 0 | 0 | 0 | 0 |

A zero-vulnerability npm audit means *no known advisories found by the
configured registry at that moment*. It does **not** certify the code,
network configuration or enterprise data-handling behavior.

## CI policy

This fork's GitHub `Dependency Review` workflow could not run because the
repository does not currently have the GitHub Dependency Graph capability
available. Until the feature is enabled, the `Dependency Audit` workflow
runs a full npm advisory audit (including development dependencies) on
pull requests and pushes to `main`, plus a separate runtime-only audit.
The regular CI also runs `npm audit --audit-level=high`.

After enabling the Dependency Graph in GitHub repository settings, the
differential `actions/dependency-review-action` check can be reinstated
alongside the npm audit. No vulnerability threshold has been suppressed.

## Recheck locally

```bash
npm ci
npm audit --audit-level=moderate
npm audit --omit=dev --audit-level=moderate
npm run check
npm test
npm run build
```

If a new alert appears, identify its direct parent and runtime impact,
prefer a patched compatible version, and regenerate the lockfile. Do not
use `npm audit fix --force` without reviewing the potential breaking changes.
