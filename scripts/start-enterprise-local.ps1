# Launch a restricted production build for a local Excel Windows smoke test.
# No real workbook data or API credentials should be used during initial testing.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ApiBaseUrl,

    [switch]$SkipInstall
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Invoke-Checked {
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )
    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Executable failed with exit code $LASTEXITCODE"
    }
}

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
Push-Location $root
try {
    $parsed = $null
    $valid = [Uri]::TryCreate($ApiBaseUrl, [UriKind]::Absolute, [ref]$parsed)
    if (-not $valid -or $parsed.Scheme -ne "https" -or -not $parsed.Host -or
        $parsed.UserInfo -or $parsed.Query -or $parsed.Fragment) {
        throw "ApiBaseUrl must be an HTTPS API base URL without credentials, a query, or a fragment."
    }

    foreach ($executable in @("node", "npm.cmd")) {
        if (-not (Get-Command $executable -ErrorAction SilentlyContinue)) {
            throw "$executable was not found. Install Node.js 22 or later and reopen PowerShell."
        }
    }

    $cert = Join-Path $root "cert.pem"
    $key = Join-Path $root "key.pem"
    if (-not (Test-Path $cert) -or -not (Test-Path $key)) {
        throw @"
Local HTTPS certificates are missing. Install mkcert, then run these commands
from the repository root (the CA installation may require IT approval):

  mkcert -install
  mkcert -cert-file cert.pem -key-file key.pem localhost 127.0.0.1

Certificates stay on your machine and are ignored by Git.
"@
    }

    if (-not $SkipInstall) {
        Invoke-Checked -Executable "npm.cmd" -Arguments @("ci")
    } elseif (-not (Test-Path (Join-Path $root "node_modules"))) {
        throw "node_modules is missing. Rerun without -SkipInstall."
    }

    # Endpoint only, not a secret. NEVER place the corporate API key in VITE_*.
    $env:VITE_PI_ENTERPRISE_MODE = "true"
    $env:VITE_PI_ENTERPRISE_API_BASE_URL = $ApiBaseUrl.TrimEnd("/")
    Invoke-Checked -Executable "npm.cmd" -Arguments @("run", "build")

    Invoke-Checked -Executable "node" -Arguments @("scripts/generate-enterprise-local-manifest.mjs")

    # These must match the build-time URL; the HTTPS server applies the CSP.
    $env:ENTERPRISE_MODE = "1"
    $env:ENTERPRISE_API_BASE_URL = $env:VITE_PI_ENTERPRISE_API_BASE_URL
    $env:PORT = "3141"

    Write-Host ""
    Write-Host "Local enterprise Pi for Excel is ready for a synthetic-data test."
    Write-Host "Manifest: $root\manifest.enterprise.local.xml"
    Write-Host "Taskpane: https://127.0.0.1:3141/src/taskpane.html"
    Write-Host "Keep this terminal open and load the manifest via Excel > Add-ins."
    Write-Host "Press Ctrl+C here to stop the server."
    Write-Host ""
    Invoke-Checked -Executable "npm.cmd" -Arguments @("run", "serve:dist")
}
finally {
    Pop-Location
}
