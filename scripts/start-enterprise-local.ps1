# Launch a restricted production build for a local Excel Windows smoke test.
# No real workbook data or API credentials should be used during initial testing.
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ApiBaseUrl,

    [switch]$SkipInstall,

    # Explicit opt-in if Office WebView2 cannot call the corporate API via CORS.
    [switch]$UseLocalApiRelay
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

    # Reuse the existing local certificate only if it covers both the
    # taskpane IP and the runner hostname; no certificate regeneration needed.
    # PowerShell passes single quotes inside a double-quoted argument intact
    # to the Windows Node executable (unlike double quotes inside single quotes).
    & node -e "const fs=require('node:fs'); const {X509Certificate}=require('node:crypto'); const cert=new X509Certificate(fs.readFileSync('cert.pem')); if(cert.checkIP('127.0.0.1') !== '127.0.0.1' || cert.checkHost('localhost') !== 'localhost') process.exit(1)"
    if ($LASTEXITCODE -ne 0) {
        throw "Local cert.pem must cover 127.0.0.1 and localhost. Use an approved certificate if either hostname is missing."
    }

    if (-not $SkipInstall) {
        Invoke-Checked -Executable "npm.cmd" -Arguments @("ci")
    } elseif (-not (Test-Path (Join-Path $root "node_modules"))) {
        throw "node_modules is missing. Rerun without -SkipInstall."
    }

    # Endpoint only, not a secret. NEVER place the corporate API key in VITE_*.
    $env:VITE_PI_ENTERPRISE_MODE = "true"
    $env:VITE_PI_ENTERPRISE_API_BASE_URL = $ApiBaseUrl.TrimEnd("/")
    $env:VITE_PI_ENTERPRISE_LOCAL_API_RELAY = if ($UseLocalApiRelay) { "true" } else { "false" }
    # Override inherited developer/CI values: the local origins are fixed.
    $env:VITE_PI_ENTERPRISE_TASKPANE_ORIGIN = "https://127.0.0.1:3141"
    $env:VITE_PI_ENTERPRISE_RUNNER_ORIGIN = "https://localhost:3142"
    Invoke-Checked -Executable "npm.cmd" -Arguments @("run", "build")

    Invoke-Checked -Executable "node" -Arguments @("scripts/generate-enterprise-local-manifest.mjs")

    # These must match the build-time URL; the HTTPS server applies the CSP.
    $env:ENTERPRISE_MODE = "1"
    $env:ENTERPRISE_API_BASE_URL = $env:VITE_PI_ENTERPRISE_API_BASE_URL
    $env:ENTERPRISE_LOCAL_API_RELAY = if ($UseLocalApiRelay) { "1" } else { "0" }
    $env:PORT = "3141"

    Write-Host ""
    Write-Host "Local enterprise Pi for Excel is ready for a synthetic-data test."
    Write-Host "Manifest: $root\manifest.enterprise.local.xml"
    Write-Host "Taskpane: https://127.0.0.1:3141/src/taskpane.html"
    Write-Host "Isolated Python runner: https://localhost:3142/enterprise-runner.html"
    Write-Host "Approved API: $($env:ENTERPRISE_API_BASE_URL)"
    Write-Host "Local API relay enabled: $($UseLocalApiRelay.IsPresent)"
    Write-Host "Keep this terminal open and load the manifest via Excel > Add-ins."
    Write-Host "Press Ctrl+C here to stop the server."
    Write-Host ""
    Invoke-Checked -Executable "npm.cmd" -Arguments @("run", "serve:dist")
}
finally {
    Pop-Location
}
