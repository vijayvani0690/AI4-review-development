param()

$ErrorActionPreference = "Stop"
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $scriptDirectory

$node = Get-Command node.exe -ErrorAction SilentlyContinue
$npm = Get-Command npm.cmd -ErrorAction SilentlyContinue
if (-not $node -or -not $npm) {
    throw "Node.js 20 or newer is required. Install the current Node.js LTS release from https://nodejs.org/ and rerun setup.ps1."
}

$majorVersion = [int]((& $node.Source --version).TrimStart("v").Split(".")[0])
if ($majorVersion -lt 20) {
    throw "Node.js 20 or newer is required. Installed version: $(& $node.Source --version)"
}

$previousSkipDownload = $env:PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD
$env:PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD = "1"
try {
    if (Test-Path -LiteralPath (Join-Path $scriptDirectory "package-lock.json")) {
        & $npm.Source ci
    } else {
        & $npm.Source install
    }
    if ($LASTEXITCODE -ne 0) {
        throw "npm install failed with exit code $LASTEXITCODE."
    }
} finally {
    $env:PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD = $previousSkipDownload
}

Write-Host ""
Write-Host "Standalone setup complete."
Write-Host "Microsoft Edge will be used as the browser; no Playwright browser download was required."
Write-Host "Next: copy the input workbook into .\input and run .\run_report.ps1 -LoginSetup"
