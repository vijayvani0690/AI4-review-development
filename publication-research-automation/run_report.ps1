param(
    [ValidateSet("weekly", "monthly")]
    [string]$Period = "weekly",
    [string]$StartDate,
    [string]$EndDate,
    [int]$MaxSites = 0,
    [string]$Site,
    [Alias("Browser")]
    [switch]$BrowserMode,
    [switch]$Headless,
    [switch]$NonInteractive,
    [switch]$LoginSetup,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$node = Get-Command node.exe -ErrorAction SilentlyContinue
if (-not $node) {
    throw "Node.js 20 or newer is required. Install Node.js from https://nodejs.org/ and run setup.ps1."
}
$majorVersion = [int]((& $node.Source --version).TrimStart("v").Split(".")[0])
if ($majorVersion -lt 20) {
    throw "Node.js 20 or newer is required. Installed version: $(& $node.Source --version)"
}
$excelJs = Join-Path $scriptDirectory "node_modules\exceljs"
$playwright = Join-Path $scriptDirectory "node_modules\playwright"
if (-not (Test-Path -LiteralPath $excelJs) -or -not (Test-Path -LiteralPath $playwright)) {
    throw "Standalone dependencies are missing. Run: powershell -ExecutionPolicy Bypass -File .\setup.ps1"
}

$arguments = @(
    (Join-Path $scriptDirectory "run.mjs"),
    "--period", $Period
)

if ($StartDate) {
    $arguments += @("--start", $StartDate)
}
if ($EndDate) {
    $arguments += @("--end", $EndDate)
}
if ($MaxSites -gt 0) {
    $arguments += @("--max-sites", $MaxSites)
}
if ($Site) {
    $arguments += @("--site", $Site)
}
if ($DryRun) {
    $arguments += "--dry-run"
}
if ($BrowserMode) {
    $arguments += "--browser"
}
if ($Headless) {
    $arguments += "--headless"
}
if ($NonInteractive) {
    $arguments += "--non-interactive"
}
if ($LoginSetup) {
    $arguments += "--login-setup"
}

& $node.Source @arguments
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
