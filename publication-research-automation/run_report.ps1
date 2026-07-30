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
$runtimeRoot = Join-Path $env:USERPROFILE ".cache\codex-runtimes\codex-primary-runtime\dependencies\node"
$bundledNode = Join-Path $runtimeRoot "bin\node.exe"
$bundledModules = Join-Path $runtimeRoot "node_modules"
$localModules = Join-Path $scriptDirectory "node_modules"

if (-not (Test-Path -LiteralPath $bundledNode)) {
    throw "Bundled Node.js was not found at $bundledNode. Open this folder in Codex once, or update `$bundledNode in run_report.ps1."
}
if (-not (Test-Path -LiteralPath $bundledModules)) {
    throw "Bundled Node.js modules were not found at $bundledModules."
}
if (-not (Test-Path -LiteralPath $localModules)) {
    New-Item -ItemType Junction -Path $localModules -Target $bundledModules | Out-Null
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

& $bundledNode @arguments
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
