param(
    [switch]$NoIcon,
    [switch]$StopRunningApp = $true
)

$ErrorActionPreference = "Stop"

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    throw "Python executable not found at $python"
}

$entry = Join-Path $PSScriptRoot "AssignmentTrackerGUI.py"
if (-not (Test-Path $entry)) {
    throw "Entry script not found at $entry"
}

$distDir = Join-Path $PSScriptRoot "dist"
$buildDir = Join-Path $PSScriptRoot "build"
$specFile = Join-Path $PSScriptRoot "AssignmentTrackerGUI.spec"
$backupDir = Join-Path $PSScriptRoot ".build_state_backup"

$stateFiles = @(
    "canvas_session.local.json",
    "sheet_endpoints.local.json",
    "app_settings.local.json",
    "keys.local.json",
    "client_secret.json",
    "google_oauth_client_secret.json",
    "google_sheets_token.local.json"
)

if (Test-Path $backupDir) {
    Remove-Item $backupDir -Recurse -Force
}

New-Item -ItemType Directory -Path $backupDir | Out-Null

foreach ($name in $stateFiles) {
    $src = Join-Path $distDir $name
    if (Test-Path $src) {
        Copy-Item $src (Join-Path $backupDir $name) -Force
        continue
    }

    $srcProject = Join-Path $PSScriptRoot $name
    if (Test-Path $srcProject) {
        Copy-Item $srcProject (Join-Path $backupDir $name) -Force
    }
}

$distOutputs = Join-Path $distDir "outputs"
if (Test-Path $distOutputs) {
    Copy-Item $distOutputs (Join-Path $backupDir "outputs") -Recurse -Force
}

& $python -m pip install pyinstaller

if ($StopRunningApp) {
    Get-Process -Name "AssignmentTrackerGUI" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

if (Test-Path $distDir) {
    Remove-Item $distDir -Recurse -Force
}

if (Test-Path $buildDir) {
    Remove-Item $buildDir -Recurse -Force
}

if (Test-Path $specFile) {
    Remove-Item $specFile -Force
}

$args = @(
    "-m", "PyInstaller",
    "--onefile",
    "--windowed",
    "--name", "AssignmentTrackerGUI",
    "--hidden-import", "resvg_py",
    "--hidden-import", "google.auth.transport.requests",
    "--hidden-import", "google.oauth2.credentials",
    "--hidden-import", "google_auth_oauthlib.flow",
    "--hidden-import", "googleapiclient.discovery",
    "--add-data", "app.ico;.",
    "--add-data", "settings.svg;."
)

$iconPath = Join-Path $PSScriptRoot "app.ico"
if (-not $NoIcon -and (Test-Path $iconPath)) {
    $args += @("--icon", $iconPath)
}

$args += $entry

& $python @args

if (-not (Test-Path $distDir)) {
    New-Item -ItemType Directory -Path $distDir | Out-Null
}

foreach ($name in $stateFiles) {
    $saved = Join-Path $backupDir $name
    if (Test-Path $saved) {
        Copy-Item $saved (Join-Path $distDir $name) -Force
    }
}

$savedOutputs = Join-Path $backupDir "outputs"
if (Test-Path $savedOutputs) {
    Copy-Item $savedOutputs (Join-Path $distDir "outputs") -Recurse -Force
}

if (Test-Path $backupDir) {
    Remove-Item $backupDir -Recurse -Force
}

$exe = Join-Path $PSScriptRoot "dist\AssignmentTrackerGUI.exe"
if (Test-Path $exe) {
    Write-Host "Build complete. Preserved local state files in dist/."
    Get-Item $exe | Select-Object FullName, Length, LastWriteTime
} else {
    throw "Build finished but EXE was not found at $exe"
}
