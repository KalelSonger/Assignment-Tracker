param(
    [switch]$NoIcon
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

$oauthSecretFiles = @(
    "client_secret.json",
    "google_oauth_client_secret.json"
)

& $python -m pip install pyinstaller

$distDir = Join-Path $PSScriptRoot "dist"
$buildDir = Join-Path $PSScriptRoot "build"
$specFile = Join-Path $PSScriptRoot "AssignmentTrackerGUI.spec"

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

foreach ($name in $oauthSecretFiles) {
    $src = Join-Path $PSScriptRoot $name
    if (Test-Path $src) {
        Copy-Item $src (Join-Path $distDir $name) -Force
    }
}

$exe = Join-Path $PSScriptRoot "dist\AssignmentTrackerGUI.exe"
if (Test-Path $exe) {
    Get-Item $exe | Select-Object FullName, Length, LastWriteTime
} else {
    throw "Build finished but EXE was not found at $exe"
}
