param(
    [switch]$Clean
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPath = Join-Path $ProjectRoot ".venv"
$PythonExe = Join-Path $VenvPath "Scripts\python.exe"
$DistApp = Join-Path $ProjectRoot "dist\MigrationFalloutDashboard"

Set-Location $ProjectRoot

if ($Clean) {
    Remove-Item -LiteralPath (Join-Path $ProjectRoot "build") -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path $ProjectRoot "dist") -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath (Join-Path $ProjectRoot "MigrationFalloutDashboard.spec") -Force -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath $PythonExe)) {
    python -m venv $VenvPath
}

& $PythonExe -m pip install --upgrade pip
& $PythonExe -m pip install -r requirements.txt

New-Item -ItemType Directory -Force -Path (Join-Path $ProjectRoot "data") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $ProjectRoot "data\attachments") | Out-Null

& $PythonExe -m PyInstaller `
    --noconfirm `
    --clean `
    --onedir `
    --name MigrationFalloutDashboard `
    --collect-all streamlit `
    --collect-all keyring `
    --hidden-import win32timezone `
    --add-data "app.py;." `
    --add-data "ai_insights.py;." `
    --add-data "config.py;." `
    --add-data "credential_store.py;." `
    --add-data "database.py;." `
    --add-data "email_matcher.py;." `
    --add-data "excel_parser.py;." `
    --add-data "filename_utils.py;." `
    --add-data "outlook_scanner.py;." `
    --add-data "report_comparator.py;." `
    --add-data "data;data" `
    launcher.py

Write-Host ""
Write-Host "Build complete:"
Write-Host $DistApp
Write-Host "Run by double-clicking MigrationFalloutDashboard.exe inside that folder."
