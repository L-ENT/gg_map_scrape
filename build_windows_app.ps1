[CmdletBinding()]
param(
  [string]$Version = 'local'
)

$ErrorActionPreference = 'Stop'
$projectRoot = $PSScriptRoot
$buildVenv = Join-Path $projectRoot '.packaging-venv'
$pythonExe = Join-Path $buildVenv 'Scripts\python.exe'

if (Get-Command py -ErrorAction SilentlyContinue) {
    & py -3.11 -m venv $buildVenv
} elseif (Get-Command python -ErrorAction SilentlyContinue) {
    # Some Windows setups have Python on PATH but do not install the `py`
    # launcher. Any supported Python 3.11+ interpreter can build this app.
    & python -m venv $buildVenv
} else {
    throw 'Python 3.11 or newer was not found. Install it from python.org, then run this script again.'
}
& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r (Join-Path $projectRoot 'requirements.txt') pyinstaller

& $pythonExe -m PyInstaller --noconfirm --clean --onedir --windowed `
  --name 'Clinic Lead Collector' `
  --add-data "$projectRoot\app.py;." `
  --collect-all streamlit `
  --collect-all pandas `
  --collect-all pyarrow `
  --collect-all selenium `
  --collect-all webdriver_manager `
  --collect-all openpyxl `
  (Join-Path $projectRoot 'desktop_launcher.py')

& $pythonExe -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name 'Clinic Lead Updater' `
  (Join-Path $projectRoot 'updater.py')

$appFolder = Join-Path $projectRoot 'dist\Clinic Lead Collector'
Copy-Item (Join-Path $projectRoot 'dist\Clinic Lead Updater.exe') (Join-Path $appFolder 'Clinic Lead Updater.exe') -Force
Set-Content -Path (Join-Path $appFolder 'version.txt') -Value $Version -NoNewline -Encoding utf8

$zipPath = Join-Path $projectRoot 'dist\Clinic-Lead-Collector-windows.zip'
if (Test-Path $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
Compress-Archive -Path $appFolder -DestinationPath $zipPath -Force

Write-Host ''
Write-Host 'Build complete. Send this ZIP file:' -ForegroundColor Green
Write-Host "  $zipPath"
Write-Host 'The recipient extracts it, then opens Clinic Lead Collector.exe. Google Chrome must be installed on their Windows PC.'
