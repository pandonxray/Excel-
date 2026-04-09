$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

python -m pip install -r .\requirements-build.txt

$distPath = Join-Path $ProjectRoot "dist"
$buildPath = Join-Path $ProjectRoot "build"
$iconPath = Join-Path $ProjectRoot "assets\wind_dashboard.ico"

if (Test-Path $distPath) { Remove-Item $distPath -Recurse -Force }
if (Test-Path $buildPath) { Remove-Item $buildPath -Recurse -Force }

python -m PyInstaller `
  --noconfirm `
  --clean `
  --onedir `
  --name WindDashboard `
  --collect-all streamlit `
  --collect-all plotly `
  --collect-all pyarrow `
  --hidden-import pandas `
  --hidden-import openpyxl `
  --hidden-import yaml `
  --icon $iconPath `
  --add-data "config;config" `
  --add-data "src;src" `
  src\launcher.py

Write-Host ""
Write-Host "[OK] 打包完成"
Write-Host "[PATH] $ProjectRoot\dist\WindDashboard\WindDashboard.exe"
