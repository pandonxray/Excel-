$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $Root

python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python .\scripts\bootstrap_env.py

Write-Host "[OK] 环境准备完成"
Write-Host "[NEXT] 激活环境: .\\.venv\\Scripts\\Activate.ps1"
Write-Host "[NEXT] 启动看板: streamlit run src/dashboard.py"
