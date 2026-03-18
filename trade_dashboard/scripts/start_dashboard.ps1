$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot
python -m streamlit run src/dashboard.py --server.headless true --server.port 8501
