$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (Test-Path $VenvPython) {
    & $VenvPython -m streamlit run src/dashboard.py --server.headless true --server.port 8501
} else {
    python -m streamlit run src/dashboard.py --server.headless true --server.port 8501
}
