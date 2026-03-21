@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" -m streamlit run src/dashboard.py --server.headless true --server.port 8501
) else (
  python -m streamlit run src/dashboard.py --server.headless true --server.port 8501
)
