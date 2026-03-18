@echo off
setlocal
cd /d "%~dp0"
python -m streamlit run src/dashboard.py --server.headless true --server.port 8501
