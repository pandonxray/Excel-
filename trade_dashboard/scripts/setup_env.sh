#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

python -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python scripts/bootstrap_env.py

echo "[OK] 环境准备完成"
echo "[NEXT] 激活环境: source .venv/bin/activate"
echo "[NEXT] 启动看板: streamlit run src/dashboard.py"
