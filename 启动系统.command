#!/bin/bash
cd "$(dirname "$0")" || exit 1
if ! command -v python3 >/dev/null 2>&1; then
  echo '请先从 https://www.python.org/downloads/ 安装 Python 3.12。'
  read -r -p '按回车退出'
  exit 1
fi
python3 -m venv .venv || exit 1
.venv/bin/python -m pip install -r requirements.txt || { read -r -p '安装失败，请检查网络，按回车退出'; exit 1; }
.venv/bin/python -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --browser.gatherUsageStats false
