@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo 请先安装 Python 3.12，安装时勾选 Add Python to PATH。
  start https://www.python.org/downloads/
  pause
  exit /b 1
)
if not exist .venv\Scripts\python.exe (
  py -3 -m venv .venv
  if errorlevel 1 goto fail
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto fail
echo 正在启动系统，请保留此窗口。停止运行请按 Ctrl+C。
.venv\Scripts\python.exe -m streamlit run app.py --server.address 127.0.0.1 --server.port 8501 --browser.gatherUsageStats false
pause
exit /b
:fail
echo 安装失败，请检查网络，并将窗口中的报错发给维护人员。
pause
exit /b 1
