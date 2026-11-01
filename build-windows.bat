@echo off
setlocal
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe py -3.12 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-desktop.txt
if errorlevel 1 goto fail
.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onedir --windowed --name "物料需求需采购计划系统" --collect-all webview desktop.py
if errorlevel 1 goto fail
echo Build complete: dist\物料需求需采购计划系统\物料需求需采购计划系统.exe
pause
exit /b 0
:fail
echo Build failed. Install Python 3.12 and check network access.
pause
exit /b 1
