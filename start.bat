@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
if exist "%PYTHON_EXE%" "%PYTHON_EXE%" --version >nul 2>nul
if errorlevel 1 set "PYTHON_EXE="
if not defined PYTHON_EXE set "PYTHON_EXE=%~dp0.venv-dev\Scripts\python.exe"
if exist "%PYTHON_EXE%" "%PYTHON_EXE%" --version >nul 2>nul
if errorlevel 1 set "PYTHON_EXE="
if not defined PYTHON_EXE set "PYTHON_EXE=python"
"%PYTHON_EXE%" check_environment.py
if %errorlevel%==1 (
  echo [错误] 存在阻断启动的问题，请查看 environment-report.txt。
  pause
  exit /b 1
)
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://127.0.0.1:5008"
"%PYTHON_EXE%" run.py
pause
