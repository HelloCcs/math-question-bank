@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

set "PY_CMD="
where py >nul 2>nul
if %errorlevel%==0 set "PY_CMD=py -3"
if not defined PY_CMD (
  where python >nul 2>nul
  if %errorlevel%==0 set "PY_CMD=python"
)

if not defined PY_CMD (
  echo [错误] 没有找到 Python。
  echo 请先安装 Python 3.11 或以上版本，并勾选“Add Python to PATH”。
  pause
  exit /b 1
)

echo [1/4] 创建虚拟环境...
%PY_CMD% -m venv .venv
if errorlevel 1 (
  echo [错误] 创建虚拟环境失败。
  pause
  exit /b 1
)

set "VENV_PY=%~dp0.venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo [错误] 虚拟环境创建失败，未找到 %VENV_PY%
  pause
  exit /b 1
)

echo [2/4] 升级 pip...
"%VENV_PY%" -m pip install --upgrade pip
if errorlevel 1 (
  echo [错误] pip 升级失败。
  pause
  exit /b 1
)

echo [3/4] 安装项目依赖...
"%VENV_PY%" -m pip install -r requirements.txt
if errorlevel 1 (
  echo [错误] 依赖安装失败。
  pause
  exit /b 1
)

if exist "%~dp0.venv\Scripts\pywin32_postinstall.py" (
  echo [4/4] 注册 pywin32...
  "%VENV_PY%" "%~dp0.venv\Scripts\pywin32_postinstall.py" -install >nul 2>nul
)

echo.
echo [4/4] 检查运行环境...
"%VENV_PY%" check_environment.py
echo.
echo 安装完成。请查看 environment-report.txt，然后双击 start.bat。
pause
