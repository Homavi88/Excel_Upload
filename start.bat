@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo   Excel Loader for TARGET.xlsx
echo ============================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
) else (
  where python >nul 2>nul
  if %errorlevel%==0 (
    set "PY=python"
  ) else (
    echo [ERROR] Python 3 not found. Install from https://www.python.org/downloads/
    echo Make sure "Add python.exe to PATH" is enabled.
    pause
    exit /b 1
  )
)

set "PYTHON="
if exist ".venv\Scripts\python.exe" goto :use_venv

echo Creating virtual environment...
%PY% -m venv .venv
if errorlevel 1 (
  echo [WARN] venv failed, using system Python
  %PY% -m pip install --user -r requirements.txt
  if errorlevel 1 (
    echo [ERROR] pip install failed
    pause
    exit /b 1
  )
  set "PYTHON=%PY%"
  goto :start_server
)

:use_venv
call ".venv\Scripts\activate.bat"
set "PYTHON=.venv\Scripts\python.exe"
echo Installing dependencies...
"%PYTHON%" -m pip install --upgrade pip >nul
"%PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 (
  echo [ERROR] pip install failed
  pause
  exit /b 1
)

:start_server
if not exist "TARGET.xlsx" (
  echo [ERROR] TARGET.xlsx not found in %cd%
  pause
  exit /b 1
)

echo.
echo Starting server at http://127.0.0.1:8765
echo Close this window to stop the loader.
echo.
start "" "http://127.0.0.1:8765"
%PYTHON% -m uvicorn app.main:app --host 127.0.0.1 --port 8765
pause
