@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Excel Loader

echo ============================================
echo   Excel Loader
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
    echo Recommended: Python 3.11 or 3.12.
    pause
    exit /b 1
  )
)

echo Using interpreter:
%PY% --version
echo.

set "PYTHON="
if exist ".venv\Scripts\python.exe" goto :use_venv

echo Creating virtual environment...
%PY% -m venv .venv
if errorlevel 1 (
  echo [WARN] venv failed, using system Python
  %PY% -m pip install --user --prefer-binary --no-cache-dir -r requirements.txt
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
"%PYTHON%" -m pip install --prefer-binary --no-cache-dir -r requirements.txt
if errorlevel 1 (
  echo.
  echo [WARN] First install failed. Recreating .venv and retrying...
  call deactivate 2>nul
  rmdir /s /q ".venv" 2>nul
  %PY% -m venv .venv
  if errorlevel 1 (
    echo [ERROR] Failed to recreate .venv
    pause
    exit /b 1
  )
  call ".venv\Scripts\activate.bat"
  set "PYTHON=.venv\Scripts\python.exe"
  "%PYTHON%" -m pip install --upgrade pip >nul
  "%PYTHON%" -m pip install --prefer-binary --no-cache-dir -r requirements.txt
  if errorlevel 1 (
    echo.
    echo [ERROR] pip install failed.
    echo 1. Delete folder ".venv"
    echo 2. Install Python 3.11 or 3.12 from python.org
    echo 3. Run start.bat again
    pause
    exit /b 1
  )
)

:start_server
if not exist "TARGET.xlsx" (
  echo [INFO] TARGET.xlsx not found — open any .xlsx via the web UI.
)

echo.
echo Starting server at http://127.0.0.1:8766
echo Use "Open Excel" in the browser to choose a workbook.
echo Or press "Выключить" in the UI to stop and close this window.
echo.
start "" "http://127.0.0.1:8766"
%PYTHON% -m uvicorn app.main:app --host 127.0.0.1 --port 8766
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
  echo.
  echo [ERROR] Server stopped with code %EXITCODE%
  pause
  exit /b %EXITCODE%
)
exit /b 0
