@echo off
setlocal

REM Bootstrap + run the AS9102 FAI GUI on Windows without PowerShell policy changes.
REM - Creates/repairs .venv (venvs are machine-specific; do not copy between PCs)
REM - Installs app dependencies
REM - Launches: python -m as9102_fai

cd /d "%~dp0\.." || exit /b 1

set "VENV_PY=.venv\Scripts\python.exe"

REM Prefer Python 3.12 via the Windows py launcher if available.
set "PY_CMD=py -3.12"
%PY_CMD% -c "import sys" >nul 2>nul
if errorlevel 1 (
    set "PY_CMD=py -3.13"
    %PY_CMD% -c "import sys" >nul 2>nul
    if errorlevel 1 (
        set "PY_CMD=python"
    )
)

REM Check if venv Python works. If not, recreate it.
%VENV_PY% -c "import sys" >nul 2>nul
if errorlevel 1 (
    echo Recreating .venv...
    if exist .venv rmdir /s /q .venv
    %PY_CMD% -m venv .venv || exit /b 1
    %VENV_PY% -m ensurepip --upgrade >nul 2>nul
    %VENV_PY% -m pip install --upgrade pip || exit /b 1
)

echo Installing dependencies...
%VENV_PY% -m pip install -r as9102_fai\requirements.txt || exit /b 1

echo Launching AS9102 FAI...
%VENV_PY% -m as9102_fai
