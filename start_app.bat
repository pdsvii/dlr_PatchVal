@echo off
setlocal

cd /d "%~dp0"

set "PYTHON=.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
    where python >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON=python"
    ) else (
        echo Python was not found. Install Python or create the .venv folder first.
        exit /b 1
    )
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    %PYTHON% -m venv .venv
    if errorlevel 1 exit /b 1
    set "PYTHON=.venv\Scripts\python.exe"
)

echo Installing requirements...
"%PYTHON%" -m pip install --upgrade pip
if errorlevel 1 exit /b 1
"%PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 exit /b 1

echo Starting Patching Validator...
"%PYTHON%" -m streamlit run src/app.py