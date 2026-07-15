@echo off
setlocal

set "ROOT_DIR=%~dp0"
if not exist "%ROOT_DIR%\.venv\Scripts\python.exe" (
  python -m venv "%ROOT_DIR%\.venv"
)

call "%ROOT_DIR%\.venv\Scripts\activate.bat"
python -m pip install -r "%ROOT_DIR%\requirements.txt"
python -m streamlit run "%ROOT_DIR%\src\dist_failure_app.py"