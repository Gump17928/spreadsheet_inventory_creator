@echo off
rem Creates the project's virtual environment and installs what it needs. Safe to run again.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 (
        echo Could not create the virtual environment. Is Python 3.11 or newer installed?
        pause
        exit /b 1
    )
)
".venv\Scripts\python.exe" -m pip install --quiet -r requirements-dev.txt
echo Setup complete.
pause
