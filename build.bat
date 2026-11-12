@echo off
rem Rebuilds Software_Inventory.xlsx and Software_Inventory_Example.xlsx from the code in tool\.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Please double-click setup.bat first.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" tool\build_workbook.py
pause
