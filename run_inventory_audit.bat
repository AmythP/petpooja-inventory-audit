@echo off
setlocal
title Kitchen Daily Inventory Audit
echo ============================================================
echo   DAILY INVENTORY AUDIT SYSTEM
echo ============================================================
echo.

REM Detect Hermes Agent bundled Python. No separate Python install needed.
set "PYEXE="
for /d %%D in ("%LOCALAPPDATA%\hermes\tools\python-*") do (
    if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
)
if "%PYEXE%"=="" (
    echo ERROR: Python not found.
    pause
    exit /b 1
)

set "SCRIPT=%~dp0src\run_inventory_audit.py"
set "INVENTORY_AUDIT_DIR=%CD%"

echo Processing Petpooja consumption and rebuilding workbook...
echo.
"%PYEXE%" "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo FAILED. Make sure Petpooja '.xls' consumption and stock files are present.
) else (
    echo.
    echo DONE. Open production+closing=yersterday.xlsx
    echo Yellow highlighted cells are for your manual data entry.
)
echo.
pause
endlocal
