@echo off
setlocal

REM Build_Master_Workbook.bat
REM Consolidates every "<Country>_University_PhD_LinkedIn_Postings.xlsx" and
REM "<Country>_University_Masters_Funded_LinkedIn_Postings.xlsx" file sitting in
REM THIS SAME FOLDER into one Excel file with a sheet per country/round plus a
REM Master Summary dashboard.
REM
REM This never touches LinkedIn and never needs a password - it only reads the
REM .xlsx files Claude has already saved into this folder. Double-click it any
REM time after adding or refreshing a country workbook.

cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this PC.
    echo Install it from https://www.python.org/downloads/ ^(check "Add Python to PATH"
    echo during setup^), then run this file again.
    pause
    exit /b 1
)

python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installing the required 'openpyxl' package...
    python -m pip install openpyxl
)

echo.
echo Building master workbook from the LinkedIn postings in this folder...
echo.
python "%~dp0build_master_workbook.py"
if errorlevel 1 (
    echo.
    echo Something went wrong - see the message above.
    pause
    exit /b 1
)

echo.
echo Done. Open All_Countries_Funded_PhD_Masters_Master.xlsx in this folder.
pause
