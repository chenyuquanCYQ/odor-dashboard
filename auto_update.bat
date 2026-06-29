@echo off
chcp 65001 >nul
setlocal
cd /d %~dp0

echo [%date% %time%] Starting daily update...

:: Step 1: Run analyze_market.py
:: Reads latest data from Google Sheets, analyzes new rows only, saves to local xlsx
echo [Step 1] Running analyze_market.py...
python "D:\02-AIProject\VOCsDetector\analyze_market.py"
if %errorlevel% neq 0 (
    echo [WARNING] analyze_market.py failed, continuing...
)

:: Step 2: Export JSON and push to GitHub
echo [Step 2] Running export_to_json.py...
python "D:\02-AIProject\odor-dashboard\export_to_json.py"
if %errorlevel% neq 0 (
    echo [ERROR] export_to_json.py failed
    goto :end
)

echo [%date% %time%] Done!
:end
endlocal