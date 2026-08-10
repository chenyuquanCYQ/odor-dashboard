@echo off
chcp 65001 >nul
setlocal
cd /d %~dp0
set "PYTHONIOENCODING=utf-8"
set "LOG=%~dp0update_log.txt"

echo [%date% %time%] ===== Starting daily update ===== >> "%LOG%"
echo [%date% %time%] Starting daily update...

:: Step 1: Run analyze_market.py
:: Reads latest data from Google Sheets, analyzes new rows only, saves to local xlsx
echo [Step 1] Running analyze_market.py...
python "D:\02-AIProject\VOCsDetector\analyze_market.py" >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [WARNING] analyze_market.py failed, continuing...
    echo [%date% %time%] [WARNING] analyze_market.py failed, continuing... >> "%LOG%"
)

:: Step 2: Export JSON (dedup + monthly trend) and push to GitHub
echo [Step 2] Running export_to_json.py...
python "D:\02-AIProject\odor-dashboard\export_to_json.py" >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] export_to_json.py FAILED - dashboard NOT updated ^(see PUSH_FAILED.txt^)
    echo [%date% %time%] [ERROR] export_to_json.py FAILED - see PUSH_FAILED.txt >> "%LOG%"
    goto :end
)

echo [%date% %time%] Done! >> "%LOG%"
echo [%date% %time%] Done!
:end
endlocal
