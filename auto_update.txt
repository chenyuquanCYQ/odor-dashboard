@echo off
:: ============================================================
:: auto_update.bat
:: 每天自動執行：分析 → 輸出 JSON → 推送 GitHub
:: 放在 Dashboard 專案根目錄，用 Windows 工作排程器呼叫
:: ============================================================

setlocal
cd /d %~dp0

echo [%date% %time%] 開始每日自動更新...

:: 1. 執行 analyze_market.py（地端 Ollama 分析新增資料）
echo [步驟 1] 執行 LLM 分析...
python "D:\02-AIProject\VOCsDetector\analyze_market.py"
if %errorlevel% neq 0 (
    echo [錯誤] analyze_market.py 執行失敗，跳過後續步驟
    goto :end
)

:: 2. 把分析結果複製到 Dashboard 目錄（如果路徑不同）
echo [步驟 2] 複製分析結果...
copy /Y "D:\02-AIProject\VOCsDetector\氣味檢測器市調_分析結果.xlsx" "%~dp0氣味檢測器市調_分析結果.xlsx"

:: 3. 輸出 JSON 並推送 GitHub
echo [步驟 3] 輸出 JSON 並推送 GitHub...
python "%~dp0export_to_json.py"
if %errorlevel% neq 0 (
    echo [錯誤] export_to_json.py 執行失敗
    goto :end
)

echo [%date% %time%] 全部完成！
:end
endlocal
pause