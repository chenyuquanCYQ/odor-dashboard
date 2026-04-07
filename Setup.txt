# 氣味檢測器 Dashboard 完整部署流程

## 所有需要的檔案

```
odor-dashboard/          ← GitHub repo 根目錄（也是你電腦上的本機資料夾）
├── index.html           ← Dashboard 網頁
├── export_to_json.py    ← 資料轉換 + git push
├── auto_update.bat      ← 每日自動排程腳本
├── SETUP.md             ← 這份說明
└── data/
    └── .gitkeep         ← 占位符（已在 GitHub 建立）
```

---

## STEP 1：設定 Google 試算表為公開

1. 開啟你的 Google 試算表
   https://docs.google.com/spreadsheets/d/1AyPfZdyqrkIwsLaKh7BNLqlgRhcXOXT9vrakUoeos0I

2. 右上角點「共用」

3. 一般存取 → 改為「知道連結的人」→「檢視者」

4. 完成（不需要 API Key，export_to_json.py 直接下載 CSV）

---

## STEP 2：把 GitHub repo clone 到本機

```bash
# 在你想放的位置開啟命令提示字元
cd D:\01-AIModel

git clone https://github.com/chenyuquanCYQ/odor-dashboard.git
```

clone 完後你的資料夾會是：
```
D:\01-AIModel\odor-dashboard\
├── auto_update.bat
├── export_to_json.py
├── index.html
└── data\
    └── .gitkeep
```

---

## STEP 3：修改 export_to_json.py 的路徑設定

用記事本或 VS Code 開啟 `export_to_json.py`，
找到第 23–29 行，確認這三個設定：

```python
# ★ 你的 Google 試算表 ID（已經填好，不需修改）
SHEET_ID = "1AyPfZdyqrkIwsLaKh7BNLqlgRhcXOXT9vrakUoeos0I"

# ★ 本機分析結果 xlsx 的完整路徑（修改成你的實際路徑）
LOCAL_XLSX = r"D:\01-AIModel\VOCsDetector\氣味檢測器市調_分析結果.xlsx"

# ★ Dashboard repo 位置（不需修改，自動偵測）
REPO_DIR = Path(__file__).parent
```

---

## STEP 4：修改 auto_update.bat 的路徑

用記事本開啟 `auto_update.bat`，
找到第 18 行，改成你的 analyze_market.py 實際路徑：

```bat
:: 改這行
python "D:\01-AIModel\VOCsDetector\analyze_market.py"

:: 改這行（xlsx 複製來源）
copy /Y "D:\01-AIModel\VOCsDetector\氣味檢測器市調_分析結果.xlsx" "%~dp0氣味檢測器市調_分析結果.xlsx"
```

---

## STEP 5：第一次手動執行測試

開啟命令提示字元，進入 repo 目錄：

```bash
cd D:\01-AIModel\odor-dashboard
python export_to_json.py
```

正常輸出應該像這樣：
```
====================================================
  export_to_json.py v2
  2025-01-15 10:30:00
====================================================
  ✅ 本機 xlsx 取得 1200 筆（含 AI 分析欄位）
📡 從 Google Sheets 讀取原始資料...
  ✅ 雲端取得 1350 筆
  ☁️  雲端補入 150 筆尚未 AI 分析的新資料

📊 合併後共 1350 筆
💾 已輸出：data\market_data.json（285.3 KB）
   已分析 1200 筆 ／ 待分析 150 筆

🚀 Git 推送...
  ✅ add
  ✅ commit
  ✅ push
🎉 完成！GitHub Pages 約 1 分鐘後更新
```

---

## STEP 6：開啟 GitHub Pages

1. 到 https://github.com/chenyuquanCYQ/odor-dashboard

2. 點上方 **Settings** 標籤

3. 左側選單點 **Pages**

4. Source → **Deploy from a branch**

5. Branch → **main** → **/ (root)** → 點 **Save**

6. 等約 1–2 分鐘，頁面上會出現：
   ```
   Your site is live at https://chenyuquancyq.github.io/odor-dashboard/
   ```

---

## STEP 7：設定 Windows 工作排程器（每日自動更新）

1. 搜尋「工作排程器」並開啟

2. 右側點「建立基本工作」

3. 名稱：`odor-dashboard 每日更新`

4. 觸發程序：每天 → 設定你想要的時間（例如 08:00）

5. 動作：啟動程式

6. 程式/指令碼：
   ```
   D:\01-AIModel\odor-dashboard\auto_update.bat
   ```

7. 完成

---

## 完整流程示意

```
每天 08:00
    ↓
auto_update.bat 自動執行
    ↓
① analyze_market.py（Ollama + Gemma）
   只處理新增筆數，跳過已分析的
    ↓
② export_to_json.py
   讀本機 xlsx（有AI分析）+ Google Sheets（新資料）
   合併輸出 data/market_data.json
    ↓
③ git push → GitHub
    ↓
GitHub Pages 自動更新
    ↓
https://chenyuquancyq.github.io/odor-dashboard/
顯示最新資料
```

---

## 常見問題

**Q: Google Sheets 讀取失敗？**
→ 確認試算表共用設定是「知道連結的人可以檢視」

**Q: git push 失敗，提示需要登入？**
→ 在命令提示字元執行：
```bash
git config --global credential.helper manager
git push  # 會跳出 GitHub 登入視窗，登入一次後就記住了
```

**Q: Dashboard 打開是空白或顯示「無法載入資料」？**
→ 確認 `data/market_data.json` 已經成功推送到 GitHub
→ 確認 GitHub Pages 已開啟（Settings → Pages）

**Q: analyze_market.py 跑太慢，可以跳過嗎？**
→ 可以。在 auto_update.bat 裡把步驟 1 那兩行加上 `::` 變成註解，
  只執行 export_to_json.py 也沒問題，
  Dashboard 仍會顯示最新雲端資料（新筆數標示為「待分析」）