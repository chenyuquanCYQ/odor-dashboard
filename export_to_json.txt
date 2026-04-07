"""
export_to_json.py  v2
同時讀取：
  1. 本機「氣味檢測器市調_分析結果.xlsx」（有 AI 分析欄位）
  2. Google Sheets 原始資料（補充尚未被 analyze_market.py 處理的新筆數）
合併後輸出 data/market_data.json，並 git push。

執行前確認：
  - Google 試算表已設為「知道連結的人可以檢視」（不需 API Key）
  - git remote 已設定好
"""

import pandas as pd
import json
import subprocess
import sys
import urllib.request
import io
from pathlib import Path
from datetime import datetime
from collections import Counter

# ════════════════════════════════════════════════════════
#  ★ 請修改這三個設定
# ════════════════════════════════════════════════════════

# 你的 Google 試算表 ID（網址中 /d/ 後面那串）
SHEET_ID = "1AyPfZdyqrkIwsLaKh7BNLqlgRhcXOXT9vrakUoeos0I"

# 本機分析結果 xlsx 路徑（analyze_market.py 的輸出）
LOCAL_XLSX = r"D:\02-AIProject\VOCsDetector\氣味檢測器市調_分析結果.xlsx"

# Dashboard git repo 根目錄（這支 py 檔所在位置）
REPO_DIR = Path(__file__).parent

# ════════════════════════════════════════════════════════
OUTPUT_FILE   = REPO_DIR / "data" / "market_data.json"
AUTO_GIT_PUSH = True


# ── Google Sheets 直接下載 CSV（不需要 API Key）──────────────────
def fetch_gsheet(sheet_id: str, gid: str = "0") -> pd.DataFrame:
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    print(f"📡 從 Google Sheets 讀取原始資料...")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            df = pd.read_csv(io.BytesIO(resp.read()))
        print(f"  ✅ 雲端取得 {len(df)} 筆")
        return df
    except Exception as e:
        print(f"  ⚠️  Google Sheets 讀取失敗：{e}")
        print(f"     請確認試算表已設為「知道連結的人可以檢視」")
        return pd.DataFrame()


# ── 讀取本機分析結果 xlsx ────────────────────────────────────────
def load_local_xlsx(path: str) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        print(f"  ⚠️  本機 xlsx 不存在：{path}（跳過，只用雲端資料）")
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, sheet_name="原始+分析")
        print(f"  ✅ 本機 xlsx 取得 {len(df)} 筆（含 AI 分析欄位）")
        return df
    except Exception as e:
        print(f"  ⚠️  本機 xlsx 讀取失敗：{e}")
        return pd.DataFrame()


# ── 欄位自動偵測 ─────────────────────────────────────────────────
def detect_col(df, keywords):
    for k in keywords:
        for c in df.columns:
            if k.lower() in str(c).lower():
                return c
    return None


# ── 標準化 ───────────────────────────────────────────────────────
def norm(val) -> str:
    if pd.isna(val) or str(val).strip() in ("", "nan", "None", "NaN"):
        return "不明"
    return str(val).strip()

def split_tags(val: str) -> list:
    if val in ("不明", "待分析"):
        return [val]
    parts = [p.strip() for p in val.replace("、", ",").split(",")]
    return [p for p in parts if p] or ["不明"]


# ── 品牌來源國推斷 ───────────────────────────────────────────────
COUNTRY_HINTS = {
    "日本":   ["tanita","shinyei","figaro","new cosmos","riken","日本","東京","大阪","japan"],
    "美國":   ["honeywell","alphasense","owlstone","breathid","foodmarble","quinTron",
               "circassia","usa","inc.","llc","corp."],
    "德國":   ["airsense","airsense analytics","gmbh","deutschland","germany"],
    "中國":   ["shownovo","中国","深圳","上海","北京","广州","beijing","shenzhen",
               "shanghai","guangzhou","科技有限","china"],
    "台灣":   ["ainos","kwangyuen","廣運","台灣","taiwan","儕仕","konos"],
    "英國":   ["owlstone","e2v","uk","united kingdom","england"],
    "法國":   ["alpha m.o.s","alpha mos","france","s.a."],
    "荷蘭":   ["enose","the enose","netherlands","b.v."],
    "韓國":   ["korea","samsung","한국"],
    "以色列": ["breathid","exalenz","israel"],
    "瑞士":   ["sensirion","switzerland","swiss"],
}

def infer_country(brand: str, desc: str) -> str:
    text = (brand + " " + desc).lower()
    for country, hints in COUNTRY_HINTS.items():
        for h in hints:
            if h.lower() in text:
                return country
    return "其他/不明"


# ── 合併兩個資料來源 ─────────────────────────────────────────────
def merge_sources(df_local: pd.DataFrame, df_cloud: pd.DataFrame) -> list:
    products = {}

    # 1. 先處理本機 xlsx（有完整 AI 分析欄位，優先）
    if not df_local.empty:
        c = {
            "name":    detect_col(df_local, ["產品名","product_name"]),
            "brand":   detect_col(df_local, ["品牌","公司","brand"]),
            "desc":    detect_col(df_local, ["描述","特色","feature"]),
            "url":     detect_col(df_local, ["網址","url"]),
            "sensor":  detect_col(df_local, ["sensor_type"]),
            "form":    detect_col(df_local, ["form_factor"]),
            "prec":    detect_col(df_local, ["precision_tier"]),
            "trl":     detect_col(df_local, ["trl"]),
            "gases":   detect_col(df_local, ["target_gases"]),
            "output":  detect_col(df_local, ["output_type"]),
            "eco":     detect_col(df_local, ["ecosystem"]),
            "segs":    detect_col(df_local, ["application_segments"]),
            "moat":    detect_col(df_local, ["competitive_moat"]),
            "feats":   detect_col(df_local, ["key_features"]),
            "conf":    detect_col(df_local, ["confidence"]),
        }
        for _, row in df_local.iterrows():
            name  = norm(row.get(c["name"],  ""))
            brand = norm(row.get(c["brand"], ""))
            desc  = norm(row.get(c["desc"],  ""))
            if name == "不明":
                continue
            key = name.lower().strip()
            products[key] = {
                "name":        name,
                "brand":       brand,
                "description": desc,
                "url":         norm(row.get(c["url"],    "")),
                "sensor_type": norm(row.get(c["sensor"], "")),
                "form_factor": norm(row.get(c["form"],   "")),
                "precision":   norm(row.get(c["prec"],   "")),
                "trl":         norm(row.get(c["trl"],    "")),
                "gases":       split_tags(norm(row.get(c["gases"],  ""))),
                "output_type": norm(row.get(c["output"], "")),
                "ecosystem":   norm(row.get(c["eco"],    "")),
                "segments":    split_tags(norm(row.get(c["segs"],   ""))),
                "moat":        norm(row.get(c["moat"],   "")),
                "features":    norm(row.get(c["feats"],  "")),
                "confidence":  norm(row.get(c["conf"],   "")),
                "country":     infer_country(brand, desc),
                "source":      "analyzed",
            }

    # 2. 雲端有但本機沒有的，補入為「待分析」
    if not df_cloud.empty:
        c2 = {
            "name":  detect_col(df_cloud, ["產品名","product"]),
            "brand": detect_col(df_cloud, ["品牌","公司","brand"]),
            "desc":  detect_col(df_cloud, ["描述","特色","feature"]),
            "url":   detect_col(df_cloud, ["網址","url"]),
        }
        new_count = 0
        for _, row in df_cloud.iterrows():
            name  = norm(row.get(c2["name"],  ""))
            brand = norm(row.get(c2["brand"], ""))
            desc  = norm(row.get(c2["desc"],  ""))
            if name == "不明":
                continue
            key = name.lower().strip()
            if key not in products:
                products[key] = {
                    "name":        name,
                    "brand":       brand,
                    "description": desc,
                    "url":         norm(row.get(c2["url"], "")),
                    "sensor_type": "待分析",
                    "form_factor": "不明",
                    "precision":   "不明",
                    "trl":         "不明",
                    "gases":       ["待分析"],
                    "output_type": "不明",
                    "ecosystem":   "不明",
                    "segments":    ["待分析"],
                    "moat":        "不明",
                    "features":    "不明",
                    "confidence":  "待分析",
                    "country":     infer_country(brand, desc),
                    "source":      "cloud_only",
                }
                new_count += 1
        print(f"  ☁️  雲端補入 {new_count} 筆尚未 AI 分析的新資料")

    return list(products.values())


# ── 統計彙總 ─────────────────────────────────────────────────────
ECOSYSTEM_SCORE = {"無":0,"藍牙App":1,"IoT雲端":2,"AI驅動":3,"SaaS平台":4,"多種整合":4}
PRECISION_SCORE = {"消費電子級":1,"工業級":2,"實驗室級":3}

def count_field(products, field, top=20):
    c = Counter(p[field] for p in products
                if p[field] not in ("不明","待分析","分析失敗",""))
    return dict(c.most_common(top))

def count_list_field(products, field, top=20):
    c = Counter()
    for p in products:
        for tag in p[field]:
            if tag not in ("不明","待分析"):
                c[tag] += 1
    return dict(c.most_common(top))

def build_quadrant(products):
    counts = Counter()
    brands = {"Q1":[],"Q2":[],"Q3":[],"Q4":[]}
    for p in products:
        if p["confidence"] == "待分析":
            continue
        x = max((v for k,v in ECOSYSTEM_SCORE.items() if k in p["ecosystem"]), default=0)
        y = max((v for k,v in PRECISION_SCORE.items() if k in p["precision"]), default=1)
        q = "Q1" if x>=2 and y>=2 else "Q2" if x>=2 else "Q3" if y>=2 else "Q4"
        counts[q] += 1
        if p["brand"] not in brands[q] and p["brand"] not in ("不明",""):
            brands[q].append(p["brand"])
    return {
        "counts": dict(counts),
        "brands": {q: list(set(v))[:8] for q,v in brands.items()},
        "labels": {
            "Q1":"Q1 高精度＋強生態（平台型）",
            "Q2":"Q2 消費IoT型",
            "Q3":"Q3 高精度儀器型",
            "Q4":"Q4 基礎感測型",
        }
    }


# ── 主流程 ───────────────────────────────────────────────────────
def main():
    print(f"\n{'='*52}")
    print(f"  export_to_json.py v2")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*52}")

    df_local = load_local_xlsx(LOCAL_XLSX)
    df_cloud = fetch_gsheet(SHEET_ID)

    if df_local.empty and df_cloud.empty:
        sys.exit("❌ 兩個資料來源都失敗，中止")

    products = merge_sources(df_local, df_cloud)
    print(f"\n📊 合併後共 {len(products)} 筆")

    analyzed = sum(1 for p in products
                   if p["confidence"] not in ("不明","待分析","分析失敗",""))

    output = {
        "meta": {
            "total":          len(products),
            "brands":         len(set(p["brand"] for p in products
                                      if p["brand"] not in ("不明",""))),
            "countries":      len(set(p["country"] for p in products
                                      if p["country"] != "其他/不明")),
            "analyzed_count": analyzed,
            "pending_count":  len(products) - analyzed,
            "updated_at":     datetime.now().isoformat(timespec="seconds"),
        },
        "charts": {
            "sensor_type":          count_field(products, "sensor_type"),
            "form_factor":          count_field(products, "form_factor"),
            "precision_tier":       count_field(products, "precision"),
            "trl":                  count_field(products, "trl"),
            "output_type":          count_field(products, "output_type"),
            "ecosystem":            count_field(products, "ecosystem"),
            "competitive_moat":     count_field(products, "moat"),
            "target_gases":         count_list_field(products, "gases"),
            "application_segments": count_list_field(products, "segments"),
            "country":              count_field(products, "country"),
        },
        "quadrant": build_quadrant(products),
        "products": products[:600],
    }

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"💾 已輸出：{OUTPUT_FILE}（{kb:.1f} KB）")
    print(f"   已分析 {analyzed} 筆 ／ 待分析 {len(products)-analyzed} 筆")

    if AUTO_GIT_PUSH:
        git_push(REPO_DIR)


def git_push(repo_dir: Path):
    print("\n🚀 Git 推送...")
    cmds = [
        ["git", "-C", str(repo_dir), "add", "data/market_data.json"],
        ["git", "-C", str(repo_dir), "commit", "-m",
         f"data: auto update {datetime.now().strftime('%Y-%m-%d %H:%M')}"],
        ["git", "-C", str(repo_dir), "push"],
    ]
    for cmd in cmds:
        r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
        if r.returncode != 0:
            if "nothing to commit" in r.stdout + r.stderr:
                print("  ⚠️  無變更，跳過 commit")
                return
            print(f"  ❌ 失敗：{r.stderr.strip()}")
            return
        print(f"  ✅ {cmd[3]}")
    print("🎉 完成！GitHub Pages 約 1 分鐘後更新")


if __name__ == "__main__":
    main()