"""
export_to_json.py  v3
同時讀取：
  1. 本機「氣味檢測器市調_分析結果.xlsx」（有 AI 分析欄位 + 發現日期）
  2. Google Sheets 原始資料（補充尚未被 analyze_market.py 處理的新筆數）
合併 → ★以「公司正規化 + 型號/網址簽章」去重（同產品不同名/來源收斂）
     → 由各產品「首次發現」月份算 monthly_trend（修好「每月新增產品數」）
     → 輸出 data/market_data.json，並 git push。

v3 變更（2026-06）：
  - 去重邏輯內建（移植自 VOCsDetector/去重整理.py），取代原本「完全同名」弱去重。
    去重為純字串運算、無 LLM/無 API 成本 → 每次 export 都從頭跑一次即可，不需 checkpoint。
    （需 checkpoint 的是 analyze_market.py 的 LLM 分析，已用 analyzed_at 跳過已分析列。）
  - monthly_trend 由 發現日期 計算（原本寫死成空 {} → 圖表永遠空）。
  - 產品多了 count(收錄次數) / url_count / first_seen / last_seen 欄位。

執行前確認：
  - Google 試算表已設為「知道連結的人可以檢視」（不需 API Key）
  - git remote 已設定好
"""

import pandas as pd
import json
import re
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


# ════════════════════════════════════════════════════════════════
#  ★ 去重邏輯（移植自 VOCsDetector/去重整理.py）
#    以「正規化公司 +（型號碼 or 來源網址）簽章」收斂同一產品的重複收錄。
# ════════════════════════════════════════════════════════════════
JUNK_SUBS = ["generic","unknown","unspecified","undefined","nobrand","no brand","未知","不明",
             "無明確","無品牌","無 (","none","notspecified","various","其他品牌","多家","n/a"]
ALIAS = {  # 正規化鍵(去空白小寫) 子字串 → 正規顯示名
    "tanita":"Tanita","タニタ":"Tanita","塔尼達":"Tanita","타니타":"Tanita",
    "airsense":"AIRSENSE Analytics","盈盛恒泰":"AIRSENSE/盈盛恒泰",
    "i-pex":"I-PEX","ipex":"I-PEX",
    "太陽誘電":"Taiyo Yuden","taiyoyuden":"Taiyo Yuden","yuden":"Taiyo Yuden",
    "owlstone":"Owlstone Medical","quintron":"QuinTron","bosch":"Bosch Sensortec",
    "alphamos":"Alpha MOS","alpham.o.s":"Alpha MOS","bactrack":"BACtrack",
    "foodmarble":"FoodMarble","senseair":"Senseair","ketoscan":"KETOSCAN",
    "figaro":"Figaro","sensirion":"Sensirion","aeris":"Aeris",
    "enose":"The eNose Company","fermion":"Fermion","dfrobot":"DFRobot",
    "exalenz":"Exalenz Bioscience","shownovo":"ShowNovo","首昕":"ShowNovo",
    "ketomojo":"Keto-Mojo","niox":"NIOX","omron":"Omron","オムロン":"Omron",
    "meridian":"Meridian Bioscience","gemelli":"Gemelli Biotech","nanose":"NaNose / Nanose Medical",
}
def norm_co(s):
    s = "" if (s is None or (isinstance(s, float) and pd.isna(s))) else str(s)
    if s == "不明":
        s = ""
    k = re.sub(r"\(.*?\)|（.*?）", "", s).lower()
    k = re.sub(r"[\s\-_/.,，、()（）\[\]:：]+", "", k)
    low = re.sub(r"\s+", "", s.lower())
    if k in ("", "nan") or any(j.replace(" ", "") in k or j.replace(" ", "") in low for j in JUNK_SUBS):
        return "__unknown__"
    for sub, canon in ALIAS.items():
        if sub in k:
            return canon
    return s.strip()

MODEL_RE = re.compile(r"[A-Za-z]{1,6}[\-‐]?\d{1,4}(?:[.\-]?\d+)?[A-Za-z]{0,4}")
STOP_CODES = {"co2","pm2","pm10","no2","h2s","nh3","ppm","ppb","app","usb","2nd","3d","ai","io"}
def codes_of(name):
    cs = {c.lower().replace("-", "").replace("‐", "") for c in MODEL_RE.findall(str(name))}
    return {c for c in cs if not c.isdigit() and len(c) >= 3 and c not in STOP_CODES}
def make_key(co, name, url):
    cs = codes_of(name)
    sig = "m:" + "+".join(sorted(cs)) if cs else "u:" + str(url).strip().lower()
    head = co if co != "__unknown__" else "unk:" + str(url).strip().lower()
    return head + "||" + sig


# ── 收集兩來源的原始列（含發現日期）─────────────────────────────
def collect_raw(df_local: pd.DataFrame, df_cloud: pd.DataFrame) -> list:
    records = []
    seen_names = set()

    if not df_local.empty:
        c = {
            "name":   detect_col(df_local, ["產品名","product_name"]),
            "brand":  detect_col(df_local, ["品牌","公司","brand"]),
            "desc":   detect_col(df_local, ["描述","特色","feature"]),
            "url":    detect_col(df_local, ["網址","url"]),
            "date":   detect_col(df_local, ["發現日期","發現","日期"]),
            "sensor": detect_col(df_local, ["sensor_type"]),
            "form":   detect_col(df_local, ["form_factor"]),
            "prec":   detect_col(df_local, ["precision_tier"]),
            "trl":    detect_col(df_local, ["trl"]),
            "gases":  detect_col(df_local, ["target_gases"]),
            "output": detect_col(df_local, ["output_type"]),
            "eco":    detect_col(df_local, ["ecosystem"]),
            "segs":   detect_col(df_local, ["application_segments"]),
            "moat":   detect_col(df_local, ["competitive_moat"]),
            "feats":  detect_col(df_local, ["key_features"]),
            "conf":   detect_col(df_local, ["confidence"]),
        }
        for _, row in df_local.iterrows():
            name = norm(row.get(c["name"], ""))
            if name == "不明":
                continue
            records.append({
                "name": name, "brand": norm(row.get(c["brand"], "")),
                "description": norm(row.get(c["desc"], "")), "url": norm(row.get(c["url"], "")),
                "date": row.get(c["date"]) if c["date"] else None,
                "sensor_type": norm(row.get(c["sensor"], "")), "form_factor": norm(row.get(c["form"], "")),
                "precision": norm(row.get(c["prec"], "")), "trl": norm(row.get(c["trl"], "")),
                "output_type": norm(row.get(c["output"], "")), "ecosystem": norm(row.get(c["eco"], "")),
                "moat": norm(row.get(c["moat"], "")), "features": norm(row.get(c["feats"], "")),
                "confidence": norm(row.get(c["conf"], "")),
                "gases_raw": norm(row.get(c["gases"], "")), "segs_raw": norm(row.get(c["segs"], "")),
                "source": "analyzed",
            })
            seen_names.add(name.lower().strip())

    if not df_cloud.empty:
        c2 = {
            "name":  detect_col(df_cloud, ["產品名", "product"]),
            "brand": detect_col(df_cloud, ["品牌", "公司", "brand"]),
            "desc":  detect_col(df_cloud, ["描述", "特色", "feature"]),
            "url":   detect_col(df_cloud, ["網址", "url"]),
            "date":  detect_col(df_cloud, ["發現日期", "發現", "日期", "date", "時間"]),
        }
        new_count = 0
        for _, row in df_cloud.iterrows():
            name = norm(row.get(c2["name"], ""))
            if name == "不明" or name.lower().strip() in seen_names:
                continue
            records.append({
                "name": name, "brand": norm(row.get(c2["brand"], "")),
                "description": norm(row.get(c2["desc"], "")), "url": norm(row.get(c2["url"], "")),
                "date": row.get(c2["date"]) if c2["date"] else None,
                "sensor_type": "待分析", "form_factor": "不明", "precision": "不明", "trl": "不明",
                "output_type": "不明", "ecosystem": "不明", "moat": "不明", "features": "不明",
                "confidence": "待分析", "gases_raw": "待分析", "segs_raw": "待分析",
                "source": "cloud_only",
            })
            new_count += 1
        print(f"  ☁️  雲端補入 {new_count} 筆尚未 AI 分析的新資料")

    return records


# ── 簽章去重 → 產品清單 ─────────────────────────────────────────
def dedup_records(records: list) -> list:
    groups = {}
    for r in records:
        co = norm_co(r["brand"])
        key = make_key(co, r["name"], r["url"])
        groups.setdefault(key, []).append((co, r))

    products = []
    for key, items in groups.items():
        co_canon = items[0][0]
        co_disp = co_canon if co_canon != "__unknown__" else "未知/不明"
        recs = [r for _, r in items]

        def mode_f(f):
            vals = [r[f] for r in recs if str(r[f]) not in ("不明", "待分析", "分析失敗", "", "nan")]
            if vals:
                return Counter(vals).most_common(1)[0][0]
            return "待分析" if all(r["source"] != "analyzed" for r in recs) else "不明"

        name = min((r["name"] for r in recs), key=lambda x: (len(x), x))   # 最短乾淨名
        desc = max((r["description"] for r in recs), key=len)              # 最長描述
        gases_raw = mode_f("gases_raw")
        segs_raw  = mode_f("segs_raw")

        ds = [pd.to_datetime(r["date"], errors="coerce") for r in recs if r["date"] is not None]
        ds = [d for d in ds if pd.notna(d)]
        first_seen = min(ds).strftime("%Y-%m-%d") if ds else "不明"
        last_seen  = max(ds).strftime("%Y-%m-%d") if ds else "不明"

        urls = [r["url"] for r in recs if r["url"] != "不明"]
        brand_text = " ".join(r["brand"] for r in recs)

        products.append({
            "name": name, "brand": co_disp, "description": desc,
            "url": urls[0] if urls else "不明",
            "sensor_type": mode_f("sensor_type"), "form_factor": mode_f("form_factor"),
            "precision": mode_f("precision"), "trl": mode_f("trl"),
            "gases": split_tags(gases_raw), "output_type": mode_f("output_type"),
            "ecosystem": mode_f("ecosystem"), "segments": split_tags(segs_raw),
            "moat": mode_f("moat"), "features": mode_f("features"), "confidence": mode_f("confidence"),
            "country": infer_country(brand_text, desc),
            "source": "analyzed" if any(r["source"] == "analyzed" for r in recs) else "cloud_only",
            "count": len(recs), "url_count": len(set(urls)),
            "first_seen": first_seen, "last_seen": last_seen,
        })

    # 收錄次數高（高曝光）排前面
    products.sort(key=lambda p: p["count"], reverse=True)
    return products


# ── 應用場景標準化對應表 ─────────────────────────────────────────
SEGMENT_NORMALIZE = {
    "環境監測": "環境監測", "環境與工業安全": "環境監測", "惡臭監測": "環境監測",
    "空氣品質": "環境監測", "氣體洩漏偵測": "環境監測",
    "工業安全": "工業安全", "工業": "工業安全", "化工廠排放": "工業安全",
    "毒氣偵查": "工業安全", "職業安全": "工業安全",
    "食品品質": "食品品質", "食品鮮度": "食品品質", "食品安全": "食品品質",
    "食品": "食品品質", "鮮度檢測": "食品品質",
    "醫療健康": "醫療健康", "醫療": "醫療健康", "呼氣診斷": "醫療健康",
    "疾病診斷": "醫療健康", "呼吸與氣味偵測": "醫療健康", "呼吸與氣味檢測": "醫療健康",
    "呼吸分析": "醫療健康", "生理氣味監測": "醫療健康",
    "個人健康": "亞健康/個人健康", "亞健康": "亞健康/個人健康", "口臭檢測": "亞健康/個人健康",
    "體臭監測": "亞健康/個人健康", "生酮監測": "亞健康/個人健康",
    "智慧家居": "智慧家居/IoT", "智慧家居/IoT": "智慧家居/IoT", "家電整合": "智慧家居/IoT", "IoT": "智慧家居/IoT",
    "農業": "農業", "農業應用": "農業",
    "半導體製程": "半導體/製造", "半導體": "半導體/製造", "製程監控": "半導體/製造",
    "國防": "國防/安全", "爆炸物偵測": "國防/安全",
    "新品研發": None, "研究": None, "學術研究": None, "不明": None, "待分析": None,
}

def normalize_segment(seg: str):
    seg = seg.strip()
    if seg in SEGMENT_NORMALIZE:
        return SEGMENT_NORMALIZE[seg]
    for k, v in SEGMENT_NORMALIZE.items():
        if k in seg or seg in k:
            return v
    if len(seg) <= 1:
        return None
    return seg


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
            if tag in ("不明","待分析"):
                continue
            if field == "segments":
                tag = normalize_segment(tag)
                if tag is None:
                    continue
            c[tag] += 1
    return dict(c.most_common(top))

def monthly_new_products(products):
    """每月『新增產品數』= 各去重產品依其『首次發現』月份歸戶後計數。"""
    c = Counter()
    for p in products:
        fs = p.get("first_seen", "不明")
        if fs and fs != "不明" and len(fs) >= 7:
            c[fs[:7]] += 1          # YYYY-MM
    return dict(sorted(c.items()))

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
        if p["brand"] not in brands[q] and p["brand"] not in ("不明","","未知/不明"):
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
    print(f"  export_to_json.py v3（內建去重 + 每月新增）")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*52}")

    df_local = load_local_xlsx(LOCAL_XLSX)
    df_cloud = fetch_gsheet(SHEET_ID)

    if df_local.empty and df_cloud.empty:
        sys.exit("❌ 兩個資料來源都失敗，中止")

    raw = collect_raw(df_local, df_cloud)
    products = dedup_records(raw)
    print(f"\n🔁 去重：原始 {len(raw)} 列 → {len(products)} 個產品"
          f"（縮減 {100*(1-len(products)/max(1,len(raw))):.0f}%）")

    analyzed = sum(1 for p in products
                   if p["confidence"] not in ("不明","待分析","分析失敗",""))

    output = {
        "meta": {
            "total":          len(products),
            "brands":         len(set(p["brand"] for p in products
                                      if p["brand"] not in ("不明","","未知/不明"))),
            "countries":      len(set(p["country"] for p in products
                                      if p["country"] != "其他/不明")),
            "analyzed_count": analyzed,
            "pending_count":  len(products) - analyzed,
            "raw_rows":       len(raw),
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
            "monthly_trend":        monthly_new_products(products),
        },
        "quadrant": build_quadrant(products),
        "products": products,  # 已去重，全部筆數
    }

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"💾 已輸出：{OUTPUT_FILE}（{kb:.1f} KB）")
    print(f"   去重後 {len(products)} 個產品｜已分析 {analyzed}｜待分析 {len(products)-analyzed}")
    mt = output["charts"]["monthly_trend"]
    print(f"   每月新增（首次發現）: {mt}")

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
            if "nothing to commit" in (r.stdout or "") + (r.stderr or ""):
                print("  ⚠️  無變更，跳過 commit")
                return
            print(f"  ❌ 失敗：{r.stderr.strip()}")
            return
        print(f"  ✅ {cmd[3]}")
    print("🎉 完成！GitHub Pages 約 1 分鐘後更新")


if __name__ == "__main__":
    main()
