# excel2json.py — 自動フォールバック対応版
import os
import json
from datetime import datetime, timedelta, timezone
import pandas as pd
from collections import defaultdict
import warnings

warnings.filterwarnings("ignore", category=UserWarning)
pd.set_option("future.no_silent_downcasting", True)

# --------- TEST_DATE (テスト時にここを編集) ----------
TEST_DATE = None
# TEST_DATE = "2026-03-22"
# TEST_DATE = "2026-04-10"
# ------------------------------------------------------

def get_today():
    if TEST_DATE:
        try:
            return datetime.strptime(TEST_DATE, "%Y-%m-%d")
        except Exception:
            pass
    return datetime.now()

# -------------------------
# Excel シート読み込み（1シート）
# -------------------------
def load_sheet(file_path, sheet_name, grade):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    except Exception as e:
        print(f"⚠ シート読み込みエラー：{file_path} / {sheet_name} → {e}")
        return []

    df = df[pd.to_datetime(df.iloc[:, 1], format="%Y-%m-%d", errors="coerce").notnull()]
    df = df.fillna("").infer_objects(copy=False)

    timetable = []
    for _, row in df.iterrows():
        date = pd.to_datetime(row.iloc[1]).strftime("%Y-%m-%d")
        comment = row.iloc[13] if len(row) > 13 else ""

        if comment:
            timetable.append({
                "grade": grade,
                "date": date,
                "period": 0,
                "courses": "",
                "room": "",
                "comment": comment
            })

        for p in range(1, 6):
            col_c = p * 2 + 1
            col_r = p * 2 + 2
            if col_r >= len(row):
                continue
            cname = row.iloc[col_c]
            room = row.iloc[col_r]
            if not cname:
                continue
            timetable.append({
                "grade": grade,
                "date": date,
                "period": p,
                "courses": cname,
                "room": room,
                "comment": ""
            })
    return timetable

# -------------------------
# 指定ファイル内のシート（全部）ロード
# -------------------------
def load_year_term(file_path, sheet_map):
    """
    sheet_map : { Excelシート名 : grade名 }
    """
    result = []
    for sheet, grade in sheet_map.items():
        part = load_sheet(file_path, sheet, grade)
        result.extend(part)
    return result

# -------------------------
# 助産統合（4年→4年助産補完）
# -------------------------
def add_schedule_to_josan(timetable):
    f4 = {}
    josan = {}
    for c in timetable:
        key = c["date"] + str(c["period"])
        if c["grade"] == "4年生":
            f4[key] = c
        elif c["grade"] == "4年助産":
            josan[key] = c
    for key, c in f4.items():
        if key in josan:
            continue
        new_c = c.copy()
        new_c["grade"] = "4年助産"
        timetable.append(new_c)
    return timetable

# -------------------------
# JSON / info 保存
# -------------------------
def save_json(data, path):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"✓ {path} を生成（{len(data)} 件）")

def save_info_json(file_path, out_path):
    if not os.path.exists(file_path):
        print(f"⚠ {file_path} がありません（空の info を作成）")
        save_json({"file_path": file_path, "last_modified": None}, out_path)
        return
    jst = timezone(timedelta(hours=9))
    ts = datetime.fromtimestamp(os.stat(file_path).st_mtime, tz=jst)
    save_json({"file_path": file_path, "last_modified": ts.strftime("%Y-%m-%d %H:%M:%S")}, out_path)

# -------------------------
# シート名マップ
# -------------------------
def sheet_names_for_year(year, term):
    yyyy = f"{year}年度"
    if term == "前期":
        return {
            f"{yyyy}(1年前期)": "1年生",
            f"{yyyy}(2年前期)": "2年生",
            f"{yyyy}(3年前期)": "3年生",
            f"{yyyy}(4年前期)": "4年生",
            f"{yyyy}(助産前期)": "4年助産",
            f"{yyyy}(M1前期)": "M1",
            f"{yyyy}(M2前期)": "M2",
            f"{yyyy}(D1前期)": "D1",
            f"{yyyy}(D23前期)": "D2/D3",
        }
    else:
        return {
            f"{yyyy}(1年後期)": "1年生",
            f"{yyyy}(2年後期)": "2年生",
            f"{yyyy}(3年後期)": "3年生",
            f"{yyyy}(4年後期)": "4年生",
            f"{yyyy}(助産後期)": "4年助産",
            f"{yyyy}(M1後期)": "M1",
            f"{yyyy}(M2後期)": "M2",
            f"{yyyy}(D1後期)": "D1",
            f"{yyyy}(D23後期)": "D2/D3",
        }

# -------------------------
# ファイルに対象シートが存在するか確認
# -------------------------
def excel_has_any_sheet(file_path, sheet_map):
    if not os.path.exists(file_path):
        return False
    try:
        x = pd.ExcelFile(file_path)
        names = x.sheet_names
        return any(s in names for s in sheet_map.keys())
    except Exception:
        return False

# -------------------------
# 年度候補から最適な年度を探す
# -------------------------
def find_year_for_file(file_path, candidates, term):
    for y in candidates:
        smap = sheet_names_for_year(y, term)
        if excel_has_any_sheet(file_path, smap):
            return y
    return None

# -------------------------
# 自動年度フォールバック読み込み
# -------------------------
def load_for_term_with_fallback(file_path, preferred_year, term):
    candidates = [preferred_year, preferred_year - 1]
    found_year = find_year_for_file(file_path, candidates, term)
    if found_year is None:
        print(f"⚠ {file_path} に {preferred_year}/{preferred_year-1} の {term} シートが見つかりません（スキップ）")
        return []
    smap = sheet_names_for_year(found_year, term)
    print(f"→ {file_path} : 使用する{term}年度 = {found_year}年度")
    return load_year_term(file_path, smap)

# -------------------------
# 後期 → 3/21〜31 のみ抽出
# -------------------------
def filter_last_march_only(timetable, year):
    result = []
    start = datetime(year, 3, 21)
    end = datetime(year, 3, 31)
    for c in timetable:
        try:
            d = datetime.strptime(c["date"], "%Y-%m-%d")
        except Exception:
            continue
        if start <= d <= end:
            result.append(c)
    return result

# -------------------------
# メイン
# -------------------------
if __name__ == "__main__":
    today = get_today()
    print(f"◆ Today: {today.strftime('%Y-%m-%d')}", end=" / ")

    if today.month <= 3:
        current_year = today.year - 1
    else:
        current_year = today.year
    next_year = current_year + 1

    if today.month == 3 and 21 <= today.day <= 31:
        mode = "mix"
    else:
        mode = "current_only"

    print(f"Mode: {mode}")
    print(f"◆ current_year={current_year}, next_year={next_year}")

    SPRING_CUR = "schedule_spring_CURRENT.xlsx"
    FALL_CUR = "schedule_fall_CURRENT.xlsx"
    SPRING_NEXT = "schedule_spring_NEXT.xlsx"
    FALL_NEXT = "schedule_fall_NEXT.xlsx"

    all_timetable = []

    if mode == "current_only":
        print("■ current_only（当年度 前期＋後期）")
        all_timetable.extend(load_for_term_with_fallback(SPRING_CUR, current_year, "前期"))
        all_timetable.extend(load_for_term_with_fallback(FALL_CUR, current_year, "後期"))

    elif mode == "mix":
        print("■ mix（CURRENT 後期 3/21〜31 + NEXT 前期 + NEXT 後期）")
        curr_fall_all = load_for_term_with_fallback(FALL_CUR, current_year, "後期")
        curr_march = filter_last_march_only(curr_fall_all, current_year)
        all_timetable.extend(curr_march)
        all_timetable.extend(load_for_term_with_fallback(SPRING_NEXT, next_year, "前期"))
        all_timetable.extend(load_for_term_with_fallback(FALL_NEXT, next_year, "後期"))

    all_timetable = add_schedule_to_josan(all_timetable)

    save_json(all_timetable, "docs/schedule.json")
    save_info_json(SPRING_CUR, "docs/info_spring_current.json")
    save_info_json(FALL_CUR, "docs/info_fall_current.json")
    save_info_json(SPRING_NEXT, "docs/info_spring_next.json")
    save_info_json(FALL_NEXT, "docs/info_fall_next.json")
