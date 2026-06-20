"""
process_data.py  (v2 — water_data.xlsx เป็น single source of truth)

Flow:
  1. อ่าน water_data.xlsx  → ข้อมูลทั้งหมดที่มีอยู่
  2. อ่าน api_historical_raw.csv  → ข้อมูล API รายวัน (ดึงโดย fetch_api_data.py)
  3. merge: เติมช่องที่ยังว่าง + เพิ่มเดือนใหม่
  4. คำนวณ THREE (DK+KY+NPL) ใหม่
  5. เขียน water_data.xlsx  (อัปเดต in-place)
  6. เขียน data.json  → dashboard
"""

import csv
import json
import os
from datetime import datetime

import openpyxl
import pandas as pd

_DIR          = os.path.dirname(os.path.abspath(__file__))

# ── ไฟล์ I/O ──
MASTER_XLSX   = os.path.join(_DIR, "water_data.xlsx")
API_HIST_CSV  = os.path.join(_DIR, "api_historical_raw.csv")
OUTPUT_JSON   = os.path.join(_DIR, "data.json")

# map API ID → reservoir key
API_ID_TO_RES = {
    "rsv357": "DK",
    "rsv359": "KY",
    "100504": "NPL",
    "100505": "PS",
}

MONTHS_TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.',
              'ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']


# ── 1. อ่าน master จาก water_data.xlsx ──
def read_master(path):
    df = pd.read_excel(path, sheet_name="master", dtype={"year": int, "month": int})
    df = df[df["reservoir"].notna()].copy()
    df["year"]  = df["year"].astype(int)
    df["month"] = df["month"].astype(int)
    print(f"✓ อ่าน master: {len(df)} แถว  ({df['year'].min()}–{df['year'].max()})")
    return df


# ── 2. อ่าน API CSV → aggregate รายเดือน ──
def read_api_monthly(csv_path):
    """
    คืน DataFrame: reservoir, year (Thai), month, inflow, outflow, level_end
    """
    if not os.path.exists(csv_path):
        print(f"  ⚠ ไม่พบ {csv_path}")
        return pd.DataFrame(columns=["reservoir","year","month","inflow","outflow","level_end"])

    rows_by_key = {}
    with open(csv_path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            k = (row.get("date_record",""), row.get("id",""))
            rows_by_key[k] = row

    monthly = {}
    for (date_str, api_id), row in rows_by_key.items():
        res = API_ID_TO_RES.get(api_id)
        if not res:
            continue
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            continue
        thai_year = dt.year + 543
        k = (res, thai_year, dt.month)
        if k not in monthly:
            monthly[k] = {"vols": [], "outs": [], "infs": []}
        try:
            vol = float(row["volume"])  if row.get("volume")  not in (None,"") else None
            out = float(row["outflow"]) if row.get("outflow") not in (None,"") else None
            inf = float(row["inflow"])  if row.get("inflow")  not in (None,"") else None
        except (TypeError, ValueError):
            vol = out = inf = None
        if vol is not None: monthly[k]["vols"].append((dt, vol))
        if out is not None: monthly[k]["outs"].append(out)
        if inf is not None: monthly[k]["infs"].append(inf)

    records = []
    for (res, thai_year, month), v in monthly.items():
        level_end = None
        if v["vols"]:
            _, lv = max(v["vols"], key=lambda x: x[0])
            level_end = round(lv, 4)
        outflow = round(sum(v["outs"]), 4) if v["outs"] else None
        inflow  = round(sum(v["infs"]), 4) if v["infs"] else None
        records.append({"reservoir": res, "year": thai_year, "month": month,
                         "inflow": inflow, "outflow": outflow, "level_end": level_end})

    df = pd.DataFrame(records)
    print(f"✓ อ่าน API CSV: {len(df)} เดือน")
    return df


# ── 3. Merge API เข้า master ──
def merge_api(df_master, df_api):
    """
    กฎ:
      • เดือนที่มีอยู่ใน master แล้ว → เติมเฉพาะช่องที่ยัง NaN
      • เดือนใหม่ที่ยังไม่มีใน master → เพิ่มแถวใหม่
    """
    if df_api.empty:
        return df_master

    # index master เพื่อ lookup เร็ว
    master_idx = {(r, y, m): i
                  for i, (r, y, m) in enumerate(
                      zip(df_master["reservoir"], df_master["year"], df_master["month"]))}

    new_rows = []
    updated = 0

    for _, api_row in df_api.iterrows():
        key = (api_row["reservoir"], int(api_row["year"]), int(api_row["month"]))
        if key in master_idx:
            idx = master_idx[key]
            for col in ["inflow", "outflow", "level_end"]:
                if pd.isna(df_master.at[idx, col]) and pd.notna(api_row[col]):
                    df_master.at[idx, col] = round(float(api_row[col]), 4)
                    updated += 1
        else:
            new_rows.append({
                "reservoir": api_row["reservoir"],
                "year":      int(api_row["year"]),
                "month":     int(api_row["month"]),
                "inflow":    api_row["inflow"]  if pd.notna(api_row["inflow"])  else None,
                "outflow":   api_row["outflow"] if pd.notna(api_row["outflow"]) else None,
                "level_end": api_row["level_end"] if pd.notna(api_row["level_end"]) else None,
            })

    if new_rows:
        df_master = pd.concat([df_master, pd.DataFrame(new_rows)], ignore_index=True)

    df_master = df_master.sort_values(["reservoir","year","month"]).reset_index(drop=True)
    print(f"✓ merge: อัปเดต {updated} ช่อง  เพิ่ม {len(new_rows)} เดือนใหม่")
    return df_master


# ── 4. คำนวณ THREE (DK+KY+NPL) ──
def calc_three(df):
    sub = df[df["reservoir"].isin(["DK","KY","NPL"])].copy()
    rows = []
    for (year, month), grp in sub.groupby(["year","month"]):
        row = {"reservoir": "THREE", "year": year, "month": month}
        for col in ["inflow","outflow","level_end"]:
            vals = grp[col].dropna()
            row[col] = round(float(vals.sum()), 4) if len(vals) == 3 else None
        rows.append(row)
    return pd.DataFrame(rows)


# ── 5. เขียน water_data.xlsx ──
def write_excel(df, path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="master", index=False)

        for res in ["DK","KY","NPL","THREE","PS"]:
            sub = df[df["reservoir"] == res][["year","month","inflow","outflow","level_end"]]
            for metric, label in [("inflow","inflow"),("outflow","outflow"),("level_end","level")]:
                if sub[metric].notna().any():
                    wide = sub.pivot(index="year", columns="month", values=metric)
                    wide.columns = MONTHS_TH[:len(wide.columns)]
                    wide.index.name = "year (Thai)"
                    if metric != "level_end":
                        wide["รวม"] = wide.sum(axis=1)
                    wide.to_excel(writer, sheet_name=f"{label}_{res}")

    print(f"✓ เขียน {path}")


# ── 6. เขียน data.json ──
def write_json(df, path):
    reservoirs = ["DK","KY","NPL","THREE","PS"]
    out = {m: {r: {} for r in reservoirs} for m in ["inflow","outflow","level_end"]}

    for res in reservoirs:
        sub = df[df["reservoir"] == res]
        for year, grp in sub.groupby("year"):
            grp = grp.sort_values("month")
            key = f"y{int(year) % 100:02d}"
            for metric in ["inflow","outflow","level_end"]:
                arr = [None] * 12
                for _, row in grp.iterrows():
                    m = int(row["month"]) - 1
                    v = row[metric]
                    if pd.notna(v):
                        arr[m] = round(float(v), 4)
                if any(v is not None for v in arr):
                    out[metric][res][key] = arr

    payload = {
        "generated":      datetime.now().strftime("%Y-%m-%d %H:%M"),
        "schema_version": 1,
        "inflow":    out["inflow"],
        "outflow":   out["outflow"],
        "level_end": out["level_end"],
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"✓ เขียน {path}")


# ── main ──
if __name__ == "__main__":
    print("=" * 50)
    print("process_data.py  — water_data.xlsx single source")
    print("=" * 50)

    # 1. อ่าน master
    df = read_master(MASTER_XLSX)

    # 2. อ่าน API
    print(f"\nอ่าน API CSV:")
    df_api = read_api_monthly(API_HIST_CSV)

    # 3. Merge
    print(f"\nMerge:")
    df = df[df["reservoir"] != "THREE"]   # ลบ THREE เก่า → คำนวณใหม่
    df = merge_api(df, df_api[df_api["reservoir"] != "THREE"])

    # 4. คำนวณ THREE ใหม่
    three_df = calc_three(df)
    df = pd.concat([df, three_df], ignore_index=True)
    df = df.sort_values(["reservoir","year","month"]).reset_index(drop=True)

    # 5 & 6. เขียน output
    print()
    write_excel(df, MASTER_XLSX)
    write_json(df, OUTPUT_JSON)

    # สรุป
    print("\n── สรุป ──")
    for res in ["DK","KY","NPL","THREE","PS"]:
        sub = df[df["reservoir"] == res]
        def rng(col):
            s = sub[sub[col].notna()]
            return f"{s['year'].min()}–{s['year'].max()}" if not s.empty else "–"
        print(f"  {res:5s}: inflow={rng('inflow')} | outflow={rng('outflow')} | level={rng('level_end')}")
