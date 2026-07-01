"""
process_data.py  (v3 — Input Data RY.xlsx เป็น primary source สำหรับ inflow)

Flow:
  1. อ่าน Input Data RY.xlsx  → inflow (DK, KY, NPL, PS) — primary truth
  2. อ่าน water_data.xlsx master  → outflow + level (+ inflow ที่ไม่มีใน Input Data)
  3. Override master inflow ด้วยค่าจาก Input Data RY.xlsx ทุก row ที่มี
  4. อ่าน api_historical_raw.csv  → เติมช่องที่ยังว่าง + เดือนปัจจุบัน (2569)
  5. คำนวณ THREE (DK+KY+NPL) ใหม่
  6. เขียน water_data.xlsx  (อัปเดต in-place)
  7. เขียน data.json  → dashboard
"""

import csv
import json
import os
from datetime import datetime

import openpyxl
import pandas as pd

_DIR          = os.path.dirname(os.path.abspath(__file__))

# ── ไฟล์ I/O ──
INPUT_DATA_XL = os.path.join(_DIR, "Input Data RY.xlsx")   # primary inflow source
MASTER_XLSX   = os.path.join(_DIR, "water_data.xlsx")
API_HIST_CSV  = os.path.join(_DIR, "api_historical_raw.csv")
OUTPUT_JSON   = os.path.join(_DIR, "data.json")
# copy ไปที่ Water folder ด้วย (สำหรับ water_dashboard.html ที่เปิดใน local)
WATER_DIR     = os.path.normpath(os.path.join(_DIR, "..", "..", "..", "..", "Water"))
OUTPUT_JSON2  = os.path.join(WATER_DIR, "data.json") if os.path.isdir(WATER_DIR) else None

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


# ── 1b. อ่าน inflow จาก Input Data RY.xlsx (primary truth) ──
def read_input_data_ry(path):
    """
    อ่าน sheet 'Inflow' ใน Input Data RY.xlsx
    คืน DataFrame: reservoir, year (Thai), month (1–12), inflow
    """
    if not os.path.exists(path):
        print(f"  ⚠ ไม่พบ {path}")
        return pd.DataFrame(columns=["reservoir","year","month","inflow"])

    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["Inflow"]
    rows = list(ws.iter_rows(values_only=True))

    # หา header row ของแต่ละ reservoir (ค่า col[1] เป็น 'DK','KY','NPL','PS')
    header_indices = [i for i, r in enumerate(rows) if r[1] in ("DK","KY","NPL","PS")]

    records = []
    for hi in header_indices:
        res_name = rows[hi][1]
        for r in rows[hi + 1:]:
            if r[1] is None:
                break
            if not isinstance(r[1], (int, float)):
                continue
            yr = int(r[1])
            for mo_idx, val in enumerate(r[2:14], start=1):
                if isinstance(val, (int, float)):
                    records.append({
                        "reservoir": res_name,
                        "year":      yr,
                        "month":     mo_idx,
                        "inflow":    round(float(val), 4),
                    })

    df = pd.DataFrame(records)
    res_list = df["reservoir"].unique().tolist() if not df.empty else []
    yr_range = f"{df['year'].min()}–{df['year'].max()}" if not df.empty else "–"
    print(f"✓ อ่าน Input Data RY.xlsx: {len(df)} เดือน  reservoir={res_list}  ปี={yr_range}")
    return df


def override_inflow(df_master, df_input):
    """
    ใช้ Input Data RY.xlsx เป็น primary source สำหรับ inflow ของ DK, KY, NPL, PS
    - ล้าง inflow ของ DK/KY/NPL/PS ในปีก่อน 2569 ทั้งหมดก่อน (ป้องกัน API data เก่าปนเข้ามา)
    - ใส่ค่าจาก Input Data RY.xlsx กลับเข้าไป (ทั้ง row ที่มีอยู่แล้วและ row ใหม่)
    """
    if df_input.empty:
        return df_master

    # ล้าง inflow เฉพาะปีก่อน 2569 สำหรับ reservoir ที่มีใน Input Data
    INPUT_RESERVOIRS = df_input["reservoir"].unique().tolist()
    CURRENT_YEAR = datetime.now().year + 543   # ปีปัจจุบัน Thai
    hist_mask = (
        df_master["reservoir"].isin(INPUT_RESERVOIRS) &
        (df_master["year"] < CURRENT_YEAR)
    )
    cleared = hist_mask.sum()
    df_master.loc[hist_mask, "inflow"] = None

    master_idx = {
        (r, y, m): i
        for i, (r, y, m) in enumerate(
            zip(df_master["reservoir"], df_master["year"], df_master["month"])
        )
    }

    new_rows = []
    updated = 0
    for _, row in df_input.iterrows():
        key = (row["reservoir"], int(row["year"]), int(row["month"]))
        if key in master_idx:
            df_master.at[master_idx[key], "inflow"] = row["inflow"]
            updated += 1
        else:
            new_rows.append({
                "reservoir": row["reservoir"],
                "year":      int(row["year"]),
                "month":     int(row["month"]),
                "inflow":    row["inflow"],
                "outflow":   None,
                "level_end": None,
            })

    if new_rows:
        df_master = pd.concat([df_master, pd.DataFrame(new_rows)], ignore_index=True)

    df_master = df_master.sort_values(["reservoir","year","month"]).reset_index(drop=True)
    print(f"✓ override inflow: ล้าง {cleared} ค่าเก่า  อัปเดต {updated}  เพิ่ม {len(new_rows)} row ใหม่")
    return df_master


# ── 2. อ่าน API CSV → aggregate รายเดือน ──
def _remove_outliers(vals):
    """
    กรอง daily outflow ที่ผิดปกติ (เช่น ทศนิยมหาย: 328 แทน 0.328)
    ถ้าค่าใดสูงกว่า median ของค่าที่เหลือ > 50 เท่า ให้ตัดทิ้ง
    """
    if len(vals) <= 2:
        return vals
    non_zero = [v for v in vals if v > 0]
    if not non_zero:
        return vals
    sorted_nz = sorted(non_zero)
    median = sorted_nz[len(sorted_nz) // 2]
    if median <= 0:
        return vals
    return [v for v in vals if v <= median * 50]

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
    latest_vol_date = {}  # reservoir → latest date string "YYYY-MM-DD"
    for (res, thai_year, month), v in monthly.items():
        level_end = None
        if v["vols"]:
            latest_dt, lv = max(v["vols"], key=lambda x: x[0])
            level_end = round(lv, 4)
            ds = latest_dt.strftime("%Y-%m-%d")
            if res not in latest_vol_date or ds > latest_vol_date[res]:
                latest_vol_date[res] = ds
        outflow = round(sum(_remove_outliers(v["outs"])), 4) if v["outs"] else None
        inflow  = round(sum(_remove_outliers(v["infs"])), 4) if v["infs"] else None
        records.append({"reservoir": res, "year": thai_year, "month": month,
                         "inflow": inflow, "outflow": outflow, "level_end": level_end})

    df = pd.DataFrame(records)
    print(f"✓ อ่าน API CSV: {len(df)} เดือน")
    return df, latest_vol_date


# ── 3. Merge API เข้า master ──
def _overwrite_months(lookback_days=5):
    """
    คืน set ของ (thai_year, month) ที่ควร overwrite
    = เดือนที่มีวันอยู่ใน lookback window (today - lookback_days)
    ตัวอย่าง: วันที่ 20 มิ.ย.  → window = 15–20 มิ.ย.  → {(2569, 6)}
              วันที่  3 ก.ค.  → window = 28 มิ.ย.–3 ก.ค. → {(2569, 6), (2569, 7)}
    """
    import calendar
    from datetime import date, timedelta
    today        = date.today()
    window_start = today - timedelta(days=lookback_days)
    months = set()
    d = window_start
    while d <= today:
        months.add((d.year + 543, d.month))
        # เลื่อนไปวันแรกของเดือนถัดไปเพื่อ loop เร็ว
        last = calendar.monthrange(d.year, d.month)[1]
        d = date(d.year, d.month, last) + timedelta(days=1)
    return months

def merge_api(df_master, df_api):
    """
    กฎ:
      • เดือนที่มีวันอยู่ใน LOOKBACK_DAYS → overwrite (ข้อมูล API อาจ update ช้า)
      • เดือนอื่นทั้งหมด → เพิ่มเฉพาะ row ใหม่ที่ยังไม่มี (ไม่แตะข้อมูลที่มีอยู่)
      • เดือนใหม่ที่ยังไม่มีใน master → เพิ่ม row ใหม่
    """
    if df_api.empty:
        return df_master

    overwrite = _overwrite_months(lookback_days=5)   # set ของ (thai_year, month)

    master_idx = {(r, y, m): i
                  for i, (r, y, m) in enumerate(
                      zip(df_master["reservoir"], df_master["year"], df_master["month"]))}

    new_rows = []
    updated = 0

    current_thai_year = datetime.now().year + 543
    # reservoir ที่ Input Data RY.xlsx เป็น primary source สำหรับ inflow
    INPUT_DATA_RES = {"DK", "KY", "NPL", "PS"}

    for _, api_row in df_api.iterrows():
        res   = api_row["reservoir"]
        year  = int(api_row["year"])
        month = int(api_row["month"])
        key   = (res, year, month)
        is_recent = (year, month) in overwrite

        if key in master_idx:
            idx = master_idx[key]
            for col in ["inflow", "outflow", "level_end"]:
                # inflow ของ DK/KY/NPL/PS ปีก่อน 2569 → ห้ามเติมจาก API
                # (Input Data RY.xlsx เป็น source เดียว)
                if col == "inflow" and res in INPUT_DATA_RES and year < current_thai_year:
                    continue
                new_val = api_row[col]
                if pd.notna(new_val):
                    if is_recent or pd.isna(df_master.at[idx, col]):
                        df_master.at[idx, col] = round(float(new_val), 4)
                        updated += 1
        else:
            # row ใหม่ที่ไม่มีใน master เลย
            inf_val = None
            if not (res in INPUT_DATA_RES and year < current_thai_year):
                # เพิ่ม inflow เฉพาะถ้าไม่ใช่ historical ของ reservoir ใน Input Data
                inf_val = api_row["inflow"] if pd.notna(api_row["inflow"]) else None
            new_rows.append({
                "reservoir": res, "year": year, "month": month,
                "inflow":    inf_val,
                "outflow":   api_row["outflow"]   if pd.notna(api_row["outflow"])   else None,
                "level_end": api_row["level_end"] if pd.notna(api_row["level_end"]) else None,
            })

    if new_rows:
        df_master = pd.concat([df_master, pd.DataFrame(new_rows)], ignore_index=True)

    df_master = df_master.sort_values(["reservoir","year","month"]).reset_index(drop=True)
    print(f"✓ merge: อัปเดต {updated} ค่า  เพิ่ม {len(new_rows)} เดือนใหม่")
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
def write_json(df, path, latest_vol_date=None):
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

    _ld = dict(latest_vol_date) if latest_vol_date else {}
    three_dates = [_ld.get(r) for r in ["DK","KY","NPL"] if _ld.get(r)]
    if three_dates:
        _ld["THREE"] = min(three_dates)
    payload = {
        "generated":      datetime.now().strftime("%Y-%m-%d %H:%M"),
        "schema_version": 1,
        "inflow":    out["inflow"],
        "outflow":   out["outflow"],
        "level_end": out["level_end"],
        "latest_date": _ld,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"✓ เขียน {path}")


# ── main ──
if __name__ == "__main__":
    print("=" * 50)
    print("process_data.py  — Input Data RY.xlsx primary inflow")
    print("=" * 50)

    # 1. อ่าน master (outflow + level + inflow เก่า)
    df = read_master(MASTER_XLSX)
    df = df[df["reservoir"] != "THREE"]   # ลบ THREE เก่า → คำนวณใหม่

    # 1b. Override inflow ด้วย Input Data RY.xlsx (primary truth)
    print(f"\nอ่าน Input Data RY.xlsx:")
    df_input = read_input_data_ry(INPUT_DATA_XL)
    print(f"\nOverride inflow:")
    df = override_inflow(df, df_input)

    # 2. อ่าน API
    print(f"\nอ่าน API CSV:")
    df_api, latest_vol_date = read_api_monthly(API_HIST_CSV)

    # 3. Merge API — เติมเฉพาะช่องที่ยังว่าง + เดือนปัจจุบัน (2569)
    print(f"\nMerge API:")
    df = merge_api(df, df_api[df_api["reservoir"] != "THREE"])

    # 4. คำนวณ THREE ใหม่
    three_df = calc_three(df)
    df = pd.concat([df, three_df], ignore_index=True)
    df = df.sort_values(["reservoir","year","month"]).reset_index(drop=True)

    # 5 & 6. เขียน output
    print()
    write_excel(df, MASTER_XLSX)
    write_json(df, OUTPUT_JSON, latest_vol_date)
    if OUTPUT_JSON2:
        import shutil
        shutil.copy2(OUTPUT_JSON, OUTPUT_JSON2)
        print(f"✓ copy → {OUTPUT_JSON2}")

    # สรุป
    print("\n── สรุป ──")
    for res in ["DK","KY","NPL","THREE","PS"]:
        sub = df[df["reservoir"] == res]
        def rng(col):
            s = sub[sub[col].notna()]
            return f"{s['year'].min()}–{s['year'].max()}" if not s.empty else "–"
        print(f"  {res:5s}: inflow={rng('inflow')} | outflow={rng('outflow')} | level={rng('level_end')}")
