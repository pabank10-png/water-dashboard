"""
process_data.py
อ่าน Input Data RY.xlsx (Inflow) + ผันน้ำฯ xlsx (Outflow/Level)
→ จัดเป็น long-format Excel + JSON

แหล่งข้อมูล:
  - Inflow      : Input Data RY.xlsx (Inflow sheet) — ทุกอ่าง ทุกปี
  - Outflow/Level: 1.ผันน้ำและการใช้น้ำ ระยอง(บน).xlsx — DK, KY, NPL ถึง ส.ค. 2563
  - PS          : รอดึงจาก API (ยังไม่มีข้อมูล)
"""

import csv
import json
import os
import openpyxl
import pandas as pd
from datetime import datetime

_DIR             = os.path.dirname(os.path.abspath(__file__))

# ── รองรับทั้ง local (ไฟล์อยู่ใน Reservoir/) และ CI (ไฟล์อยู่ใน root ของ repo) ──
def _find_file(*candidates):
    """คืนไฟล์แรกที่มีอยู่จริง หรือ candidates[0] ถ้าไม่พบเลย"""
    for p in candidates:
        if os.path.exists(p):
            return p
    return candidates[0]

INPUT_FILE       = _find_file(
    os.path.join(_DIR, "..", "Input Data RY.xlsx"),   # local layout
    os.path.join(_DIR, "Input Data RY.xlsx"),          # CI flat layout
)
RAW_DAILY_FILE   = os.path.join(_DIR, "1.ผันน้ำและการใช้น้ำ ระยอง(บน).xlsx")
EXCEL_HIST_CSV   = _find_file(
    os.path.join(_DIR, "excel_historical.csv"),
    os.path.join(_DIR, "..", "excel_historical.csv"),
)
API_HIST_CSV     = os.path.join(_DIR, "api_historical_raw.csv")
CURRENT_YEAR_CSV = _find_file(
    os.path.join(_DIR, "../reservoir_data.csv"),       # local layout
    os.path.join(_DIR, "reservoir_data.csv"),           # CI flat layout
)
OUTPUT_XLSX      = os.path.join(_DIR, "water_data.xlsx")
OUTPUT_JSON      = os.path.join(_DIR, "data.json")

# ตัดข้อมูลดิบที่เดือน 8 ปี 2563 (หลังจากนั้นใช้ API)
DAILY_CUTOFF_YEAR  = 2020
DAILY_CUTOFF_MONTH = 8

# map API ID → reservoir key
API_ID_TO_RES = {
    "rsv357": "DK",
    "rsv359": "KY",
    "100504": "NPL",
    "100505": "PS",
}

MONTHS_TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.']

# ── ตำแหน่งข้อมูลใน Input Data RY.xlsx (Inflow sheet) ──
SECTIONS = {
    'DK':  {'data_start': 4,  'data_end': 33},
    'KY':  {'data_start': 35, 'data_end': 57},
    'NPL': {'data_start': 59, 'data_end': 96},
    'PS':  {'data_start': 98, 'data_end': 118},
}

# ── ตำแหน่งข้อมูลใน ผันน้ำฯ xlsx (รายวัน) ──
# date_col : index ของคอลัมน์วันที่
# vol_col  : index ของ ปริมาณน้ำ (ลบ.ม.) → level_end
# out_col  : index ของ น้ำลงสุทธิ/น้ำลงสุทธิ2 (ลบ.ม./วัน) → outflow
# data_start: แถวแรกที่มีข้อมูล (0-based index)
DAILY_CONFIGS = {
    'NPL': {'sheet': 'หนองปลาไหล', 'date_col': 0, 'vol_col': 3, 'out_col': 15, 'data_start': 4},
    'DK':  {'sheet': 'ดอกกราย',    'date_col': 0, 'vol_col': 4, 'out_col': 14, 'data_start': 3},
    'KY':  {'sheet': 'คลองใหญ่',   'date_col': 1, 'vol_col': 5, 'out_col': 13, 'data_start': 3},
}


# ── อ่าน Inflow จาก Input Data RY.xlsx ──
def read_inflow(ws, sections):
    records = []
    for reservoir, sec in sections.items():
        for row in ws.iter_rows(
            min_row=sec['data_start'],
            max_row=sec['data_end'],
            values_only=True
        ):
            year = row[1]
            if not year or not isinstance(year, (int, float)):
                continue
            year = int(year)
            for m_idx, col_offset in enumerate(range(2, 14)):  # columns C–N
                val = row[col_offset]
                records.append({
                    'reservoir': reservoir,
                    'year':      year,
                    'month':     m_idx + 1,
                    'inflow':    round(float(val), 4) if isinstance(val, (int, float)) else None,
                    'outflow':   None,
                    'level_end': None,
                })
    return records


# ── อ่าน Outflow และ Level จากไฟล์ดิบรายวัน ──
def read_outflow_level(xlsx_path,
                       cutoff_year=DAILY_CUTOFF_YEAR,
                       cutoff_month=DAILY_CUTOFF_MONTH):
    """
    อ่านข้อมูลรายวัน DK/KY/NPL จาก xlsx
    รวมต่อเดือน:
      - outflow   = ผลรวมรายวัน (ลบ.ม.) / 1e6 → MCM
      - level_end = ปริมาณน้ำวันสุดท้ายของเดือน / 1e6 → MCM
    คืนค่า dict: {res: {(thai_year_int, month): (outflow_mcm, level_end_mcm)}}
    """
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    result = {}

    for res, cfg in DAILY_CONFIGS.items():
        ws = wb[cfg['sheet']]
        rows = list(ws.iter_rows(values_only=True))
        data_rows = rows[cfg['data_start']:]

        # สะสมรายวันแต่ละเดือน
        monthly = {}  # (thai_year, month) → {'vols': [], 'outs': []}

        for row in data_rows:
            date = row[cfg['date_col']]
            if not isinstance(date, datetime):
                continue

            # ตัดข้อมูลหลัง cutoff
            if (date.year, date.month) > (cutoff_year, cutoff_month):
                continue

            thai_year = date.year + 543
            month = date.month
            key = (thai_year, month)

            if key not in monthly:
                monthly[key] = {'vols': [], 'outs': []}

            vol = row[cfg['vol_col']]
            out = row[cfg['out_col']]

            if isinstance(vol, (int, float)):
                monthly[key]['vols'].append(vol)
            if isinstance(out, (int, float)):
                monthly[key]['outs'].append(out)

        # รวมเป็นรายเดือน
        res_data = {}
        for (thai_year, month), vals in monthly.items():
            level_end = round(vals['vols'][-1] / 1e6, 4) if vals['vols'] else None
            outflow   = round(sum(vals['outs']) / 1e6, 4) if vals['outs'] else None
            res_data[(thai_year, month)] = (outflow, level_end)

        result[res] = res_data
        print(f"  {res}: {len(res_data)} เดือน จากไฟล์ดิบ")

    return result


# ── อ่าน Outflow/Level จาก API CSV (api_historical_raw.csv + reservoir_data.csv) ──
def read_api_outflow_level(*csv_paths):
    """
    อ่านไฟล์ CSV ที่ได้จาก API (รูปแบบเดียวกับ reservoir_data.csv)
    รวมข้อมูลรายวัน → รายเดือน
      - outflow   : ผลรวมรายวัน (MCM/day sum → MCM/month)
      - level_end : ปริมาณน้ำวันสุดท้ายของเดือน (MCM)
    คืน dict: {res: {(thai_year, month): (outflow, level_end)}}
    """
    # รวมแถวจากทุกไฟล์ dedup ด้วย (date_record, id)
    rows_by_key = {}
    for path in csv_paths:
        if not os.path.exists(path):
            print(f"  ⚠ ไม่พบ {path} — ข้าม")
            continue
        with open(path, encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                key = (row.get("date_record", ""), row.get("id", ""))
                rows_by_key[key] = row

    # จัดกลุ่มเป็น {(res, thai_year, month): {'vols': [], 'outs': [], 'infs': []}}
    monthly = {}
    for (date_str, api_id), row in rows_by_key.items():
        res = API_ID_TO_RES.get(api_id)
        if res is None:
            continue
        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            continue
        thai_year = dt.year + 543
        month = dt.month
        k = (res, thai_year, month)
        if k not in monthly:
            monthly[k] = {"vols": [], "outs": [], "infs": []}
        try:
            vol = float(row["volume"]) if row.get("volume") not in (None, "") else None
            out = float(row["outflow"]) if row.get("outflow") not in (None, "") else None
            inf = float(row["inflow"])  if row.get("inflow")  not in (None, "") else None
        except (TypeError, ValueError):
            vol, out, inf = None, None, None
        if vol is not None:
            monthly[k]["vols"].append((dt, vol))
        if out is not None:
            monthly[k]["outs"].append(out)
        if inf is not None:
            monthly[k]["infs"].append(inf)

    result = {}   # {res: {(thai_year, month): (inflow, outflow, level_end)}}
    for (res, thai_year, month), vals in monthly.items():
        level_end = None
        if vals["vols"]:
            _, last_vol = max(vals["vols"], key=lambda x: x[0])
            level_end = round(last_vol, 4)
        outflow = round(sum(vals["outs"]), 4) if vals["outs"] else None
        inflow  = round(sum(vals["infs"]), 4) if vals["infs"] else None

        if res not in result:
            result[res] = {}
        result[res][(thai_year, month)] = (inflow, outflow, level_end)

    total_months = sum(len(v) for v in result.values())
    print(f"  API CSV: {total_months} เดือน จาก {len(rows_by_key)} แถวรายวัน")
    return result


# ── อ่าน Outflow/Level จาก excel_historical.csv (fallback เมื่อไม่มีไฟล์ Excel ดิบ) ──
def read_outflow_level_csv(csv_path):
    """
    อ่าน excel_historical.csv ที่ export ไว้แล้ว
    คืน dict เหมือน read_outflow_level: {res: {(thai_year, month): (outflow, level_end)}}
    """
    result = {}
    if not os.path.exists(csv_path):
        print(f"  ⚠ ไม่พบ {csv_path}")
        return result
    with open(csv_path, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            res   = row['reservoir']
            yr    = int(row['thai_year'])
            month = int(row['month'])
            outflow   = float(row['outflow'])   if row.get('outflow')   not in (None, '') else None
            level_end = float(row['level_end']) if row.get('level_end') not in (None, '') else None
            if res not in result:
                result[res] = {}
            result[res][(yr, month)] = (outflow, level_end)
    total = sum(len(v) for v in result.values())
    print(f"  CSV fallback: {total} เดือน")
    return result


# ── คำนวณ THREE = DK + KY + NPL รายเดือน ──
def calc_three(df):
    """
    รวม DK+KY+NPL เป็น reservoir 'THREE'
    แต่ละ metric จะ None ถ้าไม่ครบทั้ง 3 อ่างในเดือนนั้น
    """
    three_res = ['DK', 'KY', 'NPL']
    sub = df[df['reservoir'].isin(three_res)].copy()

    rows = []
    for (year, month), grp in sub.groupby(['year', 'month']):
        row = {'reservoir': 'THREE', 'year': year, 'month': month}
        for metric in ['inflow', 'outflow', 'level_end']:
            vals = grp[metric].dropna()
            row[metric] = round(vals.sum(), 4) if len(vals) == 3 else None
        rows.append(row)

    return pd.DataFrame(rows, columns=['reservoir','year','month','inflow','outflow','level_end'])


# ── เขียน Excel output ──
def write_excel(df, path):
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        # Sheet master: long-format ทุก reservoir
        df.to_excel(writer, sheet_name='master', index=False)

        for res in ['DK', 'KY', 'NPL', 'THREE', 'PS']:
            sub = df[df['reservoir'] == res][['year','month','inflow','outflow','level_end']]

            for metric, label in [('inflow', 'inflow'), ('outflow', 'outflow'), ('level_end', 'level')]:
                if sub[metric].notna().any():
                    wide = sub.pivot(index='year', columns='month', values=metric)
                    wide.columns = MONTHS_TH
                    wide.index.name = 'year (Thai)'
                    if metric != 'level_end':
                        wide['รวม'] = wide.sum(axis=1)
                    wide.to_excel(writer, sheet_name=f'{label}_{res}')

    print(f"✓ เขียน {path} เสร็จ")


# ── เขียน JSON ──
def write_json(df, path):
    reservoirs = ['DK', 'KY', 'NPL', 'THREE', 'PS']
    data_out = {
        metric: {res: {} for res in reservoirs}
        for metric in ['inflow', 'outflow', 'level_end']
    }

    for res in reservoirs:
        sub = df[df['reservoir'] == res]
        for year, grp in sub.groupby('year'):
            grp = grp.sort_values('month')
            yy = year % 100
            key = f'y{yy:02d}'
            for metric in ['inflow', 'outflow', 'level_end']:
                # สร้าง array 12 ช่อง — เดือนที่ไม่มีข้อมูลให้เป็น None
                monthly = [None] * 12
                for _, row in grp.iterrows():
                    m_idx = int(row['month']) - 1
                    v = row[metric]
                    if pd.notna(v):
                        monthly[m_idx] = round(float(v), 4)
                if any(v is not None for v in monthly):
                    data_out[metric][res][key] = monthly

    output = {
        'generated':      datetime.now().strftime('%Y-%m-%d %H:%M'),
        'schema_version': 1,
        'inflow':    data_out['inflow'],
        'outflow':   data_out['outflow'],
        'level_end': data_out['level_end'],
    }

    with open(path, 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    print(f"✓ เขียน {path} เสร็จ")


# ── main ──
if __name__ == '__main__':
    # 1. อ่าน Inflow จาก Input Data RY.xlsx
    wb_inf = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    ws_inf = wb_inf['Inflow']
    records = read_inflow(ws_inf, SECTIONS)
    print(f"✓ อ่าน Inflow {len(records)} แถว จาก 4 อ่าง")

    df = pd.DataFrame(records)
    df = df.sort_values(['reservoir','year','month']).reset_index(drop=True)

    # 2. อ่าน Outflow/Level จากไฟล์ดิบรายวัน หรือ CSV fallback
    if os.path.exists(RAW_DAILY_FILE):
        print(f"\nอ่าน Outflow/Level จาก {RAW_DAILY_FILE}:")
        outflow_level_excel = read_outflow_level(RAW_DAILY_FILE)
    else:
        print(f"\nไม่พบ Excel ดิบ → ใช้ {EXCEL_HIST_CSV}:")
        outflow_level_excel = read_outflow_level_csv(EXCEL_HIST_CSV)

    # เติมข้อมูล Excel ก่อน (ลำดับความสำคัญสูงกว่า API)
    def apply_outflow_level(df, data_dict):
        for res, data in data_dict.items():
            for (thai_year, month), (outflow, level_end) in data.items():
                mask = (df['reservoir'] == res) & (df['year'] == thai_year) & (df['month'] == month)
                if mask.any():
                    if outflow is not None:
                        df.loc[mask, 'outflow']   = outflow
                    if level_end is not None:
                        df.loc[mask, 'level_end'] = level_end

    apply_outflow_level(df, outflow_level_excel)

    # 2b. เติมส่วนที่ยังขาดจาก API CSV (api_historical_raw.csv + reservoir_data.csv)
    print(f"\nอ่าน Outflow/Level จาก API CSV:")
    outflow_level_api = read_api_outflow_level(API_HIST_CSV, CURRENT_YEAR_CSV)

    # เติมเฉพาะ cell ที่ยังเป็น None (ไม่ทับ Excel)
    # ถ้าปี/เดือนนั้นยังไม่มีใน df (เช่น 2569) ให้เพิ่มแถวใหม่
    new_api_rows = []
    for res, data in outflow_level_api.items():
        for (thai_year, month), (inflow_api, outflow, level_end) in data.items():
            mask = (df['reservoir'] == res) & (df['year'] == thai_year) & (df['month'] == month)
            if mask.any():
                if inflow_api is not None and df.loc[mask, 'inflow'].isna().all():
                    df.loc[mask, 'inflow'] = inflow_api
                if outflow is not None and df.loc[mask, 'outflow'].isna().all():
                    df.loc[mask, 'outflow'] = outflow
                if level_end is not None and df.loc[mask, 'level_end'].isna().all():
                    df.loc[mask, 'level_end'] = level_end
            else:
                # ปีที่ยังไม่มีในชุดข้อมูล inflow (เช่น 2569 = 2026)
                new_api_rows.append({
                    'reservoir': res, 'year': thai_year, 'month': month,
                    'inflow': inflow_api, 'outflow': outflow, 'level_end': level_end,
                })
    if new_api_rows:
        df = pd.concat([df, pd.DataFrame(new_api_rows)], ignore_index=True)
        df = df.sort_values(['reservoir', 'year', 'month']).reset_index(drop=True)

    # 3. คำนวณ THREE
    three_df = calc_three(df)
    df = pd.concat([df, three_df], ignore_index=True)
    df = df.sort_values(['reservoir','year','month']).reset_index(drop=True)

    # 4. เขียน output
    print()
    write_excel(df, OUTPUT_XLSX)
    write_json(df, OUTPUT_JSON)

    # สรุป
    print()
    print("── สรุปข้อมูล ──")
    for res in ['DK', 'KY', 'NPL', 'THREE', 'PS']:
        sub = df[df['reservoir'] == res]
        inf_rows = sub[sub['inflow'].notna()]
        out_rows = sub[sub['outflow'].notna()]
        lev_rows = sub[sub['level_end'].notna()]
        years_inf = f"{inf_rows['year'].min()}–{inf_rows['year'].max()}" if not inf_rows.empty else "–"
        years_out = f"{out_rows['year'].min()}–{out_rows['year'].max()}" if not out_rows.empty else "–"
        years_lev = f"{lev_rows['year'].min()}–{lev_rows['year'].max()}" if not lev_rows.empty else "–"
        print(f"  {res:5s}: inflow={years_inf} | outflow={years_out} | level_end={years_lev}")
