"""
fetch_api_data.py
ดึงข้อมูล volume / outflow รายวันจาก RID API
ตั้งแต่ 2020-01-01 ถึงเมื่อวาน → บันทึกเป็น api_historical_raw.csv

รองรับ resume: ถ้ารันซ้ำจะข้ามวันที่มีข้อมูลครบแล้ว
ประมาณเวลา: ~2200 วัน × 0.15s ≈ 5-6 นาที
"""

import csv
import os
import time
from datetime import date, timedelta
import requests

# ── ตั้งค่า ──
_DIR          = os.path.dirname(os.path.abspath(__file__))
OUTPUT_CSV    = os.path.join(_DIR, "api_historical_raw.csv")
FETCH_START   = date(2005, 9, 1)    # ประแสร์มีข้อมูลตั้งแต่ 1 ก.ย. 2548
DELAY_SEC     = 0.15                # delay ระหว่างแต่ละวัน
LOOKBACK_DAYS = 5                   # ย้อนหลัง 5 วันเสมอ (API อาจอัปเดตช้า)

FIELDNAMES = ["date_record", "id", "name", "volume", "inflow", "outflow", "percent_storage"]

# ID อ่าง → reservoir (DK, KY) | dam (NPL, PS)
TARGETS = {
    "reservoir": {
        "url":      "https://app.rid.go.th/reservoir/api/reservoir/public/",
        "data_key": "reservoir",
        "ids":      {"rsv357", "rsv359"},   # DK=ดอกกราย, KY=คลองใหญ่
    },
    "dam": {
        "url":      "https://app.rid.go.th/reservoir/api/dam/public/",
        "data_key": "dam",
        "ids":      {"100504", "100505"},   # NPL=หนองปลาไหล, PS=ประแสร์
    },
}


# ── โหลดไฟล์ที่มีอยู่ ──
def load_existing(path):
    """
    คืน (done_dates, rows_ทั้งหมด)
    done_dates = วันที่ครบ 4 อ่าง และเก่ากว่า LOOKBACK_DAYS วัน
    (วัน 5 วันล่าสุดจะถูกดึงซ้ำเสมอ เผื่อ API อัปเดตช้า)
    """
    if not os.path.exists(path):
        return set(), []

    rows = []
    count_by_date = {}
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
            d = row.get("date_record", "")
            count_by_date[d] = count_by_date.get(d, 0) + 1

    lookback_cutoff = (date.today() - timedelta(days=LOOKBACK_DAYS)).strftime("%Y-%m-%d")
    done = {d for d, c in count_by_date.items() if c >= 4 and d < lookback_cutoff}
    return done, rows


# ── ดึง 1 วัน ──
def fetch_day(date_str):
    results = []
    for cat, cfg in TARGETS.items():
        try:
            r = requests.get(cfg["url"] + date_str, timeout=15)
            if r.status_code != 200:
                continue
            data = r.json()
            for region in data.get("data", []):
                for item in region.get(cfg["data_key"], []):
                    if item.get("id") in cfg["ids"]:
                        results.append({
                            "date_record":     date_str,
                            "id":              item.get("id", ""),
                            "name":            item.get("name", ""),
                            "volume":          item.get("volume"),
                            "inflow":          item.get("inflow")  if item.get("inflow")  is not None else 0,
                            "outflow":         item.get("outflow") if item.get("outflow") is not None else 0,
                            "percent_storage": item.get("percent_storage"),
                        })
        except Exception as e:
            print(f"\n  ⚠ {cat} {date_str}: {e}")
    return results


# ── บันทึก CSV ──
def save_csv(path, rows):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


# ── main ──
def main():
    fetch_end = date.today()   # รวมวันนี้ด้วย

    done_dates, existing_rows = load_existing(OUTPUT_CSV)
    if existing_rows:
        print(f"พบไฟล์เดิม {len(existing_rows)} แถว | วันที่ครบแล้ว {len(done_dates)} วัน")

    # รายการวันที่ต้องดึง
    to_fetch = []
    d = FETCH_START
    while d <= fetch_end:
        if d.strftime("%Y-%m-%d") not in done_dates:
            to_fetch.append(d)
        d += timedelta(days=1)

    total = len(to_fetch)
    if total == 0:
        print("✓ ข้อมูลครบแล้ว ไม่ต้องดึงเพิ่ม")
        return

    print(f"ต้องดึง {total} วัน ({FETCH_START} → {fetch_end})")
    print(f"ประมาณเวลา: {total * DELAY_SEC * 2:.0f} วินาที\n")

    new_rows = []
    for i, d in enumerate(to_fetch):
        date_str = d.strftime("%Y-%m-%d")
        rows = fetch_day(date_str)
        new_rows.extend(rows)

        remaining = (total - i - 1) * DELAY_SEC * 2
        print(f"  [{i+1}/{total}] {date_str} ({(i+1)/total*100:.0f}%) | เหลือ ~{remaining:.0f}s    ", end="\r")

        # บันทึก checkpoint ทุก 200 วัน
        if (i + 1) % 200 == 0:
            all_rows = existing_rows + new_rows
            save_csv(OUTPUT_CSV, all_rows)
            existing_rows = all_rows
            new_rows = []
            print(f"\n  💾 checkpoint บันทึก {len(existing_rows)} แถว")

        time.sleep(DELAY_SEC)

    all_rows = existing_rows + new_rows
    save_csv(OUTPUT_CSV, all_rows)
    print(f"\n\n✓ เสร็จ! {len(all_rows)} แถว → {OUTPUT_CSV}")


if __name__ == "__main__":
    main()
