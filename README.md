# Water Reservoir Dashboard RY

**Live:** https://pabank10-png.github.io/water-dashboard/

Dashboard สำหรับติดตามและคาดการณ์ปริมาณน้ำในอ่างเก็บน้ำระยอง ประกอบด้วย 3 อ่าง (ดอกกราย / คลองใหญ่ / หนองปลาไหล) และประแสร์ รองรับ Scenario planning ล่วงหน้า 1–2 ปี

---

## Tabs

| Tab | เนื้อหา |
|-----|---------|
| 3อ่าง (DK/KY/NPL) | Water level forecast + Scenario Manager + Water Balance table |
| ประแสร์ (PS) | Water level forecast + Scenario Manager + Water Balance table |
| Inflow | ประวัติ Inflow รายเดือนแยกปี |
| Outflow | ประวัติ Outflow รายเดือนแยกปี |

---

## Key Features

- **Dynamic constants** — โหลดจาก `data.json` อัตโนมัติ ไม่ต้อง hardcode รายเดือน/รายปี
- **Scenario Manager** — สูงสุด 4 Scenarios ต่ออ่าง ปรับ Inflow ปี/ตัวคูณ/Outflow/ผันน้ำรายเดือน
- **1 ปี / 2 ปี mode** — สลับ forecast horizon ต่อ tab
- **Export / Import Settings** — บันทึก scenario ทั้งหมดเป็น `.json` → import คืนวันไหนก็ได้ ค่าจะ recalculate ตามวันที่จริงอัตโนมัติ
- **Night Mode** — สลับ Dark/Light ผ่านปุ่ม topbar
- **Download Data** — ดาวน์โหลด `water_data.xlsx` ประวัติทั้งหมด

---

## Auto-Update Pipeline

```
cron-job.org (09:00 & 12:00 Asia/Bangkok ทุกวัน)
  → POST GitHub API (workflow_dispatch)
  → GitHub Actions: fetch_api_data.py + process_data.py
  → commit data.json, api_historical_raw.csv, water_data.xlsx
  → Dashboard แสดงข้อมูลใหม่อัตโนมัติ
```

ไม่ใช้ `schedule:` cron ใน GitHub Actions (delay 3-4 ชม.) — ใช้ cron-job.org แทนเพื่อความตรงเวลา

---

## Export / Import

**Export** — กดปุ่ม `Export` ใน topbar → ดาวน์โหลด `water-scenario-YYYYMMDD.json`

ไฟล์เก็บ: simYears (1 หรือ 2 ปี), scenarios ทั้งหมด (inflow source, ตัวคูณ, outflow/day, ผันน้ำ/day)

**Import** — กดปุ่ม `Import` → เลือกไฟล์ JSON → dashboard คืนค่าทันที

ค่า `STAR_PROGRESS` และ `STAR_MONTH_IDX` **ไม่ถูก export** — คำนวณจากวันที่จริงใหม่เสมอ ดังนั้น:
- Import กลางเดือน → เดือนปัจจุบัน recalculate ตามวันที่เหลือโดยอัตโนมัติ
- Import เดือนถัดไป → เดือนก่อนกลายเป็น Actual อัตโนมัติ

---

## Changelog

| วันที่ | รายการ |
|--------|--------|
| 2026-06-17 | สร้าง Dashboard ครั้งแรก |
| 2026-06-20 | Deploy GitHub Pages |
| 2026-06-24 | Forecast 2 ปี, ย้าย folder |
| 2026-06-29 | Dynamic ACT_JUN, fix tooltip, banner sub-text |
| 2026-07-01 | Dynamic STAR_MONTH_IDX, fix fcastLine, runSc rem |
| 2026-07-02 | Fix tension:0, แยก tooltip ● vs ★, STAR_LABEL dynamic |
| 2026-07-02 | เพิ่ม Export / Import Scenario Settings |
| 2026-07-02 | ย้าย schedule trigger → cron-job.org (09:00 & 12:00) |
| 2026-07-02 | fix: มิ.ย. (ACT_IDX) แสดง inflow/outflow จาก Scenario แทนค่า 0 |
| 2026-07-02 | ตาราง Monthly Balance เพิ่มคอลัมน์ สิ้นอ่าง (Actual+Scenario) และ สิ้นอ่าง (Scenario) |
