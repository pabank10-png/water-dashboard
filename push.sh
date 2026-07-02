#!/bin/bash
# push.sh — Water Dashboard deploy script
# วิธีใช้: bash push.sh "commit message"

MSG="${1:-update}"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"

cd "$REPO_DIR"
find .git -name "*.lock" -delete 2>/dev/null

# เพิ่มเฉพาะไฟล์ที่แก้มือ — ไม่รวม data.json / water_data.xlsx / api_historical_raw.csv
# (ไฟล์เหล่านั้น pipeline จัดการอยู่แล้ว ถ้า add ด้วยจะ conflict ตอน push)
git add index.html README.md push.sh
git add .github/workflows/ 2>/dev/null || true   # add ถ้ามีการแก้ workflow

git commit -m "$MSG"
git push origin master:main
