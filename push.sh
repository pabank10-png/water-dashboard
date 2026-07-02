#!/bin/bash
# push.sh — Water Dashboard deploy script
# วิธีใช้: bash push.sh "commit message"

MSG="${1:-update}"
REPO_DIR="$(cd "$(dirname "$0")" && pwd)"

cd "$REPO_DIR"
find .git -name "*.lock" -delete 2>/dev/null
git add -A
git commit -m "$MSG"
git push origin master:main
