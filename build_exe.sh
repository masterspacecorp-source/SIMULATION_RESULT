#!/usr/bin/env bash
set -euo pipefail

if ! python -c "import PyInstaller" >/dev/null 2>&1; then
  echo "[ERROR] PyInstaller가 설치되어 있지 않습니다."
  echo "        python -m pip install pyinstaller"
  exit 1
fi

python -m PyInstaller \
  --noconfirm \
  --clean \
  --onefile \
  --name sudp_consolidator \
  sudp_core.py

echo "[DONE] dist/sudp_consolidator 생성"
