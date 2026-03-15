#!/usr/bin/env bash
set -euo pipefail

if ! python -m PyInstaller --version >/dev/null 2>&1; then
  echo "[오류] PyInstaller가 설치되어 있지 않습니다."
  echo "      python -m pip install pyinstaller"
  exit 1
fi

TARGET="sudp_gui.py"
if [[ ! -f "$TARGET" ]]; then
  TARGET="sudp_core.py"
fi

echo "[빌드] target=$TARGET"
python -m PyInstaller --clean --onefile --noconsole --name sudp_consolidator "$TARGET"
echo "[완료] dist/sudp_consolidator(.exe)"
