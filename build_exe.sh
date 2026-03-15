#!/usr/bin/env bash
set -euo pipefail

# GUI 실행형 EXE 생성 (권장)
python -m PyInstaller --clean --onefile --noconsole --name sudp_consolidator_gui sudp_gui.py

# 필요 시 코어 CLI EXE도 생성하려면 아래 주석 해제
# python -m PyInstaller --clean --onefile --console --name sudp_consolidator_core sudp_core.py

echo "Done. EXE output: ./dist/sudp_consolidator_gui(.exe)"
