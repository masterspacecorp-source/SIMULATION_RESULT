# SUDP 결과 취합기 EXE 빌드 가이드

현재 저장소에는 `sudp_core.py` 1개로 동작하는 실행 스크립트가 있으며,
Windows에서 아래 방식으로 `exe`를 생성할 수 있습니다.

## 1) 준비

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas numpy openpyxl xlsxwriter
```

## 2) 빌드

```powershell
python -m PyInstaller --noconfirm --clean --onefile --name sudp_consolidator sudp_core.py
```

빌드가 완료되면 아래 파일이 생성됩니다.

- `dist/sudp_consolidator.exe`

## 3) 실행 예시

```powershell
.\dist\sudp_consolidator.exe --root "C:\Program Files (x86)\MasterSpace\M-Core\SUDP" --codes "2731,2732" --start-year 2026 --end-year 2035 --result-mode "시간별" --hhv 5000 --ru-price 0.0 --out "C:\temp\result.xlsx"
```

## 4) 콘솔 창 숨기기(선택)

GUI로만 쓸 계획이면 `--onefile` 빌드 시 `--noconsole` 옵션을 추가하세요.
단, 표준출력 로그 확인이 어려워질 수 있습니다.
