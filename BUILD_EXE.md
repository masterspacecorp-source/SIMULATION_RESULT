# EXE 빌드 가이드 (Windows)

## 1) 준비
```powershell
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install pyinstaller pandas openpyxl xlsxwriter
```

## 2) GUI EXE 빌드 (권장)
`sudp_gui.py`를 엔트리포인트로 빌드하면 EXE 실행 시 바로 GUI가 뜹니다.

```powershell
python -m PyInstaller --clean --onefile --noconsole --name sudp_consolidator sudp_gui.py
```

결과물: `dist\sudp_consolidator.exe`

## 3) 코어 파일로 빌드해도 GUI 자동 실행
아래처럼 `sudp_core.py`로 빌드해도, **인자 없이 실행하면 GUI가 자동 실행**되도록 처리했습니다.

```powershell
python -m PyInstaller --clean --onefile --noconsole --name sudp_consolidator sudp_core.py
```

## 4) 실행
- 더블클릭 후 GUI에서 시작연도/종료연도/자원코드/결과형식/HHV 입력
- 결과 엑셀 생성

## 5) 참고
- 콘솔 버전으로 직접 실행하려면 `--noconsole` 없이 빌드하고 CLI 인자를 넘기면 됩니다.
