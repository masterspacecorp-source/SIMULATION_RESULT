# EXE 빌드 가이드

SUDP 결과 취합 도구는 **GUI 입력 방식**으로 사용하는 것을 권장합니다.

## 1) 권장 빌드 (GUI EXE)
아래 명령으로 빌드하세요.

```bash
python -m PyInstaller --clean --onefile --noconsole --name sudp_consolidator_gui sudp_gui.py
```

- 실행 파일: `dist/sudp_consolidator_gui.exe`
- 실행하면 GUI 화면에서 시작/종료연도, 자원코드, 예비력 단가 등을 입력해 처리할 수 있습니다.

## 2) (선택) 코어 CLI EXE
코어를 직접 EXE로 만드는 경우에는 콘솔 인자를 전달해 실행해야 합니다.

```bash
python -m PyInstaller --clean --onefile --console --name sudp_consolidator_core sudp_core.py
```

예시:

```bash
sudp_consolidator_core.exe --codes INCC1,INCC2 --start-year 2026 --end-year 2035 --ru-price 1.0
```

> 참고: `--noconsole`로 `sudp_core.py`를 빌드해 인자 없이 더블클릭 실행하면,
> 코드에서 GUI 폴백을 시도하지만 운영 환경에 따라 제한될 수 있어 GUI EXE(`sudp_gui.py` 기반) 빌드를 권장합니다.
