# sudp_gui.py
import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ★ 코어 모듈 (sudp_core.py의 run 함수를 사용)
try:
    from sudp_core import run as generate_report
except Exception as e:
    raise ImportError(
        "sudp_core.py에서 run 함수를 import하지 못했습니다. "
        "코어 스크립트를 sudp_core.py로 저장했는지 확인하세요.\n"
        f"원인: {e}"
    )

DEFAULT_ROOT = r"C:\Program Files (x86)\MasterSpace\M-Core\SUDP"

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SUDP 결과 취합")
        self.geometry("820x620")
        self.resizable(False, False)

        # DPI 인식(윈도우에서 글자 또렷하게)
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        main = ttk.Frame(self, padding=18)
        main.pack(fill="both", expand=True)

        # ================== 1) 기본 입력 ==================
        row1 = ttk.Frame(main); row1.pack(fill="x", pady=(0, 10))
        ttk.Label(row1, text="시작연도", width=10).pack(side="left")
        self.var_start = tk.StringVar()
        ttk.Entry(row1, textvariable=self.var_start, width=14).pack(side="left", padx=(5, 22))

        ttk.Label(row1, text="종료연도", width=10).pack(side="left")
        self.var_end = tk.StringVar()
        ttk.Entry(row1, textvariable=self.var_end, width=14).pack(side="left", padx=5)

        # 예비력용량가치 단가
        ttk.Label(row1, text="예비력 단가", width=12).pack(side="left", padx=(24, 0))
        self.var_ru_price = tk.StringVar()
        ttk.Entry(row1, textvariable=self.var_ru_price, width=14).pack(side="left", padx=5)
        ttk.Label(row1, text="(예: 1.0)", foreground="#666").pack(side="left")

        # 자원코드
        row2 = ttk.Frame(main); row2.pack(fill="x", pady=10)
        ttk.Label(row2, text="자원코드", width=10).pack(side="left")
        self.var_codes = tk.StringVar()
        ttk.Entry(row2, textvariable=self.var_codes).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Label(row2, text="콤마로 구분 (예: INCC1,INCC2)", foreground="#666").pack(side="left", padx=(10, 0))

        # 저장경로(결과 엑셀)
        row3 = ttk.Frame(main); row3.pack(fill="x", pady=10)
        ttk.Label(row3, text="저장경로", width=10).pack(side="left")
        self.var_out = tk.StringVar()
        ttk.Entry(row3, textvariable=self.var_out).pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row3, text="찾아보기", command=self.choose_out_path).pack(side="left", padx=5)

        # ================== 2) 스냅샷 옵션 ==================
        box = ttk.Labelframe(main, text="스냅샷 옵션 (SUDP 폴더 백업)")
        box.pack(fill="x", pady=12)

        # 스냅샷 사용 여부
        row_s1 = ttk.Frame(box); row_s1.pack(fill="x", pady=(6, 6))
        ttk.Label(row_s1, text="스냅샷", width=10).pack(side="left")

        self.var_snapshot = tk.StringVar(value="no")   # "yes" / "no"
        rb_no  = ttk.Radiobutton(row_s1, text="No",  value="no",  variable=self.var_snapshot, command=self._refresh_snapshot_state)
        rb_yes = ttk.Radiobutton(row_s1, text="Yes", value="yes", variable=self.var_snapshot, command=self._refresh_snapshot_state)
        rb_no.pack(side="left", padx=2)
        rb_yes.pack(side="left", padx=12)

        # 형식 zip / copy
        row_s2 = ttk.Frame(box); row_s2.pack(fill="x", pady=(0, 6))
        ttk.Label(row_s2, text="형식", width=10).pack(side="left")
        self.var_snap_mode = tk.StringVar(value="zip")
        self.rb_zip  = ttk.Radiobutton(row_s2, text="zip",  value="zip",  variable=self.var_snap_mode)
        self.rb_copy = ttk.Radiobutton(row_s2, text="copy", value="copy", variable=self.var_snap_mode)
        self.rb_zip.pack(side="left", padx=2); self.rb_copy.pack(side="left", padx=12)

        # 스냅샷 이름
        row_s3 = ttk.Frame(box); row_s3.pack(fill="x", pady=(0, 6))
        ttk.Label(row_s3, text="이름", width=10).pack(side="left")
        self.var_snap_name = tk.StringVar()
        self.ent_snap_name = ttk.Entry(row_s3, textvariable=self.var_snap_name)
        self.ent_snap_name.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Label(row_s3, text="(빈칸이면 자동 생성)").pack(side="left")

        # 스냅샷 경로 + 기본 경로 체크
        row_s4 = ttk.Frame(box); row_s4.pack(fill="x", pady=(0, 8))
        ttk.Label(row_s4, text="저장폴더", width=10).pack(side="left")
        self.var_snap_dir = tk.StringVar()
        self.ent_snap_dir = ttk.Entry(row_s4, textvariable=self.var_snap_dir)
        self.ent_snap_dir.pack(side="left", fill="x", expand=True, padx=5)
        ttk.Button(row_s4, text="찾아보기", command=self.choose_snap_dir).pack(side="left", padx=5)

        self.var_snap_use_default = tk.BooleanVar(value=True)
        self.chk_snap_default = ttk.Checkbutton(
            row_s4,
            text="결과 엑셀과 같은 폴더 사용",
            variable=self.var_snap_use_default,
            command=self._refresh_snapshot_state
        )
        self.chk_snap_default.pack(side="left", padx=(12, 0))

        # ================== 3) 상태/실행 ==================
        row4 = ttk.Frame(main); row4.pack(fill="both", expand=True, pady=(10, 0))
        self.var_status = tk.StringVar(value="준비 완료")
        ttk.Label(row4, textvariable=self.var_status, foreground="#444").pack(anchor="w")

        self.txt = tk.Text(row4, height=10, state="disabled")
        self.txt.pack(fill="both", expand=True, pady=(5, 0))

        row5 = ttk.Frame(main); row5.pack(fill="x", pady=15)
        self.btn_run = ttk.Button(row5, text="실행", command=self.on_run_clicked)
        self.btn_run.pack(pady=5)
        self.prog = ttk.Progressbar(row5, mode="indeterminate")

        # 기본값
        self.var_start.set("")
        self.var_end.set("")
        self.var_codes.set("")
        self.var_out.set("")
        self.var_ru_price.set("")  # 사용자가 입력
        self._refresh_snapshot_state()  # 초기 비활성화 상태 반영

    # -------- 유틸 ----------
    def choose_out_path(self):
        start = (self.var_start.get() or "start")
        end = (self.var_end.get() or "end")
        default_name = f"SUDP_{start}_{end}.xlsx"
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.var_out.set(path)
            # 결과 경로 바뀌면 '같은 폴더 사용' 체크 시 스냅샷 폴더도 맞춰줌
            if self.var_snap_use_default.get():
                self.var_snap_dir.set(str(Path(path).parent))

    def choose_snap_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.var_snap_dir.set(d)

    def _refresh_snapshot_state(self):
        enabled = (self.var_snapshot.get() == "yes")
        state = "normal" if enabled else "disabled"
        for w in (self.rb_zip, self.rb_copy, self.ent_snap_name, self.ent_snap_dir, self.chk_snap_default):
            w.configure(state=state)

        if enabled:
            # 기본 폴더 사용 체크 시 경로 비활성화 + 자동 지정
            if self.var_snap_use_default.get():
                self.ent_snap_dir.configure(state="disabled")
                out_path = self.var_out.get().strip()
                if out_path:
                    self.var_snap_dir.set(str(Path(out_path).parent))
            else:
                self.ent_snap_dir.configure(state="normal")

    # -------- 실행 ----------
    def on_run_clicked(self):
        # 입력 검증
        try:
            start = int(self.var_start.get().strip())
            end = int(self.var_end.get().strip())
        except Exception:
            messagebox.showerror("입력 오류", "시작/종료연도를 숫자로 입력해주세요.")
            return

        codes = self.var_codes.get().strip()
        if not codes:
            messagebox.showerror("입력 오류", "자원코드를 콤마로 입력해주세요.")
            return

        # 예비력 단가
        ru_text = self.var_ru_price.get().strip()
        try:
            reserve_price = float(ru_text)
        except Exception:
            messagebox.showerror("입력 오류", "예비력 단가를 숫자로 입력해주세요. (예: 1.0)")
            return

        out_path = self.var_out.get().strip()
        if not out_path:
            s, e = sorted([start, end])
            out_path = str(Path.cwd() / f"SUDP_{s}_{e}.xlsx")
            self.var_out.set(out_path)

        # 스냅샷 파라미터 구성
        snapshot = (self.var_snapshot.get() == "yes")
        snapshot_mode = self.var_snap_mode.get() if snapshot else None
        snapshot_name = self.var_snap_name.get().strip() if snapshot else None

        if snapshot:
            if self.var_snap_use_default.get():
                snap_dir = str(Path(out_path).parent)
            else:
                snap_dir = self.var_snap_dir.get().strip()
                if not snap_dir:
                    messagebox.showerror("입력 오류", "스냅샷 폴더를 선택해주세요.")
                    return
        else:
            snap_dir = None

        # 진행 UI
        self.btn_run.config(state="disabled")
        self.prog.pack(pady=(5, 0))
        self.prog.start(12)
        self.var_status.set("처리 중... 잠시만 기다려주세요.")

        # 로그
        self._append_log(f"[시작] {start}~{end}, 코드: {codes}\n저장: {out_path}\n")
        self._append_log(f"예비력 단가: {reserve_price}\n")
        if snapshot:
            self._append_log(f"스냅샷: {snapshot_mode}, 이름='{snapshot_name or '(자동)'}', 폴더={snap_dir}\n")
        else:
            self._append_log("스냅샷: 사용 안 함\n")

        # 백그라운드 실행
        t = threading.Thread(
            target=self._run_job,
            args=(codes, start, end, out_path, reserve_price, snapshot, snapshot_mode, snap_dir, snapshot_name, self.var_snap_use_default.get()),
            daemon=True
        )
        t.start()

    def _run_job(self, codes, start, end, out_path, reserve_price, snapshot, snapshot_mode, snap_dir, snapshot_name, use_default_snapshot_dir):
        try:
            # sudp_core.run의 최신 시그니처에 맞춰 호출
            generate_report(
                root=DEFAULT_ROOT,
                codes_csv=codes,
                start_year=start,
                end_year=end,
                out_path=out_path,
                reserve_price=reserve_price,          # ★ 단가 전달
                snapshot=snapshot,
                snapshot_mode=(snapshot_mode or "zip"),
                snapshot_out=snap_dir,
                snapshot_name=snapshot_name,
                use_default_snapshot_dir=use_default_snapshot_dir,
            )
            self._append_log("[완료] 엑셀 생성 성공\n")
            self._set_status("완료! 엑셀 파일이 생성되었습니다.")
            messagebox.showinfo("완료", "엑셀 파일 생성이 완료되었습니다.")
        except Exception as e:
            self._append_log("".join(traceback.format_exc()))
            self._set_status("실패: 오류가 발생했습니다.")
            messagebox.showerror("오류", f"실행 중 오류가 발생했습니다.\n\n{e}")
        finally:
            self.btn_run.config(state="normal")
            self.prog.stop()
            self.prog.pack_forget()

    # -------- 공통 ----------
    def _append_log(self, msg: str):
        self.txt.config(state="normal")
        self.txt.insert("end", msg)
        self.txt.see("end")
        self.txt.config(state="disabled")

    def _set_status(self, msg: str):
        self.var_status.set(msg)

if __name__ == "__main__":
    app = App()
    app.mainloop()
