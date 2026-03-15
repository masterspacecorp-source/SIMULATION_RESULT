from __future__ import annotations

import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from sudp_core import DEFAULT_ROOT, run


class SudpApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("SUDP 결과 취합")
        self.geometry("960x720")

        self.var_start_year = tk.StringVar()
        self.var_start_month = tk.StringVar()
        self.var_end_year = tk.StringVar()
        self.var_end_month = tk.StringVar()
        self.var_reserve_price = tk.StringVar()
        self.var_hhv = tk.StringVar()
        self.var_result_mode = tk.StringVar(value="시간별")
        self.var_codes = tk.StringVar()
        self.var_out = tk.StringVar()
        self.var_root = tk.StringVar(value=str(DEFAULT_ROOT))

        self.var_snapshot = tk.StringVar(value="No")
        self.var_snapshot_mode = tk.StringVar(value="zip")
        self.var_snapshot_name = tk.StringVar()
        self.var_snapshot_dir = tk.StringVar()
        self.var_snapshot_same_folder = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 8, "pady": 6}
        top = ttk.Frame(self)
        top.pack(fill="x", padx=12, pady=12)

        ttk.Label(top, text="시작연도").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_start_year, width=12).grid(row=0, column=1, sticky="w", **pad)

        ttk.Label(top, text="시작월").grid(row=0, column=2, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_start_month, width=8).grid(row=0, column=3, sticky="w", **pad)

        ttk.Label(top, text="종료연도").grid(row=0, column=4, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_end_year, width=12).grid(row=0, column=5, sticky="w", **pad)

        ttk.Label(top, text="종료월").grid(row=0, column=6, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_end_month, width=8).grid(row=0, column=7, sticky="w", **pad)

        ttk.Label(top, text="예비력 단가").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_reserve_price, width=12).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(top, text="적용 발열량(HHV, kcal/kg)").grid(row=1, column=2, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_hhv, width=14).grid(row=1, column=3, sticky="w", **pad)

        ttk.Label(top, text="결과 형식").grid(row=1, column=4, sticky="w", **pad)
        ttk.Combobox(top, textvariable=self.var_result_mode, values=["시간별", "연도별"], width=10, state="readonly").grid(
            row=1, column=5, sticky="w", **pad
        )

        ttk.Label(top, text="자원코드").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_codes, width=55).grid(row=2, column=1, columnspan=7, sticky="we", **pad)

        ttk.Label(top, text="SUDP 루트").grid(row=3, column=0, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_root, width=80).grid(row=3, column=1, columnspan=6, sticky="we", **pad)
        ttk.Button(top, text="찾아보기", command=self._choose_root).grid(row=3, column=7, sticky="e", **pad)

        ttk.Label(top, text="저장경로").grid(row=4, column=0, sticky="w", **pad)
        ttk.Entry(top, textvariable=self.var_out, width=80).grid(row=4, column=1, columnspan=6, sticky="we", **pad)
        ttk.Button(top, text="찾아보기", command=self._choose_out).grid(row=4, column=7, sticky="e", **pad)

        snap = ttk.LabelFrame(self, text="스냅샷 옵션 (SUDP 폴더 백업)")
        snap.pack(fill="x", padx=12, pady=(0, 10))

        ttk.Label(snap, text="스냅샷").grid(row=0, column=0, sticky="w", **pad)
        ttk.Radiobutton(snap, text="No", variable=self.var_snapshot, value="No", command=self._sync_snapshot_state).grid(row=0, column=1, sticky="w", **pad)
        ttk.Radiobutton(snap, text="Yes", variable=self.var_snapshot, value="Yes", command=self._sync_snapshot_state).grid(row=0, column=2, sticky="w", **pad)

        ttk.Label(snap, text="형식").grid(row=1, column=0, sticky="w", **pad)
        self.rb_zip = ttk.Radiobutton(snap, text="zip", variable=self.var_snapshot_mode, value="zip")
        self.rb_copy = ttk.Radiobutton(snap, text="copy", variable=self.var_snapshot_mode, value="copy")
        self.rb_zip.grid(row=1, column=1, sticky="w", **pad)
        self.rb_copy.grid(row=1, column=2, sticky="w", **pad)

        ttk.Label(snap, text="이름").grid(row=2, column=0, sticky="w", **pad)
        self.ent_snap_name = ttk.Entry(snap, textvariable=self.var_snapshot_name, width=35)
        self.ent_snap_name.grid(row=2, column=1, columnspan=3, sticky="we", **pad)

        ttk.Label(snap, text="저장폴더").grid(row=3, column=0, sticky="w", **pad)
        self.ent_snap_dir = ttk.Entry(snap, textvariable=self.var_snapshot_dir, width=55)
        self.ent_snap_dir.grid(row=3, column=1, columnspan=3, sticky="we", **pad)
        self.btn_snap_dir = ttk.Button(snap, text="찾아보기", command=self._choose_snapshot_dir)
        self.btn_snap_dir.grid(row=3, column=4, sticky="e", **pad)
        self.chk_same_dir = ttk.Checkbutton(
            snap,
            text="결과 엑셀과 같은 폴더 사용",
            variable=self.var_snapshot_same_folder,
            command=self._sync_snapshot_state,
        )
        self.chk_same_dir.grid(row=3, column=5, sticky="w", **pad)

        self._sync_snapshot_state()

        log_frame = ttk.Frame(self)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(4, 10))
        ttk.Label(log_frame, text="상태").pack(anchor="w")
        self.txt = tk.Text(log_frame, height=18)
        self.txt.pack(fill="both", expand=True)

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        self.btn_run = ttk.Button(btns, text="실행", command=self._on_run)
        self.btn_run.pack()

    def _choose_root(self) -> None:
        d = filedialog.askdirectory(title="SUDP 루트 폴더 선택")
        if d:
            self.var_root.set(d)

    def _choose_out(self) -> None:
        p = filedialog.asksaveasfilename(
            title="결과 엑셀 저장 경로",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if p:
            self.var_out.set(p)

    def _choose_snapshot_dir(self) -> None:
        d = filedialog.askdirectory(title="스냅샷 저장 폴더 선택")
        if d:
            self.var_snapshot_dir.set(d)

    def _sync_snapshot_state(self) -> None:
        enabled = self.var_snapshot.get() == "Yes"
        state = "normal" if enabled else "disabled"
        for w in [self.rb_zip, self.rb_copy, self.ent_snap_name, self.chk_same_dir]:
            w.configure(state=state)

        same_dir = self.var_snapshot_same_folder.get()
        self.ent_snap_dir.configure(state=("disabled" if (not enabled or same_dir) else "normal"))
        self.btn_snap_dir.configure(state=("disabled" if (not enabled or same_dir) else "normal"))

    def _append_log(self, text: str) -> None:
        self.txt.insert("end", text + "\n")
        self.txt.see("end")

    def _on_run(self) -> None:
        self.btn_run.configure(state="disabled")
        self._append_log("실행 중...")
        t = threading.Thread(target=self._run_job, daemon=True)
        t.start()

    def _run_job(self) -> None:
        try:
            start_year = int(self.var_start_year.get().strip())
            start_month = int((self.var_start_month.get().strip() or "1"))
            end_year = int(self.var_end_year.get().strip())
            end_month = int((self.var_end_month.get().strip() or "12"))
            reserve_price = float(self.var_reserve_price.get().strip() or 0)
            hhv = float(self.var_hhv.get().strip() or 4500)
            codes_csv = self.var_codes.get().strip()
            if not codes_csv:
                raise ValueError("자원코드를 입력하세요.")

            out_path = self.var_out.get().strip()
            if not out_path:
                ys, ye = sorted([start_year, end_year])
                out_path = str(Path.cwd() / f"SUDP_{ys}_{ye}.xlsx")

            do_snapshot = self.var_snapshot.get() == "Yes"
            snapshot_out = self.var_snapshot_dir.get().strip() or None
            use_default_snapshot_dir = self.var_snapshot_same_folder.get()

            run(
                root=self.var_root.get().strip() or str(DEFAULT_ROOT),
                codes_csv=codes_csv,
                start_year=start_year,
                start_month=start_month,
                end_year=end_year,
                end_month=end_month,
                out_path=out_path,
                reserve_price=reserve_price,
                result_mode=self.var_result_mode.get().strip() or "시간별",
                applied_hhv=hhv,
                snapshot=do_snapshot,
                snapshot_mode=self.var_snapshot_mode.get().strip() or "zip",
                snapshot_name=self.var_snapshot_name.get().strip() or None,
                snapshot_out=snapshot_out,
                use_default_snapshot_dir=use_default_snapshot_dir,
            )
            self.after(0, lambda: self._append_log(f"완료: {out_path}"))
        except Exception as e:
            tb = traceback.format_exc()
            self.after(0, lambda: self._append_log(f"오류: {e}\n{tb}"))
            self.after(0, lambda: messagebox.showerror("오류", str(e)))
        finally:
            self.after(0, lambda: self.btn_run.configure(state="normal"))


def main() -> None:
    app = SudpApp()
    app.mainloop()


if __name__ == "__main__":
    main()
