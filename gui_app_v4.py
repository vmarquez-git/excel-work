# gui_app_v4.py
# Tkinter GUI wrapper for snapshot comparison (sheet-based, no Excel tables required).

from __future__ import annotations

import logging
import re
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

import app_v4 as app


YYYY_MM_RE = re.compile(r"^\d{4}-\d{2}$")


def _validate_month(value: str, field: str) -> str:
    value = (value or "").strip()
    if not YYYY_MM_RE.match(value):
        raise ValueError(f"{field} must be YYYY-MM (e.g., 2025-11). Got: '{value}'")
    return value


class SnapshotCompareGUI(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Snapshot Compare (A vs B)")
        self.geometry("820x380")

        self.fileA = tk.StringVar()
        self.fileB = tk.StringVar()
        self.a_actuals = tk.StringVar(value="2025-11")
        self.a_forecast = tk.StringVar(value="2025-12")
        self.b_actuals = tk.StringVar(value="2025-11")
        self.b_forecast = tk.StringVar(value="2025-12")
        self.outdir = tk.StringVar(value="Outputs")

        self.status = tk.StringVar(value="Ready.")
        self._build_ui()

    def _build_ui(self) -> None:
        pad = {"padx": 10, "pady": 6}

        self.columnconfigure(1, weight=1)
        self.columnconfigure(3, weight=0)

        def add_file_row(row: int, label: str, var: tk.StringVar):
            tk.Label(self, text=label, anchor="w").grid(row=row, column=0, sticky="w", **pad)
            tk.Entry(self, textvariable=var, width=72).grid(row=row, column=1, columnspan=2, sticky="we", **pad)
            tk.Button(self, text="Browse...", command=lambda: self._browse_file(var)).grid(
                row=row, column=3, sticky="e", **pad
            )

        add_file_row(0, "File A:", self.fileA)
        add_file_row(1, "File B:", self.fileB)

        tk.Label(self, text="A_actuals_month (YYYY-MM):", anchor="w").grid(row=2, column=0, sticky="w", **pad)
        tk.Entry(self, textvariable=self.a_actuals, width=18).grid(row=2, column=1, sticky="w", **pad)

        tk.Label(self, text="B_actuals_month (YYYY-MM):", anchor="w").grid(row=2, column=2, sticky="w", **pad)
        tk.Entry(self, textvariable=self.b_actuals, width=18).grid(row=2, column=3, sticky="w", **pad)

        tk.Label(self, text="A_forecast_start (YYYY-MM):", anchor="w").grid(row=3, column=0, sticky="w", **pad)
        tk.Entry(self, textvariable=self.a_forecast, width=18).grid(row=3, column=1, sticky="w", **pad)

        tk.Label(self, text="B_forecast_start (YYYY-MM):", anchor="w").grid(row=3, column=2, sticky="w", **pad)
        tk.Entry(self, textvariable=self.b_forecast, width=18).grid(row=3, column=3, sticky="w", **pad)

        tk.Label(self, text="Output folder:", anchor="w").grid(row=4, column=0, sticky="w", **pad)
        tk.Entry(self, textvariable=self.outdir, width=72).grid(row=4, column=1, columnspan=2, sticky="we", **pad)
        tk.Button(self, text="Browse...", command=self._browse_outdir).grid(row=4, column=3, sticky="e", **pad)

        self.run_btn = tk.Button(self, text="Run", command=self._run_clicked)
        self.run_btn.grid(row=5, column=2, sticky="e", padx=10, pady=10)

        tk.Button(self, text="Close", command=self.destroy).grid(row=5, column=3, sticky="e", padx=10, pady=10)

        tk.Label(self, textvariable=self.status, anchor="w", fg="blue").grid(
            row=6, column=0, columnspan=4, sticky="we", padx=10, pady=6
        )

    def _browse_file(self, target_var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
        )
        if path:
            target_var.set(path)

    def _browse_outdir(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.outdir.set(path)

    def _set_status(self, message: str) -> None:
        self.after(0, lambda: self.status.set(message))

    def _run_clicked(self) -> None:
        self.status.set("Starting process...")
        self.update_idletasks()
        threading.Thread(target=self._run_job, daemon=True).start()

    def _run_job(self) -> None:
        try:
            self._set_running(True)

            self._set_status("Checking selected files...")
            file_a = Path(self.fileA.get().strip()).expanduser().resolve()
            file_b = Path(self.fileB.get().strip()).expanduser().resolve()

            if not file_a.exists():
                raise FileNotFoundError(f"File A not found: {file_a}")
            if not file_b.exists():
                raise FileNotFoundError(f"File B not found: {file_b}")

            a_act = _validate_month(self.a_actuals.get(), "A_actuals_month")
            a_fc = _validate_month(self.a_forecast.get(), "A_forecast_start")
            b_act = _validate_month(self.b_actuals.get(), "B_actuals_month")
            b_fc = _validate_month(self.b_forecast.get(), "B_forecast_start")

            out_dir = Path(self.outdir.get().strip() or "Outputs").expanduser().resolve()
            log_path = app.setup_logging(out_dir)

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = out_dir / f"Compare_{ts}.xlsx"

            logging.info("GUI run started")
            logging.info("FileA=%s", file_a)
            logging.info("FileB=%s", file_b)
            logging.info("A: actuals=%s forecast=%s", a_act, a_fc)
            logging.info("B: actuals=%s forecast=%s", b_act, b_fc)
            logging.info("Output=%s", out_path)
            logging.info("Log=%s", log_path)

            self._set_status("Validating Excel sheets...")
            logging.info("Validating required sheets in File A")
            app.validate_required_sheets(file_a)

            logging.info("Validating required sheets in File B")
            app.validate_required_sheets(file_b)

            self._set_status("Running SnapshotCheck for File A...")
            logging.info("Running SnapshotCheck for File A")
            snap_a = app.snapshot_check(file_a, a_act, a_fc)

            self._set_status("Running PathCheck for File A...")
            logging.info("Running PathCheck for File A")
            path_a = app.path_check(file_a)

            self._set_status("Running SnapshotCheck for File B...")
            logging.info("Running SnapshotCheck for File B")
            snap_b = app.snapshot_check(file_b, b_act, b_fc)

            self._set_status("Running PathCheck for File B...")
            logging.info("Running PathCheck for File B")
            path_b = app.path_check(file_b)

            self._set_status("Running comparison sheets...")
            logging.info("Running SnapshotCheck comparison")
            snap_compare = app.snapshotcheck_compare(snap_a, snap_b)

            logging.info("Running Actuals by month comparison")
            actuals_cmp = app.actuals_compare_by_month(file_a, file_b)

            logging.info("Running EAC total comparison")
            eac_total = app.eac_compare_total(file_a, file_b)

            logging.info("Running EAC by month comparison")
            eac_by_month = app.eac_compare_by_month(file_a, file_b)

            sheets = {
                "SnapshotCheck_A": snap_a,
                "PathCheck_A": path_a,
                "SnapshotCheck_B": snap_b,
                "PathCheck_B": path_b,
                "SnapshotCheck_Compare": snap_compare,
                "Actuals_CompareByMonth": actuals_cmp,
                "EAC_Compare_Total": eac_total,
                "EAC_Compare_ByMonth": eac_by_month,
            }

            self._set_status("Writing output workbook...")
            logging.info("Writing output workbook")
            app.write_outputs(out_path, sheets)

            logging.info("Process completed successfully")
            self.after(
                0,
                lambda: self.status.set(f"Done. Output: {out_path.name} | Log: {Path(log_path).name}")
            )
            self.after(
                0,
                lambda: messagebox.showinfo("Complete", f"Created:\n{out_path}\n\nLog:\n{log_path}")
            )

        except Exception as e:
            logging.exception("Process failed")
            err_msg = str(e)
            self.after(0, lambda: self.status.set(f"Failed: {err_msg}"))
            self.after(0, lambda: messagebox.showerror("Error", err_msg))

        finally:
            self._set_running(False)

    def _set_running(self, running: bool) -> None:
        def _apply():
            self.run_btn.config(state=("disabled" if running else "normal"))
            if running:
                self.status.set("Running...")
        self.after(0, _apply)


if __name__ == "__main__":
    SnapshotCompareGUI().mainloop()
