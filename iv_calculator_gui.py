"""
iv_calculator_gui.py — Desktop GUI for the Intrinsic Value Calculator.

Just enter a ticker (and optionally change the year), click Generate, and
it fetches everything from Yahoo Finance and writes a filled-in copy of
Intrinsic_Value_Calculator.xlsx.

Run:
    pip install yfinance openpyxl --break-system-packages
    python iv_calculator_gui.py

Requires Intrinsic_Value_Calculator.xlsx (the template) to be in the same
folder as this script, or pick a different one via "Browse...".
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import iv_core

APP_TITLE = "Intrinsic Value Calculator"
DEFAULT_TEMPLATE_NAME = "Intrinsic_Value_Calculator.xlsx"


class IVCalculatorApp:
    def __init__(self, root):
        self.root = root
        root.title(APP_TITLE)
        root.geometry("640x560")
        root.minsize(560, 480)

        pad = {"padx": 10, "pady": 6}

        # --- Template file picker ---
        frm_template = ttk.LabelFrame(root, text="Template workbook")
        frm_template.pack(fill="x", **pad)

        default_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), DEFAULT_TEMPLATE_NAME)
        self.template_var = tk.StringVar(value=default_template if os.path.exists(default_template) else "")

        ttk.Entry(frm_template, textvariable=self.template_var).pack(side="left", fill="x", expand=True, padx=(10, 5), pady=8)
        ttk.Button(frm_template, text="Browse...", command=self.browse_template).pack(side="left", padx=(0, 10), pady=8)

        # --- Main inputs ---
        frm_inputs = ttk.LabelFrame(root, text="Stock")
        frm_inputs.pack(fill="x", **pad)

        ttk.Label(frm_inputs, text="Ticker / Symbol:").grid(row=0, column=0, sticky="w", padx=10, pady=6)
        self.ticker_var = tk.StringVar()
        ttk.Entry(frm_inputs, textvariable=self.ticker_var, width=20).grid(row=0, column=1, sticky="w", padx=10, pady=6)

        ttk.Label(frm_inputs, text="Current Year:").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        self.year_var = tk.StringVar(value=str(iv_core.default_year()))
        ttk.Entry(frm_inputs, textvariable=self.year_var, width=20).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        # --- Advanced (optional overrides) ---
        frm_adv = ttk.LabelFrame(root, text="Advanced (optional — leave blank to auto-fetch)")
        frm_adv.pack(fill="x", **pad)

        ttk.Label(frm_adv, text="Last Close Price:").grid(row=0, column=0, sticky="w", padx=10, pady=6)
        self.last_close_var = tk.StringVar()
        ttk.Entry(frm_adv, textvariable=self.last_close_var, width=20).grid(row=0, column=1, sticky="w", padx=10, pady=6)

        ttk.Label(frm_adv, text="Discount Rate (e.g. 0.086):").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        self.discount_rate_var = tk.StringVar()
        ttk.Entry(frm_adv, textvariable=self.discount_rate_var, width=20).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        # --- Output file picker ---
        frm_out = ttk.LabelFrame(root, text="Output file")
        frm_out.pack(fill="x", **pad)
        self.output_var = tk.StringVar()
        ttk.Entry(frm_out, textvariable=self.output_var).pack(side="left", fill="x", expand=True, padx=(10, 5), pady=8)
        ttk.Button(frm_out, text="Browse...", command=self.browse_output).pack(side="left", padx=(0, 10), pady=8)

        # --- Generate button ---
        self.generate_btn = ttk.Button(root, text="Generate", command=self.on_generate)
        self.generate_btn.pack(pady=(4, 8))

        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.pack(fill="x", padx=10)

        # --- Log output ---
        frm_log = ttk.LabelFrame(root, text="Log")
        frm_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frm_log, height=12, state="disabled", wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)

        self.ticker_var.trace_add("write", self._sync_output_name)

    def _sync_output_name(self, *_):
        ticker = self.ticker_var.get().strip().upper()
        if ticker:
            self.output_var.set(f"{ticker}_Intrinsic_Value.xlsx")

    def browse_template(self):
        path = filedialog.askopenfilename(title="Select template workbook", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.template_var.set(path)

    def browse_output(self):
        default_name = self.output_var.get() or "Intrinsic_Value.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save output as", defaultextension=".xlsx",
            initialfile=default_name, filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def log(self, msg):
        def _append():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", str(msg) + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, _append)

    def on_generate(self):
        template = self.template_var.get().strip()
        ticker = self.ticker_var.get().strip().upper()
        year_str = self.year_var.get().strip()
        output = self.output_var.get().strip() or f"{ticker or 'output'}_Intrinsic_Value.xlsx"

        if not template or not os.path.exists(template):
            messagebox.showerror(APP_TITLE, "Please choose a valid template workbook.")
            return
        if not ticker:
            messagebox.showerror(APP_TITLE, "Please enter a stock ticker.")
            return
        try:
            year = int(year_str) if year_str else iv_core.default_year()
        except ValueError:
            messagebox.showerror(APP_TITLE, "Year must be a whole number.")
            return

        last_close_str = self.last_close_var.get().strip()
        discount_rate_str = self.discount_rate_var.get().strip()
        try:
            manual_last_close = float(last_close_str) if last_close_str else None
            manual_discount_rate = float(discount_rate_str) if discount_rate_str else None
        except ValueError:
            messagebox.showerror(APP_TITLE, "Last Close and Discount Rate must be numbers.")
            return

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        self.generate_btn.configure(state="disabled")
        self.progress.start(12)

        def worker():
            try:
                iv_core.build_workbook(
                    template, output, ticker, year,
                    manual_last_close=manual_last_close,
                    manual_discount_rate=manual_discount_rate,
                    log=self.log,
                )
                self.root.after(0, lambda: messagebox.showinfo(APP_TITLE, f"Done! Saved to:\n{output}"))
            except Exception as e:
                self.log(f"ERROR: {e}")
                self.root.after(0, lambda: messagebox.showerror(APP_TITLE, f"Failed to generate:\n{e}"))
            finally:
                self.root.after(0, self.progress.stop)
                self.root.after(0, lambda: self.generate_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = IVCalculatorApp(root)
    root.mainloop()
