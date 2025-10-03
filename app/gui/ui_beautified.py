# app/gui/main_window_bootstrap.py
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from pathlib import Path

from app.config import default_save_dir
from app.services.excel_utils import (
    unique_path, merge_many_files,
    remove_spaces_first_col_and_drop_blank_rows,
    convert_formulas_to_values,
)
from app.gui.formula_cleaner import ExcelFormulaRemover

COLUMN_NAMES = ['Year','Month','Country','Model-Type','Model-Size','Model','Currency','Column1','Model Price','Column3','Forecase Qty']

class MainWindowBootstrap:
    def __init__(self, root: tb.Window):
        self.root = root
        self.root.geometry("380x520")

        container = tb.Frame(root, padding=20)
        container.pack(fill=BOTH, expand=YES)

        tb.Label(container, text="Select a function",
                 bootstyle="inverse-primary", font=("Segoe UI", 14, "bold")
        ).pack(fill=X, pady=(0,12))

        tb.Separator(container).pack(fill=X, pady=6)

        sec1 = tb.LabelFrame(container, text="1) Remove blanks", padding=12, bootstyle="secondary")
        sec1.pack(fill=X, pady=8)
        tb.Label(sec1, text='Select the file which you want to remove spaces.').pack(anchor=W, pady=(0,6))
        tb.Button(sec1, text="Remove Spaces", bootstyle="success", width=22, command=self.on_remove).pack()

        sec2 = tb.LabelFrame(container, text='2) Formula → Value', padding=12, bootstyle="secondary")
        sec2.pack(fill=X, pady=8)
        tb.Button(sec2, text="Convert File", bootstyle="primary", width=22, command=self.on_convert).pack()

        sec3 = tb.LabelFrame(container, text="3) Merge worksheets", padding=12, bootstyle="secondary")
        sec3.pack(fill=X, pady=8)
        tb.Button(sec3, text="Merge Files", bootstyle="info", width=22, command=self.on_merge).pack()

        sec4 = tb.LabelFrame(container, text="4) Remove formulas on Excel cell", padding=12, bootstyle="secondary")
        sec4.pack(fill=X, pady=8)
        tb.Button(sec4, text="Formula Cleaner", bootstyle="warning", width=22, command=self.open_formula_cleaner).pack()

        tb.Separator(container).pack(fill=X, pady=10)
        
    def on_remove(self):
        src = filedialog.askopenfilename(title="Select file", filetypes=[("Excel","*.xlsx *.xlsm *.xls")])
        if not src: return
        out = unique_path(default_save_dir() / "AfterRemove.xlsx")
        res = remove_spaces_first_col_and_drop_blank_rows(src, str(out), start_row=3)
        messagebox.showinfo("Hint", f"Finished!\nSaved: {res}")

    def on_convert(self):
        src = filedialog.askopenfilename(title="Select file", filetypes=[("Excel","*.xlsx *.xlsm *.xls")])
        if not src: return
        out = unique_path(default_save_dir() / "AfterConvert.xlsx")
        res = convert_formulas_to_values(src, str(out))
        messagebox.showinfo("Hint", f"Finished!\nSaved: {res}")

    def on_merge(self):
        folder = filedialog.askdirectory(title="Select folder")
        if not folder: return
        results = merge_many_files(folder, column_names=COLUMN_NAMES)
        messagebox.showinfo("Hint", "Finished!\n" + ("\n".join(Path(p).name for p in results) if results else "No Excel files."))

    def open_formula_cleaner(self):
        ExcelFormulaRemover(master=self.root)
