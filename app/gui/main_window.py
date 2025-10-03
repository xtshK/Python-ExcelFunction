import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

from app.config import APP_TITLE, default_save_dir
from app.services.excel_utils import (
    unique_path,
    merge_many_files,
    remove_spaces_first_col_and_drop_blank_rows,
    convert_formulas_to_values,
)

COLUMN_NAMES = ['Year','Month','Country','Model-Type','Model-Size','Model','Currency','Column1','Model Price','Column3','Forecase Qty']

class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self._center(340, 440)

        font_style = ('Helvetica', 12, 'bold')
        tk.Label(root, text='Select a function', font=font_style).pack(pady=15)

        # Remove
        tk.Label(root, text='1. Select the file \n which you want to remove spaces.').pack(pady=2)
        tk.Button(root, text="Remove Spaces", command=self.on_remove, width=18).pack(pady=10)

        # Convert
        tk.Label(root, text='2. Convert cells from "Formula" to "Value".').pack(pady=2)
        tk.Button(root, text="Convert File", command=self.on_convert, width=18).pack(pady=10)

        # Merge
        tk.Label(root, text='3. Combine the Allocate Files \n worksheets into one worksheet.').pack(pady=2)
        tk.Button(root, text="Merge Files", command=self.on_merge, width=18).pack(pady=10)

        tk.Button(root, text='Exit', command=self.root.destroy, width=8).pack(pady=14)

    def _center(self, w: int, h: int):
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = (sw - w)//2, (sh - h)//2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def on_remove(self):
        src = filedialog.askopenfilename(
            title="Select the FileInOne_Result file.",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]
        )
        if not src:
            return
        out_dir = default_save_dir()
        out = unique_path(out_dir / "AfterRemove.xlsx")
        try:
            res = remove_spaces_first_col_and_drop_blank_rows(src, str(out), start_row=3)
            messagebox.showinfo("Hint", f"Finished!\nSaved: {res}")
        except Exception as e:
            messagebox.showerror("Error", f"Remove failed:\n{e}")

    def on_convert(self):
        src = filedialog.askopenfilename(
            title="Select the file to convert to values.",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]
        )
        if not src:
            return
        out_dir = default_save_dir()
        out = unique_path(out_dir / "AfterConvert.xlsx")
        try:
            res = convert_formulas_to_values(src, str(out))
            messagebox.showinfo("Hint", f"Finished!\nSaved: {res}")
        except Exception as e:
            messagebox.showerror("Error", f"Convert failed:\n{e}")

    def on_merge(self):
        folder = filedialog.askdirectory(title="Select the AfertFill_Allocate folder.")
        if not folder:
            return
        try:
            results = merge_many_files(folder, column_names=COLUMN_NAMES)
            if results:
                names = "\n".join(Path(p).name for p in results)
                messagebox.showinfo("Hint", f"Finished!\nCreated:\n{names}")
            else:
                messagebox.showinfo("Hint", "No Excel files found in the selected folder.")
        except Exception as e:
            messagebox.showerror("Error", f"Merge failed:\n{e}")
