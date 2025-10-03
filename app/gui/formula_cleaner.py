# app/gui/formula_cleaner.py
import os, platform, tkinter as tk
from tkinter import filedialog, messagebox
from app.services.excel_utils import clean_excel_file_to_values

# macOS 的 Tk 提示抑制（可留在這裡）
if platform.system() == 'Darwin':
    os.environ['TK_SILENCE_DEPRECATION'] = '1'

class ExcelFormulaRemover(tk.Toplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Excel Clean System")
        self.geometry("400x200")

        tk.Button(self, text="Upload Excel File",
                  command=self.process_file, font=("Arial", 12)
        ).pack(expand=True, pady=50)

    def process_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel file to clean",
            filetypes=[
                ("Excel Files", "*.xlsx *.xls *.xlsm"),
                ("Excel 2007+", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("All Files", "*.*")
            ]
        )
        if not file_path:
            return
        try:
            out_path = clean_excel_file_to_values(file_path)  # 可加 out_dir=...
            messagebox.showinfo("Success", f"Formulas removed and file saved to:\n{out_path}")
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred:\n{e}")
