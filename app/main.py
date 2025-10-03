import tkinter as tk
from app.gui.main_window import MainWindow

def main():
    root = tk.Tk()
    root.title("File Dialog")
    app = MainWindow(root)
    root.mainloop()
