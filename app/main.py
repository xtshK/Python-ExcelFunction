import tkinter as tk
import ttkbootstrap as tb
from app.gui.ui_beautified import MainWindowBootstrap

def main():
    root = tb.Window(themename="darkly")
    root.title("CSC 4M Excel Formatter")
    MainWindowBootstrap(root)
    root.mainloop()