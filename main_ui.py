import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlwings as xw
from excel_handler import ExcelHandler
import pandas as pd
from find_replace import FindReplace
from vlookup import VlookUp

class MainUI(tk.Tk):
    def __init__(self):
        super().__init__()
        # Configure the main window
        self.title("Excel Interaction UI")
        self.geometry("500x350")

        self.notebook = ttk.Notebook(self)

        self.excel_handler = ExcelHandler()

        # Create tabs
        self.tab1 = FindReplace(self.notebook, self.excel_handler)
        self.tab2 = VlookUp(self.notebook, self.excel_handler)

        self.notebook.add(self.tab1, text="Find & Replace Char")
        self.notebook.add(self.tab2, text="Vlookup")

        self.notebook.pack(expand=1, fill="both")

if __name__ == "__main__":
    app = MainUI()
    app.mainloop()
