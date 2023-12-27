import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlwings as xw
from excel_handler import ExcelHandler
import os
from temp_table_manager import TempTableManager

class TabOne(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.create_widgets()

    def create_widgets(self):
        label = tk.Label(self, text="This is Tab 1")
        label.pack(padx=20, pady=20)

class TabTwo(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.create_widgets()

    def create_widgets(self):
        label = tk.Label(self, text="This is Tab 2")
        label.pack(padx=20, pady=20)


class MainUI(tk.Tk):
    def __init__(self):
        super().__init__()
        # Configure the main window
        self.window.title("Excel Interaction UI")
        self.window.geometry("600x450")

        self.notebook = ttk.Notebook(self)

        # Create tabs
        self.tab1 = TabOne(self.notebook)
        self.tab2 = TabTwo(self.notebook)

        self.notebook.add(self.tab1, text="Tab 1")
        self.notebook.add(self.tab2, text="Tab 2")

        self.notebook.pack(expand=1, fill="both")

if __name__ == "__main__":
    app = MainUI()
    app.mainloop()
