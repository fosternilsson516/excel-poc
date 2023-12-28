import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlwings as xw
from excel_handler import ExcelHandler
import os
import pandas as pd
from find_replace import FindReplace


class TabTwo(tk.Frame):
    def __init__(self, master=None, excel_handler=None):
        super().__init__(master)
        self.excel_handler = excel_handler
        self.create_widgets()

    def create_widgets(self):
        label = tk.Label(self, text="This is Tab 2")
        label.pack(padx=20, pady=20)

        # Add button to open Excel file in Tab Two
        button = tk.Button(self, text="Open Excel File", command=self.open_excel_file_dialog)
        button.pack(pady=10)

    def open_excel_file_dialog(self):
        if self.excel_handler:
            self.excel_handler.open_excel_file_dialog()
            self.get_and_print_selected_data()  

    def get_and_print_selected_data(self):
        if self.excel_handler:
            df = self.excel_handler.get_selected_data()
            if df is not None:
                print("Selected Data:")
                print(df)

                # You can save the DataFrame to a file or perform other operations here
                self.save_data_to_csv(df)          


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
        self.tab2 = TabTwo(self.notebook, self.excel_handler)

        self.notebook.add(self.tab1, text="Find & Replace Char")
        self.notebook.add(self.tab2, text="Tab 2")

        self.notebook.pack(expand=1, fill="both")

if __name__ == "__main__":
    app = MainUI()
    app.mainloop()
