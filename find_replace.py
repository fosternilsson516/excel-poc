import pandas as pd
import tkinter as tk
from tkinter import filedialog

class FindReplace(tk.Frame):
    def __init__(self, master=None, excel_handler=None):
        super().__init__(master)
        self.excel_handler = excel_handler
        self.create_widgets()
        self.set_df = None
        self.template_df = None

    def create_widgets(self):
        label = tk.Label(self, text="Make Sure There Is Only One Sheet In Each File", font=("Helvetica", 10))
        label.place(x=100, y=15)

        label = tk.Label(self, text="1st: Select Template Find And Replace File", font=("Helvetica", 12))
        label.place(x=105, y=50)

        # Add button to open Excel file in Tab One
        button = tk.Button(self, text="Connect & Load Template File", command=self.template_excel_file)
        button.place(x=166, y=90)

        label = tk.Label(self, text="2nd: Select File That Needs Characters Fixed", font=("Helvetica", 12))
        label.place(x=100, y=140)

        # Add button to open Excel file in Tab One
        button = tk.Button(self, text="Connect & Load Working File", command=self.set_excel_df)
        button.place(x=167, y=180)

    def set_excel_df(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if file_path:
            success = self.excel_handler.open_excel_file(file_path)
            if success:  
                self.set_df = self.excel_handler.get_selected_data() 
                if self.set_df is not None: 
                    print("DataFrame Set:")
                    print(self.set_df) 
                    self.process_selected_data(self.template_df)
                    print(file_path)     

    def template_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if file_path:
            success = self.excel_handler.open_excel_file(file_path)
            self.template_df = self.excel_handler.get_selected_data()
            if self.template_df is not None:
                print("Template DataFrame:")
                print(self.template_df)
                 
                             
                  

    def process_selected_data(self, template_df):
        if self.set_df is not None and self.template_df is not None:
            # Create a dictionary from "Find" and "Replace" columns in the template
            find_replace_dict = self.create_find_replace_dict(self.template_df)

            for find_value, replace_value in find_replace_dict.items():
                # Replace values in the set_df
                self.set_df = self.replace_values(self.set_df, find_value, replace_value)

            print("DataFrame after replacement:")
            print(self.set_df)  
            self.excel_handler.update_sheet_with_dataframe(self.set_df, new_sheet_name="Cleaned Data")    

    def create_find_replace_dict(self, df):
        find_replace_dict = dict(zip(df.iloc[:, 0], df.iloc[:, 1]))
        return find_replace_dict

    def replace_values(self, df, find_value, replace_value):
        # Apply replacements to the entire DataFrame
        return df.map(lambda entry: self.replace_entry(entry, find_value, replace_value))

    def replace_entry(self, entry, find_value, replace_value):
        # Check for both values and break into characters for further processing
        if isinstance(entry, str):
            # Replace entire values
            entry = entry.replace(find_value, replace_value)

            # Break into characters and check for individual characters
            entry = "".join([self.replace_char(char, find_value, replace_value) for char in entry])

        return entry

    def replace_char(self, char, find_value, replace_value):
        # Replace based on the dictionary for individual characters
        return replace_value if char == find_value else char    