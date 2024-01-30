import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class VlookUp(tk.Frame):
    def __init__(self, master=None, excel_handler=None):
        super().__init__(master)
        self.excel_handler = excel_handler
        self.radio_selection = tk.StringVar()
        self.radio_selection.set(None)
        self.create_widgets()
        self.excel_df = None
        self.list_df = None
        
    def create_widgets(self):
        label = tk.Label(self, text="Make Sure To Match The Column Headers In Both Files", font=("Helvetica", 10))
        label.place(x=80, y=15)

        label = tk.Label(self, text="1st: Select File That Has The List To Lookup", font=("Helvetica", 12))
        label.place(x=100, y=50)

        # Add button to open Excel file in Tab One
        button = tk.Button(self, text="Connect & Load Lookup File", command=self.lookup_list_df)
        button.place(x=167, y=85)

        label = tk.Label(self, text="2nd: Check Box To Keep Rows From List Or Remove Matching Rows", font=("Helvetica", 12))
        label.place(x=10, y=125)

        radio_button1 = tk.Radiobutton(self, text="Keep Rows", variable=self.radio_selection, value="Keep", command=self.on_radio_change)
        radio_button1.place(x=250, y=170)

        radio_button2 = tk.Radiobutton(self, text="Remove Rows", variable=self.radio_selection, value="Remove", command=self.on_radio_change)
        radio_button2.place(x=140, y=170)

        label = tk.Label(self, text="3rd: Select Working File That Needs To Be Cleaned", font=("Helvetica", 12))
        label.place(x=80, y=220)

        # Add button to open Excel file in Tab One
        button = tk.Button(self, text="Connect & Load Working File", command=self.working_excel_df)
        button.place(x=167, y=255)

    def on_radio_change(self):
        print("Selected option:", self.radio_selection.get())    

    def working_excel_df(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if file_path:
            success = self.excel_handler.open_excel_file(file_path)
            if success:  
                self.excel_df = self.excel_handler.get_selected_data() 
                if self.excel_df is not None: 
                    print("Working Excel DataFrame:")
                    print(self.excel_df) 
                    self.process_selected_data(self.list_df, self.radio_selection.get())
                    print(file_path)     

    def lookup_list_df(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
        if file_path:
            success = self.excel_handler.open_excel_file(file_path)
            if success:
                self.list_df = self.excel_handler.get_selected_data()
                if self.list_df is not None:
                    print("List DataFrame:")
                    print(self.list_df)
                 
    def process_selected_data(self, list_df, radio_selection):
        if self.excel_df is not None and self.list_df is not None:

            list_column_header = list_df.iloc[0, 0].strip() if not list_df.empty else None

            if list_column_header is None:
                tkinter.messagebox.showerror("Error", "Make sure column header is in the very first cell")
                return


            common_column = None
            for idx, column in enumerate(self.excel_df.iloc[0]):
                if str(column).strip() == list_column_header:
                    common_column = idx
                    break

            if common_column is None:
                return    

            list_values = set(list_df.iloc[1:, 0].astype(str).tolist())      

            condition = self.excel_df[common_column].astype(str).isin(list_values) if radio_selection == "Keep" else ~self.excel_df[common_column].astype(str).isin(list_values)
            
            self.excel_df = self.excel_df[condition]

            self.excel_df = self.excel_df.reset_index(drop=True)

            print("DataFrame after processing selected data:")
            print(self.excel_df)
            self.excel_handler.update_sheet_with_dataframe(self.excel_df, new_sheet_name="Cleaned Data")