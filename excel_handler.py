import tkinter as tk
from tkinter import filedialog
import xlwings as xw
import pandas as pd


class ExcelHandler():
    def __init__(self):
        self.app = None
        self.wb = None
        self.sheet = None
        self.selected_data = None
        

    def open_excel_file(self, file_path):
        try:
            self.app = xw.apps.active or xw.App(visible=True, add_book=False)
            self.wb = self.app.books.open(file_path)

            return True
        except Exception as e:
            print(f"Error interacting with Excel file: {e}")
            return False

    def close_excel_file(self):
        if self.wb:
            self.wb.close()
        if self.app:
            self.app.quit()

    def get_selected_data(self):
        if self.wb:
            # Get all the data from the active sheet
            active_sheet = self.wb.sheets.active
            # Assuming you want to get data from the used range
            used_range = active_sheet.used_range
            # Retrieve the values from the used range
            data = used_range.value

            # Convert the data into a Pandas DataFrame
            df = pd.DataFrame(data)

            # Assign the DataFrame to selected_data
            self.selected_data = df

            return df

        return None

    def update_sheet_with_dataframe(self, df):
        if self.wb:
            try:
                # Assuming you want to update the active sheet
                active_sheet = self.wb.sheets.active

                # Clear existing data in the sheet
                active_sheet.clear()

                # Write the new DataFrame to the sheet starting from cell A1
                active_sheet.range("A1").value = df.values.tolist()

                # Save the workbook to apply changes
                self.wb.save()

                return True
            except Exception as e:
                print(f"Error updating sheet: {e}")
                return False

        return False    
