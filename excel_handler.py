def open_excel_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            success = self.excel_handler.open_excel_file(file_path)
            if success:
                self.add_excel_tab(file_path)