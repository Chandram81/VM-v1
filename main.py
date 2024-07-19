import tkinter as tk
from tkinter import ttk
from csv_to_excel_tab import CsvToExcelTab
from excel_sheet_extractor_tab import ExcelSheetExtractorTab

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("File Converter")

        # Create a notebook (tabs container)
        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill='both')

        # Add CSV to Excel tab
        csv_to_excel_tab = CsvToExcelTab(notebook)
        notebook.add(csv_to_excel_tab.frame, text="CSV to Excel")

        # Add Excel Sheet Extractor tab
        excel_sheet_extractor_tab = ExcelSheetExtractorTab(notebook)
        notebook.add(excel_sheet_extractor_tab.frame, text="Excel Sheet Extractor")

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
