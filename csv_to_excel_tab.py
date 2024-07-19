import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import xlsxwriter

class CsvToExcelTab:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)

        # Create the widgets for the CSV to Excel tab
        file_label = ttk.Label(self.frame, text="CSV File:")
        file_label.grid(row=0, column=0, padx=10, pady=10)

        self.file_entry = ttk.Entry(self.frame, width=50)
        self.file_entry.grid(row=0, column=1, padx=10, pady=10)

        browse_button = ttk.Button(self.frame, text="Browse", command=self.browse_file)
        browse_button.grid(row=0, column=2, padx=10, pady=10)

        convert_button = ttk.Button(self.frame, text="Convert to Excel", command=self.convert_csv_to_excel)
        convert_button.grid(row=1, columnspan=3, pady=20)

    def browse_file(self):
        # Open file dialog to select a CSV file
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def convert_csv_to_excel(self):
        csv_file_path = self.file_entry.get()
        if csv_file_path:
            try:
                # Convert the file path to an Excel file path
                excel_file_path = csv_file_path.replace('.csv', '.xlsx')

                # Create an XlsxWriter workbook and worksheet
                workbook = xlsxwriter.Workbook(excel_file_path)
                worksheet = workbook.add_worksheet()

                # Read the CSV file and write its content to the Excel worksheet
                with open(csv_file_path, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    for row_idx, row in enumerate(reader):
                        for col_idx, cell in enumerate(row):
                            worksheet.write(row_idx, col_idx, cell)

                # Close the workbook
                workbook.close()

                # Show a success message
                messagebox.showinfo("Success", f"CSV file successfully converted to Excel:\n{excel_file_path}")
            except Exception as e:
                # Show an error message
                messagebox.showerror("Error", f"An error occurred:\n{e}")
        else:
            messagebox.showwarning("No file selected", "Please select a CSV file to convert.")
