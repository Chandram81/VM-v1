import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from excel_utils import load_sheet_names, load_columns, combine_sheets
import os

class ExcelSheetExtractorTab:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)

        file_label = ttk.Label(self.frame, text="Excel File:")
        file_label.grid(row=0, column=0, padx=10, pady=10)

        self.file_entry = ttk.Entry(self.frame, width=50)
        self.file_entry.grid(row=0, column=1, padx=10, pady=10)

        browse_button = ttk.Button(self.frame, text="Browse", command=self.browse_file)
        browse_button.grid(row=0, column=2, padx=10, pady=10)

        sheet_label = ttk.Label(self.frame, text="Select Sheets:")
        sheet_label.grid(row=1, column=0, padx=10, pady=10)

        self.sheet_frame = ttk.Frame(self.frame)
        self.sheet_frame.grid(row=1, column=1, padx=10, pady=10)

        self.extract_button = ttk.Button(self.frame, text="Combine Selected Columns", command=self.combine_sheets)
        self.extract_button.grid(row=3, columnspan=3, pady=20)

        self.sheet_checkbuttons = []
        self.column_checkbuttons = {}
        self.current_sheet_vars = None

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.load_sheet_names(file_path)

    def load_sheet_names(self, file_path):
        try:
            sheet_names = load_sheet_names(file_path)

            for widget, _ in self.sheet_checkbuttons:
                widget.destroy()
            self.sheet_checkbuttons = []

            for sheet_name in self.column_checkbuttons:
                for widget, _ in self.column_checkbuttons[sheet_name]:
                    widget.destroy()
            self.column_checkbuttons = {}

            for idx, sheet_name in enumerate(sheet_names):
                var = tk.BooleanVar()
                cb = tk.Checkbutton(self.sheet_frame, text=sheet_name, variable=var, command=lambda sn=sheet_name, v=var: self.load_columns(sn, v))
                cb.grid(row=idx, column=0, sticky='w')
                self.sheet_checkbuttons.append((cb, var))
                self.column_checkbuttons[sheet_name] = []

        except RuntimeError as e:
            messagebox.showerror("Error", str(e))

    def load_columns(self, sheet_name, sheet_var):
        if sheet_var.get():
            file_path = self.file_entry.get()
            try:
                column_names = load_columns(file_path, sheet_name)

                for widget, _ in self.column_checkbuttons[sheet_name]:
                    widget.destroy()
                self.column_checkbuttons[sheet_name] = []

                for idx, column_name in enumerate(column_names):
                    var = tk.BooleanVar()
                    cb = tk.Checkbutton(self.sheet_frame, text=column_name, variable=var)
                    cb.grid(row=idx, column=1, sticky='w')
                    self.column_checkbuttons[sheet_name].append((cb, var))

            except RuntimeError as e:
                messagebox.showerror("Error", str(e))
        else:
            for _, var in self.column_checkbuttons[sheet_name]:
                var.set(False)

    def combine_sheets(self):
        excel_file_path = self.file_entry.get()
        selected_sheets = []
        for cb, var in self.sheet_checkbuttons:
            if var.get():
                sheet_name = cb.cget('text')
                selected_columns = [(var, col.cget('text')) for col, var in self.column_checkbuttons[sheet_name] if var.get()]
                if selected_columns:
                    selected_sheets.append((sheet_name, selected_columns))

        if excel_file_path and selected_sheets:
            try:
                new_excel_file_path = combine_sheets(excel_file_path, selected_sheets)
                messagebox.showinfo("Success", f"Selected sheets successfully combined into:\n{new_excel_file_path}")
                os.startfile(os.path.dirname(new_excel_file_path))
            except RuntimeError as e:
                messagebox.showerror("Error", str(e))
        else:
            messagebox.showwarning("Missing Information", "Please select an Excel file and at least one sheet to combine.")
