import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime, timedelta
from tkinter import ttk


class CSVExtractorTool:
    def __init__(self, master):
        self.master = master
        master.title("CSV Data Extractor")

        self.csv_file_path = ""
        self.excel_file_path = ""

        # UI Design
        self.label_csv = tk.Label(master, text="Select CSV File:")
        self.label_csv.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        self.button_browse_csv = tk.Button(master, text="Browse", command=self.browse_csv)
        self.button_browse_csv.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        self.label_excel = tk.Label(master, text="Select Excel File:")
        self.label_excel.grid(row=1, column=0, padx=10, pady=10, sticky="w")

        self.button_browse_excel = tk.Button(master, text="Browse", command=self.browse_excel)
        self.button_browse_excel.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        self.button_extract = tk.Button(master, text="Extract Data", command=self.extract_data)
        self.button_extract.grid(row=2, column=0, columnspan=2, pady=20)

        self.progress_bar = ttk.Progressbar(master, mode='indeterminate')
        self.progress_bar.grid(row=3, column=0, columnspan=2, pady=10)

    def browse_csv(self):
        self.csv_file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        print("Selected CSV File:", self.csv_file_path)

    def browse_excel(self):
        self.excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        print("Selected Excel File:", self.excel_file_path)

    def extract_data(self):
        if not self.csv_file_path or not self.excel_file_path:
            print("Please select both CSV and Excel files.")
            return

        self.progress_bar.start()

        try:
            # Read CSV file without header
            df = pd.read_csv(self.csv_file_path, header=None)

            # Assume the first column is timestamp and the second column is generation
            timestamp_column = df.iloc[:, 0]

            # Convert timestamp to datetime type
            timestamp_column = pd.to_datetime(timestamp_column)

            # Extract data for the last 30 days from today
            today = datetime.now().date()
            start_date = today.replace(day=1) - timedelta(days=30)
            filtered_data = df[timestamp_column.dt.date >= start_date]

            print("Filtered Data:")
            print(filtered_data)

            # Write to Excel file
            with pd.ExcelWriter(self.excel_file_path, engine='openpyxl', mode='a') as writer:
                filtered_data.to_excel(writer, sheet_name='Sheet1', index=False, header=False, columns=[0, 1], startrow=0)

            print("Data extraction completed.")

        except Exception as e:
            print(f"Error during data extraction: {e}")

        self.progress_bar.stop()


if __name__ == "__main__":
    root = tk.Tk()
    app = CSVExtractorTool(root)
    root.mainloop()
