import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter.simpledialog import askstring, askinteger

def select_xlsx_files():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    return root, file_paths

def load_sheets(file_paths):
    all_sheets = {}
    for file_path in file_paths:
        xls = pd.ExcelFile(file_path)
        sheet_names = [name for name in xls.sheet_names if name != "Commission"]
        all_sheets[file_path] = sheet_names
    return all_sheets

def user_select_sheets(all_sheets):
    selected_sheets = {}
    for file_path, sheet_names in all_sheets.items():
        print(f"\nFile: {file_path}")
        for i, name in enumerate(sheet_names, start=1):
            print(f"{i}. {name}")
        selection = input("Enter sheet numbers to process (comma-separated), 'a' for all, or 'skip' to skip this file: ")
        if selection.lower() in ['a', 'all']:
            selected_sheets[file_path] = sheet_names
        elif selection.lower() == 'skip':
            continue
        else:
            selected_numbers = [int(num) for num in selection.split(',') if num.isdigit()]
            selected_sheets[file_path] = [sheet_names[i-1] for i in selected_numbers if i-1 < len(sheet_names)]
    return selected_sheets

def sum_no_of_nights(selected_sheets):
    results = []
    for file_path, sheet_names in selected_sheets.items():
        for sheet_name in sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            if "No of Nights" in df.columns:
                sum_nights = df["No of Nights"].sum()
                print(f"Sheet: {sheet_name}, Sum of No of Nights: {sum_nights}")
                results.append((sheet_name, "No of Nights", sum_nights))
    return results

def save_results_as_csv(results):
    filename = askstring("Save File", "Enter filename for CSV (without extension): ")
    if filename:
        df = pd.DataFrame(results, columns=["Sheet Name", "Column Name", "Sum"])
        df.to_csv(f"{filename}.csv", index=False)
        print(f"Results saved to {filename}.csv")
    else:
        print("No filename provided, not saving.")

if __name__ == "__main__":
    root, file_paths = select_xlsx_files()
    if not file_paths:
        print("No files selected.")
    else:
        all_sheets = load_sheets(file_paths)
        selected_sheets = user_select_sheets(all_sheets)
        results = sum_no_of_nights(selected_sheets)
        if results:
            save_results_as_csv(results)
        root.destroy()
