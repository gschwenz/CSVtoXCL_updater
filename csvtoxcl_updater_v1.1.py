import pandas as pd
import os
from tkinter import Tk, filedialog
from openpyxl import load_workbook

# --- HIDE TKINTER ROOT WINDOW ---
root = Tk()
root.withdraw()

# --- SELECT CSV FILE ---
csv_path = filedialog.askopenfilename(
    title="Wähle die CSV-Datei aus\\Select the CSV file",
    filetypes=[("CSV files", "*.csv")])

if not csv_path:
    raise FileNotFoundError("Keine CSV-Datei ausgewählt!")

# --- SELECT EXCEL FILE ---
sales_path = filedialog.askopenfilename(
    title="Wähle die Excel-Datei aus, zu der Daten hinzugefügt werden sollen\\Select the Excel file to which you want to add data",
    filetypes=[("Excel files", "*.xlsx")])

if not sales_path:
    raise FileNotFoundError("Keine Excel-Datei ausgewählt!\nNo Excel file selected!")

print(f"CSV-Datei\\CSV File: {csv_path}")
print(f"Ziel-Excel-Datei\\Target Excel File: {sales_path}")

# --- READ DATA ---
csv_df = pd.read_csv(csv_path)
sheet_name = pd.ExcelFile(sales_path).sheet_names[0]  # get first sheet name
existing_df = pd.read_excel(sales_path, sheet_name=sheet_name)

# --- APPEND DATA ---
updated_df = pd.concat([existing_df, csv_df], ignore_index=True)

# --- WRITE BACK WITHOUT TOUCHING .book/.sheets ---
with pd.ExcelWriter(sales_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    updated_df.to_excel(writer, sheet_name=sheet_name, index=False)

print("✅ Daten erfolgreich an die Excel-Datei angehängt.\n✅ Data successfully appended to the Excel file.")
input("Press Enter to close...")
