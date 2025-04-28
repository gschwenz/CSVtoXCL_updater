import pandas as pd
import os
from tkinter import Tk, filedialog
from openpyxl import load_workbook

# --- HIDE TKINTER ROOT WINDOW ---
root = Tk()
root.withdraw()

# --- SELECT CSV FILE ---
csv_path = filedialog.askopenfilename(
    title="Wähle die CSV-Datei aus",
    filetypes=[("CSV files", "*.csv")])

if not csv_path:
    input("❌ Keine CSV-Datei ausgewählt. Drücke ENTER zum Beenden...")
    raise SystemExit()

# --- SELECT EXCEL FILE ---
sales_path = filedialog.askopenfilename(
    title="Wähle die Excel-Datei aus, zu der Daten hinzugefügt werden sollen",
    filetypes=[("Excel files", "*.xlsx")])

if not sales_path:
    input("❌ Keine Excel-Datei ausgewählt. Drücke ENTER zum Beenden...")
    raise SystemExit()

print(f"\n📄 CSV-Datei: {csv_path}")
print(f"📊 Ziel-Excel-Datei: {sales_path}")

# --- READ DATA ---
csv_df = pd.read_csv(csv_path)
sheet_name = pd.ExcelFile(sales_path).sheet_names[0]
existing_df = pd.read_excel(sales_path, sheet_name=sheet_name)

# --- APPEND DATA ---
updated_df = pd.concat([existing_df, csv_df], ignore_index=True)

# --- WRITE BACK TO EXCEL ---
with pd.ExcelWriter(sales_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    updated_df.to_excel(writer, sheet_name=sheet_name, index=False)

print("\n✅ Daten erfolgreich an die Excel-Datei angehängt.")
input("🔚 Drücke ENTER zum Beenden...")
