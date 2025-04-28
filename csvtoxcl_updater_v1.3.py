import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox
from openpyxl import load_workbook
from datetime import datetime

try:
    # --- HIDE TKINTER ROOT WINDOW ---
    root = Tk()
    root.withdraw()

    # --- SELECT CSV FILE ---
    csv_path = filedialog.askopenfilename(
        title="Wähle die CSV-Datei aus",
        filetypes=[("CSV files", "*.csv")])

    if not csv_path:
        messagebox.showwarning("Abgebrochen", "❌ Keine CSV-Datei ausgewählt.")
        raise SystemExit()

    # --- SELECT EXCEL FILE ---
    sales_path = filedialog.askopenfilename(
        title="Wähle die Excel-Datei aus, zu der Daten hinzugefügt werden sollen",
        filetypes=[("Excel files", "*.xlsx")])

    if not sales_path:
        messagebox.showwarning("Abgebrochen", "❌ Keine Excel-Datei ausgewählt.")
        raise SystemExit()

    # --- LOAD DATA ---
    csv_df = pd.read_csv(csv_path)
    sheet_name = pd.ExcelFile(sales_path).sheet_names[0]
    existing_df = pd.read_excel(sales_path, sheet_name=sheet_name)

    # --- APPEND DATA ---
    updated_df = pd.concat([existing_df, csv_df], ignore_index=True)
    added_rows = len(csv_df)

    # --- WRITE TO EXCEL ---
    with pd.ExcelWriter(sales_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        updated_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # --- LOGGING ---
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_entry = (f"{timestamp} | ✅ {added_rows} Zeilen von '{os.path.basename(csv_path)}' "
                 f"zu '{os.path.basename(sales_path)}' (Sheet: {sheet_name}) hinzugefügt.\n")

    log_path = os.path.join(os.path.dirname(__file__), "import_log.txt")
    with open(log_path, "a", encoding="utf-8") as log_file:
        log_file.write(log_entry)

    # --- GUI CONFIRMATION POPUP ---
    messagebox.showinfo(
        "Import abgeschlossen",
        f"{added_rows} Zeilen wurden erfolgreich angehängt.\n\n"
        f"📄 CSV-Datei: {os.path.basename(csv_path)}\n"
        f"📊 Ziel-Excel-Datei: {os.path.basename(sales_path)}\n"
        f"📈 Arbeitsblatt: {sheet_name}\n\n"
        f"(Eintrag im Protokoll gespeichert)"
    )

except Exception as e:
    import traceback
    print("\n❌ Ein Fehler ist aufgetreten:")
    traceback.print_exc()
    input("🔚 Drücke ENTER zum Beenden...")
