import os
import time
import shutil
from parser import parse_label
from excel_writer import update_excel
from datetime import datetime

WATCH_FOLDER  = r"C:\Users\aayan.boradia\Downloads\Labels\Incoming"
DONE_FOLDER   = r"C:\Users\aayan.boradia\Downloads\Labels\Done"
EXCEL_PATH = r"C:\Users\aayan.boradia\Downloads\TRACKING_SHIPMENT.xlsx"


def process_file(file_path, excel_path=None, done_folder=None):
    # Use app-provided paths if given, else fall back to defaults
    _excel = excel_path or r"C:\Users\aayan.boradia\Downloads\TRACKING_SHIPMENT.xlsx"
    _done  = done_folder or r"C:\Users\aayan.boradia\Downloads\ShipmentLabels\Done"
    
    print(f"[INFO] Processing: {file_path}")
    try:
        records = parse_label(file_path)

        if not records:
            print(f"[WARN] No records extracted from {file_path}")
            return

        for record in records:
            if not record.get("sheet"):
                print(f"[SKIP] Could not determine sheet for invoice '{record.get('invoice')}' — manual review needed")
                continue

            if not record.get("tracking_number"):
                print(f"[WARN] No tracking number found in {file_path}")

            update_excel(_excel, record)
            print(f"[OK] Written to sheet '{record['sheet']}': {record}")

        os.makedirs(_done, exist_ok=True)
        shutil.move(file_path, os.path.join(_done, os.path.basename(file_path)))
        print(f"[DONE] Moved to {_done}")

    except Exception as e:
        print(f"[ERROR] Failed on {file_path}: {e}")

# Alias
process_label = process_file