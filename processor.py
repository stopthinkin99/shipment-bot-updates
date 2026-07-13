import os
import time
import shutil
from parser import parse_label
from excel_writer import update_excel
from datetime import datetime

WATCH_FOLDER  = r"C:\Users\aayan.boradia\Downloads\Labels\Incoming"
DONE_FOLDER   = r"C:\Users\aayan.boradia\Downloads\Labels\Done"
EXCEL_PATH    = r"C:\Users\aayan.boradia\Downloads\Labels\Template_label.xlsx"


def process_file(file_path):
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

            update_excel(EXCEL_PATH, record)
            print(f"[OK] Written to sheet '{record['sheet']}': {record}")

        # Move to done
        os.makedirs(DONE_FOLDER, exist_ok=True)
        shutil.move(file_path, os.path.join(DONE_FOLDER, os.path.basename(file_path)))
        print(f"[DONE] Moved to {DONE_FOLDER}")

    except Exception as e:
        print(f"[ERROR] Failed on {file_path}: {e}")