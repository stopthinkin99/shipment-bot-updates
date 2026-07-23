"""Background FedEx status scheduler and Excel updater."""
from __future__ import annotations
import json, os, re, threading, time
from datetime import date, datetime
from pathlib import Path
import openpyxl
from fedex_credentials import has_credentials
from fedex_tracking import MAX_TRACKING_NUMBERS_PER_REQUEST, track_numbers

SHEET_TO_PROFILE={"UNI":"UNI","FENIX":"FENIX"}
COL_DATE=1; COL_CARRIER=4; COL_TRACKING=5; COL_REMARK=7
FEDEX_RE=re.compile(r"\b(?:FEDEX|FED\s*EX|BX\s*FX|FX\s*S/O|FX\s*P/O|STANDARD\s+OVERNIGHT|PRIORITY\s+OVERNIGHT|FIRST\s+OVERNIGHT)\b",re.I)

def _settings_path():
    base=os.environ.get("LOCALAPPDATA"); folder=Path(base)/"UniCreation"/"ShipmentBot" if base else Path.home()/".uni_creation_shipment_bot"
    folder.mkdir(parents=True,exist_ok=True); return folder/"fedex_tracking_settings.json"
SETTINGS=_settings_path()
def load_tracking_time(default="16:00"):
    try: return str(json.loads(SETTINGS.read_text(encoding="utf-8")).get("fedex_tracking_time",default)).strip() or default
    except Exception: return default
def save_tracking_time(value):
    try: SETTINGS.write_text(json.dumps({"fedex_tracking_time":str(value).strip()},indent=2),encoding="utf-8")
    except Exception: pass

def _as_date(value):
    if isinstance(value,datetime): return value.date()
    if isinstance(value,date): return value
    for fmt in ("%m/%d/%Y","%m/%d/%y","%Y-%m-%d","%d%b%y","%d%b%Y"):
        try: return datetime.strptime(str(value).strip(),fmt).date()
        except Exception: pass
    return None

def _tracking(value): return "".join(ch for ch in str(value or "") if ch.isalnum())
def _batches(values,size):
    for i in range(0,len(values),size): yield values[i:i+size]

def update_fedex_statuses(excel_path, *, environment="production", today=None, log=print):
    today=today or date.today()
    if not excel_path or not os.path.isfile(excel_path):
        msg=f"[FEDEX] Workbook not found: {excel_path}"; log(msg); return {"updated":0,"unchanged":0,"failed":0,"message":msg}
    log("[FEDEX] Starting shipment-status update.")
    log(f"[FEDEX] Workbook: {excel_path}")
    log(f"[FEDEX] Environment: {environment.upper()}")
    log(
        f"[FEDEX] Tracking current month only: "
        f"{today.strftime('%B %Y')}."
    )
    log(
        "[FEDEX] Rows already marked DELIVERED will be skipped."
    )
    wb=None
    for attempt in range(5):
        try: wb=openpyxl.load_workbook(excel_path); break
        except (PermissionError,OSError):
            if attempt==4: raise PermissionError("Close the workbook in Excel and try again.")
            time.sleep(2)
    changed=False; updated=unchanged=failed=0
    month_start = today.replace(day=1)
    try:
        for sheet_name,profile in SHEET_TO_PROFILE.items():
            actual=next((s for s in wb.sheetnames if s.strip().upper()==sheet_name),None)
            if not actual: log(f"[FEDEX] {sheet_name}: worksheet not found; skipped."); continue
            ws=wb[actual]; rows=[]
            for row in range(1,ws.max_row+1):
                carrier=ws.cell(row,COL_CARRIER).value; number=_tracking(ws.cell(row,COL_TRACKING).value); status=str(ws.cell(row,COL_REMARK).value or "").strip().upper(); d=_as_date(ws.cell(row,COL_DATE).value)
                # Track only FedEx rows from the current calendar month.
                # Any row already marked Delivered is final and skipped.
                if not FEDEX_RE.search(str(carrier or "")):
                    continue
                if not number:
                    continue
                if status == "DELIVERED":
                    continue
                if d is None:
                    log(
                        f"[FEDEX] {sheet_name} row {row}: skipped because "
                        "the shipment date is missing or invalid."
                    )
                    continue
                if d.year != today.year or d.month != today.month:
                    continue
                rows.append((row,number,status))
            log(f"[FEDEX] {sheet_name}: found {len(rows)} eligible shipment(s).")
            if not rows: continue
            if not has_credentials(profile): failed+=len(rows); log(f"[FEDEX] {profile}: credentials missing; rows skipped."); continue
            result_map={}; unique=list(dict.fromkeys(n for _,n,_ in rows))
            for i,batch in enumerate(_batches(unique,MAX_TRACKING_NUMBERS_PER_REQUEST),1):
                try:
                    log(f"[FEDEX] {profile}: tracking batch {i} containing {len(batch)} number(s).")
                    result_map.update(track_numbers(profile,batch,environment=environment,log=log))
                except Exception as exc:
                    failed+=len(batch); log(f"[FEDEX] {profile}: batch {i} failed: {type(exc).__name__}: {exc}")
            for row,number,old in rows:
                result=result_map.get(number)
                if not result: continue
                status=str(result.get("status") or "UNKNOWN").strip().upper()
                if status==old: unchanged+=1; log(f"[FEDEX] {sheet_name} row {row}: {number} remains {status}"); continue
                ws.cell(row,COL_REMARK).value=status; changed=True; updated+=1; log(f"[FEDEX] {sheet_name} row {row}: {number} → {status}")
        if changed: wb.save(excel_path); log("[FEDEX] Workbook saved successfully.")
        else: log("[FEDEX] No workbook changes were required.")
        msg=f"[FEDEX] Finished: updated={updated}, unchanged={unchanged}, failed={failed}."; log(msg)
        return {"updated":updated,"unchanged":unchanged,"failed":failed,"message":msg}
    finally:
        if wb: wb.close()

class FedExStatusScheduler(threading.Thread):
    def __init__(self,get_excel_path,get_target,log=print,*,environment="production",poll_seconds=20):
        super().__init__(daemon=True); self.get_excel_path=get_excel_path; self.get_target=get_target; self.log=log; self.environment=environment; self.poll_seconds=poll_seconds; self._stop=threading.Event(); self._last=None
    def stop(self): self._stop.set()
    def run(self):
        self.log("FedEx status scheduler started.")
        while not self._stop.is_set():
            now=datetime.now(); target=str(self.get_target() or "").strip()
            if target and now.strftime("%H:%M")==target and self._last!=now.date():
                self._last=now.date(); path=str(self.get_excel_path() or "").strip()
                if path and os.path.isfile(path):
                    self.log("[FEDEX] Running scheduled status update..."); update_fedex_statuses(path,environment=self.environment,today=now.date(),log=self.log)
                else: self.log("[FEDEX] Scheduled update skipped: Excel path is missing.")
            self._stop.wait(self.poll_seconds)
        self.log("FedEx status scheduler stopped.")