"""Official FedEx OAuth + Track API client."""
from __future__ import annotations
import threading, time, uuid
from dataclasses import dataclass
from typing import Any, Dict, Iterable
import requests
from fedex_credentials import load_credentials

PRODUCTION_BASE_URL = "https://apis.fedex.com"
SANDBOX_BASE_URL = "https://apis-sandbox.fedex.com"
DEFAULT_ENVIRONMENT = "production"
REQUEST_TIMEOUT_SECONDS = 35
MAX_TRACKING_NUMBERS_PER_REQUEST = 30

class FedExTrackingError(RuntimeError): pass
@dataclass
class _Token: value: str; expires_at: float
_cache: Dict[str,_Token] = {}; _lock = threading.Lock()

def _base_url(env: str) -> str:
    env=(env or "").strip().lower()
    if env in {"sandbox","test"}: return SANDBOX_BASE_URL
    if env in {"production","prod","live"}: return PRODUCTION_BASE_URL
    raise ValueError("environment must be production or sandbox")

def _token(profile: str, environment: str, log=print) -> str:
    profile=profile.strip().upper(); key=f"{environment}:{profile}"
    api, secret = load_credentials(profile)
    if not api or not secret: raise FedExTrackingError(f"FedEx credentials are not configured for {profile}.")
    with _lock:
        cached=_cache.get(key)
        if cached and cached.expires_at > time.time()+120: return cached.value
        log(f"[FEDEX] {profile}: requesting OAuth token.")
        r=requests.post(f"{_base_url(environment)}/oauth/token", headers={"Content-Type":"application/x-www-form-urlencoded"}, data={"grant_type":"client_credentials","client_id":api,"client_secret":secret}, timeout=REQUEST_TIMEOUT_SECONDS)
        try: payload=r.json()
        except ValueError: payload={}
        if r.status_code >= 400: raise FedExTrackingError(f"{profile} OAuth failed ({r.status_code}): {payload.get('errors') or payload.get('error_description') or r.text[:500]}")
        value=str(payload.get("access_token") or "").strip()
        if not value: raise FedExTrackingError("OAuth response had no access_token")
        _cache[key]=_Token(value, time.time()+int(payload.get("expires_in") or 3600))
        return value

def _normalize(code, description, derived):
    c=" ".join(str(x or "").upper() for x in (code,description,derived))
    if "DELIVERED" in c or str(code).upper()=="DL": return "DELIVERED"
    if "OUT FOR DELIVERY" in c or "ON FEDEX VEHICLE" in c: return "OUT FOR DELIVERY"
    if any(x in c for x in ("EXCEPTION","FAILED ATTEMPT","UNABLE TO DELIVER","DELAY")): return "DELIVERY EXCEPTION"
    if "PICKED UP" in c or "PICKUP" in c: return "PICKED UP"
    if any(x in c for x in ("IN TRANSIT","AT FEDEX FACILITY","DEPARTED FEDEX","ARRIVED AT FEDEX")): return "IN TRANSIT"
    if any(x in c for x in ("LABEL CREATED","SHIPMENT INFORMATION SENT","INFORMATION SENT")): return "LABEL CREATED"
    if "RETURN" in c: return "RETURNED"
    return str(description or derived or code or "UNKNOWN").strip().upper()[:80]

def track_numbers(profile: str, tracking_numbers: Iterable[str], *, environment="production", include_detailed_scans=False, log=print):
    numbers=[]; seen=set()
    for value in tracking_numbers:
        n="".join(ch for ch in str(value or "") if ch.isalnum())
        if n and n not in seen: seen.add(n); numbers.append(n)
    if not numbers: return {}
    if len(numbers)>30: raise ValueError("Maximum 30 tracking numbers per request")
    token=_token(profile, environment, log)
    body={"includeDetailedScans": bool(include_detailed_scans), "trackingInfo":[{"trackingNumberInfo":{"trackingNumber":n}} for n in numbers]}
    headers={"Authorization":f"Bearer {token}","Content-Type":"application/json","X-locale":"en_US","X-customer-transaction-id":str(uuid.uuid4())}
    r=requests.post(f"{_base_url(environment)}/track/v1/trackingnumbers", headers=headers, json=body, timeout=REQUEST_TIMEOUT_SECONDS)
    if r.status_code==401:
        with _lock: _cache.pop(f"{environment}:{profile.strip().upper()}",None)
        headers["Authorization"]=f"Bearer {_token(profile, environment, log)}"
        r=requests.post(f"{_base_url(environment)}/track/v1/trackingnumbers", headers=headers, json=body, timeout=REQUEST_TIMEOUT_SECONDS)
    try: payload=r.json()
    except ValueError: payload={}
    if r.status_code>=400: raise FedExTrackingError(f"Track API failed ({r.status_code}): {payload.get('errors') or r.text[:700]}")
    parsed={}
    for complete in (payload.get("output") or {}).get("completeTrackResults") or []:
        requested=str(complete.get("trackingNumber") or "").strip()
        for result in complete.get("trackResults") or []:
            number=str((result.get("trackingNumberInfo") or {}).get("trackingNumber") or requested).strip()
            latest=result.get("latestStatusDetail") or {}
            desc=latest.get("description") or latest.get("statusByLocale") or ""
            parsed[number]={"tracking_number":number,"status":_normalize(latest.get("code"),desc,latest.get("derivedCode")),"raw_description":str(desc).strip()}
    for n in numbers: parsed.setdefault(n,{"tracking_number":n,"status":"UNKNOWN","raw_description":"No result returned by FedEx."})
    return parsed