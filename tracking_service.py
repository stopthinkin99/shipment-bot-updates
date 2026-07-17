"""EasyPost-backed tracking for UPS, FedEx, and USPS."""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

import requests

API_BASE = "https://api.easypost.com/v2"
TIMEOUT = 30


def _app_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    folder = Path(base) / "UniCreation" / "ShipmentBot" if base else Path.home() / ".uni_creation_shipment_bot"
    folder.mkdir(parents=True, exist_ok=True)
    return folder


CONFIG_FILE = _app_dir() / "tracking_config.json"
CACHE_FILE = _app_dir() / "tracker_cache.json"


def save_api_key(api_key: str) -> Path:
    key = str(api_key or "").strip()
    if not key:
        raise ValueError("EasyPost API key cannot be blank.")
    CONFIG_FILE.write_text(json.dumps({"easypost_api_key": key}, indent=2), encoding="utf-8")
    return CONFIG_FILE


def load_api_key() -> str:
    key = os.environ.get("EASYPOST_API_KEY", "").strip()
    if key:
        return key
    try:
        return str(json.loads(CONFIG_FILE.read_text(encoding="utf-8")).get("easypost_api_key", "")).strip()
    except Exception:
        return ""


def _load_cache() -> dict[str, str]:
    try:
        data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _save_cache(cache: dict[str, str]) -> None:
    CACHE_FILE.write_text(json.dumps(cache, indent=2, sort_keys=True), encoding="utf-8")


def normalize_tracking_number(value: Any) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def normalize_carrier(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "").strip().upper())
    if "USPS" in text or "UNITED STATES POSTAL" in text:
        return "USPS"
    if "UPS" in text:
        return "UPS"
    if "FEDEX" in text or "FED EX" in text or "BX FX" in text or re.search(r"\bFX\b", text):
        return "FedEx"
    return ""


STATUS_MAP = {
    "pre_transit": "LABEL CREATED",
    "in_transit": "IN TRANSIT",
    "out_for_delivery": "OUT FOR DELIVERY",
    "delivered": "DELIVERED",
    "available_for_pickup": "AVAILABLE FOR PICKUP",
    "return_to_sender": "RETURN TO SENDER",
    "failure": "DELIVERY EXCEPTION",
    "cancelled": "CANCELLED",
    "error": "TRACKING ERROR",
    "unknown": "UNKNOWN",
}


def _latest_message(tracker: dict[str, Any]) -> str:
    details = tracker.get("tracking_details") or []
    if not isinstance(details, list) or not details:
        return ""
    latest = details[-1]
    return str(latest.get("message") or latest.get("status_detail") or "").strip() if isinstance(latest, dict) else ""


def _normalize_status(tracker: dict[str, Any]) -> str:
    message = _latest_message(tracker).upper()
    if any(term in message for term in (
        "DELIVERY ATTEMPT", "ATTEMPTED DELIVERY", "NOTICE LEFT", "NO ACCESS",
        "MISSED DELIVERY", "RECIPIENT NOT AVAILABLE", "CUSTOMER NOT AVAILABLE"
    )):
        return "DELIVERY ATTEMPTED"
    raw = str(tracker.get("status") or "unknown").strip().lower()
    return STATUS_MAP.get(raw, raw.replace("_", " ").upper())


class TrackingServiceError(RuntimeError):
    pass


def _request(method: str, endpoint: str, api_key: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
    try:
        response = requests.request(
            method,
            f"{API_BASE}{endpoint}",
            auth=(api_key, ""),
            json=payload,
            timeout=TIMEOUT,
        )
    except requests.RequestException as exc:
        raise TrackingServiceError(f"Tracking request failed: {exc}") from exc

    if response.status_code not in {200, 201}:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise TrackingServiceError(f"EasyPost HTTP {response.status_code}: {detail}")

    try:
        result = response.json()
    except ValueError as exc:
        raise TrackingServiceError("EasyPost returned invalid JSON.") from exc
    if not isinstance(result, dict):
        raise TrackingServiceError("EasyPost returned an unexpected response.")
    return result


def get_tracking_status(carrier: str, tracking_number: str, api_key: str | None = None) -> dict[str, str]:
    key = str(api_key or load_api_key()).strip()
    if not key:
        raise TrackingServiceError("EasyPost API key is not configured. Run configure_tracking.py.")

    carrier_name = normalize_carrier(carrier)
    tracking = normalize_tracking_number(tracking_number)
    if not carrier_name:
        raise TrackingServiceError(f"Unsupported carrier: {carrier!r}")
    if not tracking:
        raise TrackingServiceError("Tracking number is blank.")

    cache = _load_cache()
    cache_key = f"{carrier_name.upper()}::{tracking}"
    tracker_id = cache.get(cache_key, "")
    tracker = None

    if tracker_id:
        try:
            tracker = _request("GET", f"/trackers/{tracker_id}", key)
        except TrackingServiceError:
            cache.pop(cache_key, None)
            _save_cache(cache)

    if tracker is None:
        tracker = _request(
            "POST",
            "/trackers",
            key,
            {"tracker": {"tracking_code": tracking, "carrier": carrier_name}},
        )
        new_id = str(tracker.get("id") or "").strip()
        if new_id:
            cache[cache_key] = new_id
            _save_cache(cache)

    return {
        "carrier": carrier_name,
        "tracking_number": tracking,
        "status": _normalize_status(tracker),
        "raw_status": str(tracker.get("status") or "unknown"),
        "message": _latest_message(tracker),
    }