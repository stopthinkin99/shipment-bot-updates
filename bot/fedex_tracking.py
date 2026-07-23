"""
fedex_tracking.py
-----------------
Official FedEx OAuth + Track API client.

TLS behavior:
- Uses Windows' native certificate store through `truststore`.
- This supports managed company PCs where a firewall, proxy, or antivirus
  installs a corporate root certificate trusted by Windows.
- SSL verification remains enabled. Never use verify=False.
"""

from __future__ import annotations

import threading
import time
import uuid
from dataclasses import dataclass
from typing import Dict, Iterable

# IMPORTANT: truststore must be injected before importing requests/urllib3.
try:
    import truststore

    truststore.inject_into_ssl()
    TRUSTSTORE_ACTIVE = True
    TRUSTSTORE_IMPORT_ERROR = ""
except Exception as exc:
    TRUSTSTORE_ACTIVE = False
    TRUSTSTORE_IMPORT_ERROR = f"{type(exc).__name__}: {exc}"

import requests
from requests.adapters import HTTPAdapter
from requests.exceptions import SSLError

from fedex_credentials import load_credentials


PRODUCTION_BASE_URL = "https://apis.fedex.com"
SANDBOX_BASE_URL = "https://apis-sandbox.fedex.com"
DEFAULT_ENVIRONMENT = "production"
REQUEST_TIMEOUT_SECONDS = 35
MAX_TRACKING_NUMBERS_PER_REQUEST = 30


class FedExTrackingError(RuntimeError):
    pass


@dataclass
class _Token:
    value: str
    expires_at: float


_cache: Dict[str, _Token] = {}
_lock = threading.Lock()
_tls_logged = False


def _base_url(environment: str) -> str:
    environment = str(environment or "").strip().lower()

    if environment in {"sandbox", "test"}:
        return SANDBOX_BASE_URL

    if environment in {"production", "prod", "live"}:
        return PRODUCTION_BASE_URL

    raise ValueError(
        "FedEx environment must be 'production' or 'sandbox'."
    )


def _log_tls_backend(log=print) -> None:
    global _tls_logged

    if _tls_logged:
        return

    _tls_logged = True

    if TRUSTSTORE_ACTIVE:
        log(
            "[FEDEX] HTTPS certificate verification: "
            "Windows certificate store enabled."
        )
    else:
        log(
            "[FEDEX] WARNING: Windows certificate-store support is not "
            "available because the truststore package could not be loaded."
        )
        log(
            "[FEDEX] truststore load error: "
            f"{TRUSTSTORE_IMPORT_ERROR or 'unknown error'}"
        )


def _session() -> requests.Session:
    """
    Return a normal Requests session.

    Because truststore is injected before Requests is imported, urllib3 uses
    the operating system's trusted roots while retaining full HTTPS
    certificate and hostname verification.
    """
    session = requests.Session()

    adapter = HTTPAdapter(
        pool_connections=4,
        pool_maxsize=8,
        max_retries=0,
    )
    session.mount("https://", adapter)

    return session


def _ssl_error_message(profile: str, exc: Exception) -> str:
    truststore_help = (
        "Install truststore in the same Python environment used to run or "
        "build Shipment Bot: py -m pip install truststore"
    )

    if TRUSTSTORE_ACTIVE:
        truststore_help = (
            "Windows certificate-store support is active, but Windows still "
            "does not trust the certificate presented for apis.fedex.com. "
            "Ask IT to verify that the company proxy/antivirus root "
            "certificate is installed under Trusted Root Certification "
            "Authorities for this Windows user or computer."
        )

    return (
        f"{profile} could not establish a verified HTTPS connection to "
        f"FedEx: {exc}. {truststore_help}. "
        "Do not disable SSL verification."
    )


def _token(
    profile: str,
    environment: str,
    log=print,
) -> str:
    profile = profile.strip().upper()
    environment = str(environment or "").strip().lower()
    cache_key = f"{environment}:{profile}"

    api_key, secret_key = load_credentials(profile)

    if not api_key or not secret_key:
        raise FedExTrackingError(
            f"FedEx credentials are not configured for {profile}."
        )

    _log_tls_backend(log)

    with _lock:
        cached = _cache.get(cache_key)

        if cached and cached.expires_at > time.time() + 120:
            return cached.value

        log(f"[FEDEX] {profile}: requesting OAuth token.")

        try:
            with _session() as session:
                response = session.post(
                    f"{_base_url(environment)}/oauth/token",
                    headers={
                        "Content-Type":
                            "application/x-www-form-urlencoded",
                    },
                    data={
                        "grant_type": "client_credentials",
                        "client_id": api_key,
                        "client_secret": secret_key,
                    },
                    timeout=REQUEST_TIMEOUT_SECONDS,
                )
        except SSLError as exc:
            raise FedExTrackingError(
                _ssl_error_message(profile, exc)
            ) from exc
        except requests.RequestException as exc:
            raise FedExTrackingError(
                f"{profile} OAuth connection failed: "
                f"{type(exc).__name__}: {exc}"
            ) from exc

        try:
            payload = response.json()
        except ValueError:
            payload = {}

        if response.status_code >= 400:
            detail = (
                payload.get("errors")
                or payload.get("error_description")
                or payload.get("error")
                or response.text[:500]
            )

            raise FedExTrackingError(
                f"{profile} OAuth failed "
                f"({response.status_code}): {detail}"
            )

        value = str(payload.get("access_token") or "").strip()

        if not value:
            raise FedExTrackingError(
                f"{profile} OAuth response did not contain access_token."
            )

        expires_in = int(payload.get("expires_in") or 3600)

        _cache[cache_key] = _Token(
            value=value,
            expires_at=time.time() + expires_in,
        )

        log(
            f"[FEDEX] {profile}: OAuth token received successfully."
        )

        return value


def _normalize(code, description, derived):
    combined = " ".join(
        str(value or "").upper()
        for value in (code, description, derived)
    )

    if "DELIVERED" in combined or str(code).upper() == "DL":
        return "DELIVERED"

    if (
        "OUT FOR DELIVERY" in combined
        or "ON FEDEX VEHICLE" in combined
    ):
        return "OUT FOR DELIVERY"

    if any(
        term in combined
        for term in (
            "EXCEPTION",
            "FAILED ATTEMPT",
            "UNABLE TO DELIVER",
            "DELAY",
        )
    ):
        return "DELIVERY EXCEPTION"

    if "PICKED UP" in combined or "PICKUP" in combined:
        return "PICKED UP"

    if any(
        term in combined
        for term in (
            "IN TRANSIT",
            "AT FEDEX FACILITY",
            "DEPARTED FEDEX",
            "ARRIVED AT FEDEX",
        )
    ):
        return "IN TRANSIT"

    if any(
        term in combined
        for term in (
            "LABEL CREATED",
            "SHIPMENT INFORMATION SENT",
            "INFORMATION SENT",
        )
    ):
        return "LABEL CREATED"

    if "RETURN" in combined:
        return "RETURNED"

    return str(
        description or derived or code or "UNKNOWN"
    ).strip().upper()[:80]


def _perform_tracking_request(
    environment: str,
    token: str,
    body: dict,
    profile: str,
):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "X-locale": "en_US",
        "X-customer-transaction-id": str(uuid.uuid4()),
    }

    try:
        with _session() as session:
            return session.post(
                f"{_base_url(environment)}/track/v1/trackingnumbers",
                headers=headers,
                json=body,
                timeout=REQUEST_TIMEOUT_SECONDS,
            )
    except SSLError as exc:
        raise FedExTrackingError(
            _ssl_error_message(profile, exc)
        ) from exc
    except requests.RequestException as exc:
        raise FedExTrackingError(
            f"{profile} Track API connection failed: "
            f"{type(exc).__name__}: {exc}"
        ) from exc


def track_numbers(
    profile: str,
    tracking_numbers: Iterable[str],
    *,
    environment="production",
    include_detailed_scans=False,
    log=print,
):
    profile = profile.strip().upper()
    environment = str(environment or "").strip().lower()

    numbers = []
    seen = set()

    for value in tracking_numbers:
        number = "".join(
            character
            for character in str(value or "")
            if character.isalnum()
        )

        if number and number not in seen:
            seen.add(number)
            numbers.append(number)

    if not numbers:
        return {}

    if len(numbers) > MAX_TRACKING_NUMBERS_PER_REQUEST:
        raise ValueError(
            "Maximum 30 tracking numbers per request."
        )

    token = _token(
        profile,
        environment,
        log,
    )

    body = {
        "includeDetailedScans": bool(include_detailed_scans),
        "trackingInfo": [
            {
                "trackingNumberInfo": {
                    "trackingNumber": number,
                }
            }
            for number in numbers
        ],
    }

    response = _perform_tracking_request(
        environment,
        token,
        body,
        profile,
    )

    if response.status_code == 401:
        with _lock:
            _cache.pop(
                f"{environment}:{profile}",
                None,
            )

        token = _token(
            profile,
            environment,
            log,
        )

        response = _perform_tracking_request(
            environment,
            token,
            body,
            profile,
        )

    try:
        payload = response.json()
    except ValueError:
        payload = {}

    if response.status_code >= 400:
        detail = (
            payload.get("errors")
            or response.text[:700]
        )

        raise FedExTrackingError(
            f"{profile} Track API failed "
            f"({response.status_code}): {detail}"
        )

    parsed = {}

    complete_results = (
        (payload.get("output") or {})
        .get("completeTrackResults")
        or []
    )

    for complete in complete_results:
        requested = str(
            complete.get("trackingNumber") or ""
        ).strip()

        for result in complete.get("trackResults") or []:
            number = str(
                (result.get("trackingNumberInfo") or {})
                .get("trackingNumber")
                or requested
            ).strip()

            latest = result.get("latestStatusDetail") or {}
            description = (
                latest.get("description")
                or latest.get("statusByLocale")
                or ""
            )

            parsed[number] = {
                "tracking_number": number,
                "status": _normalize(
                    latest.get("code"),
                    description,
                    latest.get("derivedCode"),
                ),
                "raw_description": str(description).strip(),
            }

    for number in numbers:
        parsed.setdefault(
            number,
            {
                "tracking_number": number,
                "status": "UNKNOWN",
                "raw_description":
                    "No result returned by FedEx.",
            },
        )

    return parsed