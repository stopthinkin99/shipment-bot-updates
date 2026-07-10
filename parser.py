import re


# ------------------------------------------------------------------ #
#  LOOKUP TABLES
# ------------------------------------------------------------------ #

# ORIGIN ID -> shipping company.
#   OGSA        = Brinx Fedex   -> carrier "BX FX ..."
#   NYSA / NYCA = MalcaAmit     -> carrier "M / E"
ORIGIN_ID_MAP = {
    "OGSA": "BRINX_FEDEX",
    "NYSA": "MALCAAMIT",
    "NYCA": "MALCAAMIT",
}

# FedEx service text -> short code used after the "BX FX" prefix
FEDEX_SERVICE_SHORT = {
    "PRIORITY OVERNIGHT": "P/O",
    "STANDARD OVERNIGHT": "S/O",
}

# Routing: leading digits of the INV/PO/REF number -> Excel sheet.
# Order = most specific first (2030 before 20-anything, etc.)
INVOICE_PREFIX_SHEET = [
    ("2030", "EMBY"),
    ("82",   "FENIX"),
    ("47",   "FENIX"),
    ("10",   "UNI"),
]

# Where routing / invoice numbers can appear on a label. Checked in
# THIS priority order (PO first, Reference last).
NUMBER_MARKERS = [
    r'PO[:\s#]*([0-9]{3,})',
    r'INV[:\s#]*([0-9]{3,})',
    r'REF[:\s#]*([0-9]{3,})',
    r'Reference\s*#?\s*\d*[:\s]*([0-9]{3,})',
]


# ------------------------------------------------------------------ #
#  BASIC FIELD HELPERS
# ------------------------------------------------------------------ #

def _find_origin_id(text):
    """Grab the ORIGIN ID code (OGSA, NYSA, NYCA...) from the header."""
    m = re.search(r'ORIGIN\s*ID[:\s]*([A-Z0-9]+)', text, re.IGNORECASE)
    return m.group(1).strip().upper() if m else ""


def _find_date(text):
    """
    Ship date in compact form, e.g. 09JUL26 / 06JUL26.
    We match the DATE VALUE directly instead of relying on the words
    'SHIP DATE', because OCR often mangles them (e.g. 'Son DATE:').
    """
    # Prefer a value sitting next to the word DATE
    m = re.search(r'DATE[:;\s]*([0-9]{1,2}[A-Z]{3}[0-9]{2,4})',
                  text, re.IGNORECASE)
    if m:
        return m.group(1).upper()

    # Fallback: any compact DDMONYY token anywhere on the label
    m = re.search(r'\b([0-9]{1,2}[A-Z]{3}[0-9]{2,4})\b', text, re.IGNORECASE)
    return m.group(1).upper() if m else ""


def _sheet_from_invoice(number):
    """Map a number's leading digits to the correct sheet name."""
    for prefix, sheet in INVOICE_PREFIX_SHEET:
        if number.startswith(prefix):
            return sheet
    return ""


def _gather_numbers(text):
    """Collect all candidate numbers from PO / INV / REF / Reference."""
    nums = []
    for pat in NUMBER_MARKERS:
        for m in re.finditer(pat, text, re.IGNORECASE):
            n = m.group(1)
            if n not in nums:
                nums.append(n)
    return nums


def _route(text):
    """
    Return (invoice_number, sheet_name).
    Picks the first candidate number whose prefix maps to a known sheet;
    that same number is what we write into the invoice column.
    """
    candidates = _gather_numbers(text)

    for n in candidates:
        sheet = _sheet_from_invoice(n)
        if sheet:
            return n, sheet

    # nothing mapped -> keep first number for the cell, but no sheet
    return (candidates[0] if candidates else ""), ""


def _is_zales(text):
    return "ZALES" in text.upper()


# ------------------------------------------------------------------ #
#  CARRIER
# ------------------------------------------------------------------ #

def _extract_fedex_service(text):
    """Return the FedEx service line, e.g. 'PRIORITY OVERNIGHT'."""
    up = text.upper()
    for service in FEDEX_SERVICE_SHORT:
        if service in up:
            return service
    m = re.search(r'([A-Z][A-Z ]*OVERNIGHT)', up)
    return m.group(1).strip() if m else ""


def _extract_ups_service(text):
    """Pull a UPS service like 'UPS GROUND', 'UPS 2ND DAY AIR'."""
    m = re.search(r'(UPS[ A-Z0-9]*?(?:GROUND|DAY|AIR|NEXT|SAVER|EXPEDITED))',
                  text.upper())
    if m:
        return re.sub(r'\s+', ' ', m.group(1)).strip()
    return "UPS" if "UPS" in text.upper() else ""


def _build_carrier(shipper, text):
    """Carrier column value based on the decoded shipper."""
    if shipper == "MALCAAMIT":
        return "M / E"

    if shipper == "BRINX_FEDEX":
        service = _extract_fedex_service(text)
        short = FEDEX_SERVICE_SHORT.get(service)
        if short:
            return f"BX FX {short}"          # BX FX P/O  /  BX FX S/O
        if service:
            return f"BX FX {service.title()}"  # BX FX <other service>
        return "BX FX"

    # unknown origin -> best-effort raw service
    return _extract_fedex_service(text)


# ------------------------------------------------------------------ #
#  TRACKING
# ------------------------------------------------------------------ #

def _extract_tracking(text, zales=False):
    """
    FedEx: '8741 6222 5747' (12 digits, spaced).
    UPS (Zales): '1Z...' 18 chars, often OCR'd with spaces:
        '1Z Y99 G35 03 1047 0932' -> 1ZY99G350310470932
    """
    if zales:
        m = re.search(r'1Z[\sA-Z0-9]{14,}', text.upper())
        if m:
            cleaned = re.sub(r'\s+', '', m.group(0))
            return cleaned[:18]   # a 1Z tracking number is 18 chars

    m = re.search(r'\b(\d{4}\s\d{4}\s\d{4})\b', text)
    return m.group(1) if m else ""


# ------------------------------------------------------------------ #
#  STANDARD (FedEx) SHIP-TO
# ------------------------------------------------------------------ #

def _standard_ship_to(text):
    """
    FedEx 'TO' blocks carry two name lines. RULE: take the SECOND name
    line (the real recipient). Fall back to the only line if just one.
        TO  INDIAN DC          <- line 1
            INDIAN DND         <- line 2  == what we want
            29 EAST MADISON... <- address (starts with a digit)
    """
    m = re.search(r'^[ \t]*TO\b[ \t]*(.*)$', text, re.IGNORECASE | re.MULTILINE)
    if not m:
        return ""

    block = text[m.start(1):]

    names = []
    for line in block.splitlines():
        c = line.strip()
        if not c:
            if names:
                break
            continue
        if re.match(r'^[\(\d]', c):        # street address / phone -> stop
            break
        if re.match(r"^[A-Z][A-Z0-9 .,&'\-]+$", c):
            names.append(c)
        elif names:
            break

    if len(names) >= 2:
        return names[1].strip()
    if names:
        return names[0].strip()
    return ""


# ------------------------------------------------------------------ #
#  ZALES SHIP-TO
# ------------------------------------------------------------------ #

def _zales_ship_to(text):
    """
    Zales 'SHIP TO' is either a store pickup id (ZJC...) + a person, or
    just a person. If a ZJC id exists -> 'ZJC..., PERSON'. Else -> person.
    """
    zjc = re.search(r'\b(ZJC[0-9]+)\b', text.upper())

    name = ""
    m = re.search(r'SHIP\s*TO[:\s]*(.*)', text, re.IGNORECASE | re.DOTALL)
    if m:
        for line in m.group(1).splitlines():
            c = line.strip()
            if not c or c.upper().startswith("ZJC"):
                continue
            if re.match(r'^[\(\d]', c):     # phone / all-digit line -> skip
                continue
            if re.match(r"^[A-Za-z][A-Za-z .,&'\-]+$", c):
                name = c.upper()
                break

    if zjc and name:
        return f"{zjc.group(1)}, {name}"
    if zjc:
        return zjc.group(1)
    return name


# ------------------------------------------------------------------ #
#  MAIN PARSE
# ------------------------------------------------------------------ #

def parse_label(text):

    data = {
        "date": "",
        "ship_to": "",
        "invoice": "",
        "carrier": "",
        "tracking_number": "",
        "sheet": "",
        "remark": "",
        "shipper": "",
    }

    zales = _is_zales(text)

    # ---- DATE (shared) ---- #
    data["date"] = _find_date(text)

    # ---- ROUTING NUMBER -> invoice + sheet (shared) ---- #
    invoice, sheet = _route(text)
    data["invoice"] = invoice
    data["sheet"] = sheet

    # ------------------------------------------------------------------ #
    #  ZALES BRANCH  (UPS delivery, always FENIX)
    # ------------------------------------------------------------------ #
    if zales:
        data["shipper"] = "ZALES"
        data["ship_to"] = _zales_ship_to(text)
        data["carrier"] = _extract_ups_service(text)
        data["tracking_number"] = _extract_tracking(text, zales=True)
        data["sheet"] = sheet or "FENIX"       # Zales is Fenix (82...)
        data["remark"] = ("ZALES STORE ACC"
                          if "ZJC" in text.upper()
                          else "ZALES CUST ACC")
        return data

    # ------------------------------------------------------------------ #
    #  STANDARD BRANCH  (Brinx Fedex / MalcaAmit)
    # ------------------------------------------------------------------ #
    origin_id = _find_origin_id(text)
    data["shipper"] = ORIGIN_ID_MAP.get(origin_id, "")

    data["ship_to"] = _standard_ship_to(text)
    data["carrier"] = _build_carrier(data["shipper"], text)
    data["tracking_number"] = _extract_tracking(text)

    return data