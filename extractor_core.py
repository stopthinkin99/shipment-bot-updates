import cv2
import pytesseract
import re
import os
import numpy as np
import fitz
import pandas as pd
from openpyxl.styles import Font
from datetime import datetime



# Runtime-safe OCR paths
import sys
from pathlib import Path

def _find_tesseract_paths():
    """
    Find bundled or installed Tesseract without using a user-specific path.
    """
    if getattr(sys, "frozen", False):
        app_dir = Path(sys.executable).resolve().parent
    else:
        app_dir = Path(__file__).resolve().parent

    meipass = Path(
        getattr(sys, "_MEIPASS", app_dir)
    )

    candidates = [
        app_dir / "Tesseract-OCR",
        app_dir / "bundle" / "Tesseract-OCR",
        meipass / "Tesseract-OCR",
        Path("C:/Program Files/Tesseract-OCR"),
        Path("C:/Program Files (x86)/Tesseract-OCR"),
    ]

    for folder in candidates:
        exe = folder / "tesseract.exe"
        tessdata = folder / "tessdata"

        if exe.is_file() and tessdata.is_dir():
            return exe, tessdata

    checked = "\n".join(
        f" - {folder}" for folder in candidates
    )

    raise FileNotFoundError(
        "Tesseract OCR could not be found.\n"
        "Checked:\n"
        f"{checked}"
    )


_TESSERACT, _TESSDATA = _find_tesseract_paths()

pytesseract.pytesseract.tesseract_cmd = str(
    _TESSERACT
)

os.environ["TESSDATA_PREFIX"] = str(
    _TESSDATA
)

def deep_scan_preprocess(img):
    scale = 3.5
    max_dim = 3000
    if img.shape[1] * scale > max_dim or img.shape[0] * scale > max_dim:
        scale = min(max_dim / img.shape[1], max_dim / img.shape[0])
    scale = max(1.0, scale)
    width = int(img.shape[1] * scale)
    height = int(img.shape[0] * scale)
    upscaled = cv2.resize(img, (width, height), interpolation=cv2.INTER_CUBIC)
    blurred = cv2.GaussianBlur(upscaled, (3, 3), 0.8)
    gray = cv2.cvtColor(blurred, cv2.COLOR_BGR2GRAY) if len(blurred.shape) == 3 else blurred
    high_contrast = cv2.convertScaleAbs(gray, alpha=1.8, beta=10)
    _, binarized = cv2.threshold(high_contrast, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return binarized


def extract_confident_lines(img, config):
    data = pytesseract.image_to_data(img, config=config, output_type=pytesseract.Output.DICT)
    lines = {}
    for i in range(len(data['text'])):
        text = data['text'][i].strip()
        conf = float(data['conf'][i])
        if not text:
            continue
        line_key = f"{data['block_num'][i]}_{data['par_num'][i]}_{data['line_num'][i]}"
        if line_key not in lines:
            lines[line_key] = {'text': [], 'conf': []}
        lines[line_key]['text'].append(text)
        if conf != -1:
            lines[line_key]['conf'].append(conf)
    confident_lines = []
    for key, val in lines.items():
        if not val['text']:
            continue
        avg_conf = sum(val['conf']) / len(val['conf']) if val['conf'] else 0
        if avg_conf > 45:
            confident_lines.append(" ".join(val['text']))
    return "\n".join(confident_lines)


def process_extracted_text(raw_text):
    lines = raw_text.split('\n')
    valid_lines = []
    for line in lines:
        line = line.strip()
        if not line:
            continue
        chars = len(line)
        symbols = len(re.findall(r'[^a-zA-Z0-9\s.,:-]', line))
        if symbols > chars * 0.35:
            continue
        lowers = len(re.findall(r'[a-z]', line))
        uppers = len(re.findall(r'[A-Z]', line))
        if lowers > 0 and lowers > uppers * 1.5:
            continue
        if len(line) <= 2 and not re.match(r'^\d+$', line):
            continue

        def force_numbers(prefix_regex, current_line):
            pattern = re.compile(f'({prefix_regex})([a-zA-Z0-9]+)', re.IGNORECASE)
            def repl(m):
                nums = re.sub(r'[Oo]', '0', m.group(2))
                nums = re.sub(r'[Il]', '1', nums)
                nums = re.sub(r'[Ss]', '5', nums)
                return m.group(1) + nums
            return pattern.sub(repl, current_line)

        line = force_numbers(r'\bR[\s.]*E[\s.]*F[\s:]*', line)
        line = force_numbers(r'\bI[\s.]*N[\s.]*V[\s:]*', line)

        def fix_cad(m):
            parts = m.group(2).split('/')
            if parts:
                parts[0] = re.sub(r'[Oo]', '0', parts[0])
                parts[0] = re.sub(r'[Il]', '1', parts[0])
            return m.group(1) + '/'.join(parts)

        line = re.sub(r'(\bC[\s.]*A[\s.]*D[\s:]*)(.*)', fix_cad, line, flags=re.IGNORECASE)
        line = re.sub(r'\b([0-9]{1,5})[Oo]([0-9]{1,5})\b', r'\g<1>0\g<2>', line)
        line = re.sub(r'\b[Oo]([0-9]{3,})\b', r'0\g<1>', line)
        line = re.sub(r'\b([0-9]{3,})[Oo]\b', r'\g<1>0', line)
        valid_lines.append(line)
    return '\n'.join(valid_lines)



_COMPANY_SUFFIX_PATTERN = re.compile(
    r"\b(?:AG|INC|INCORPORATED|LLC|LTD|LIMITED|CORP|CORPORATION|"
    r"COMPANY|CO|LP|PLC|GMBH|S\.A\.|SA|BV|NV)\b",
    re.IGNORECASE,
)

_ADDRESS_WORD_PATTERN = re.compile(
    r"\b(?:ST|STREET|AVE|AVENUE|ROAD|RD|BLVD|BOULEVARD|"
    r"TURNPIKE|DRIVE|DR|HIGHWAY|HWY|LANE|LN|COURT|CT)\b",
    re.IGNORECASE,
)

_STOP_BLOCK_PATTERN = re.compile(
    r"\b(?:UPS|FEDEX|TRACKING|BILLING|SIGNATURE|REF|REFERENCE|"
    r"WEIGHT|SERVICE|PACKAGE)\b",
    re.IGNORECASE,
)


def _extract_ship_to_company(text):
    """
    Extract the actual company from the SHIP TO block.

    Example:
        SHIP TO:
        BRAND THIES
        (401) 463-2509
        SWAROVSKI AG.
        7830 NATIONAL TURNPIKE

    Returns:
        SWAROVSKI AG
    """
    lines = [
        line.strip()
        for line in str(text or "").splitlines()
        if line.strip()
    ]

    ship_to_index = None

    for index, line in enumerate(lines):
        if re.search(r"\bSHIP\s*TO\s*:?", line, re.IGNORECASE):
            ship_to_index = index
            break

    if ship_to_index is None:
        return ""

    candidates = []

    for line in lines[ship_to_index + 1:ship_to_index + 12]:
        cleaned = line.strip()
        upper = cleaned.upper().strip(" .")

        if _STOP_BLOCK_PATTERN.search(upper):
            break

        if re.fullmatch(r"[\d\s()+\-]+", cleaned):
            continue

        if upper.startswith(("C/O ", "C/O:", "ATTN ", "ATTN:")):
            continue

        if _ADDRESS_WORD_PATTERN.search(upper):
            continue

        if re.search(r"\b[A-Z]{2}\s+\d{5}(?:-\d{4})?\b", upper):
            continue

        letters = len(re.findall(r"[A-Z]", upper))
        digits = len(re.findall(r"\d", upper))

        if letters < 3 or digits > letters:
            continue

        candidates.append(cleaned.rstrip(" ."))

    for candidate in candidates:
        if _COMPANY_SUFFIX_PATTERN.search(candidate):
            return candidate

    company_keywords = re.compile(
        r"\b(?:DIAMOND|JEWEL|JEWELRY|DESIGN|CREATION|"
        r"INTERNATIONAL|INDUSTRIES|GROUP|SUPPLY)\b",
        re.IGNORECASE,
    )

    for candidate in candidates:
        if company_keywords.search(candidate):
            return candidate

    return candidates[0] if candidates else ""



def _normalize_document_number(value):
    value = str(value or "").strip()
    value = re.sub(r"\s+", " ", value)
    value = re.sub(r"\s*/\s*", "/", value)
    value = re.sub(r"\s*-\s*", "-", value)
    value = re.sub(r"\s*,\s*", ", ", value)
    return value.strip(" .,:;-")


def _extract_document_number(text, labels):
    label_pattern = "|".join(labels)

    for raw_line in str(text or "").splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        match = re.search(
            rf"\b(?:{label_pattern})\b"
            r"\s*(?:NUMBER|NO\.?|#)?\s*[:\-]?\s*"
            r"([A-Z0-9][A-Z0-9/,\- ]{2,60})",
            line,
            re.IGNORECASE,
        )
        if not match:
            continue

        candidate = match.group(1)
        candidate = re.split(
            r"\s{2,}|\b(?:DATE|SHIP TO|TRACKING|WEIGHT|PHONE|REF(?:ERENCE)?)\b",
            candidate,
            maxsplit=1,
            flags=re.IGNORECASE,
        )[0]
        candidate = _normalize_document_number(candidate)

        if re.search(r"\d", candidate):
            return candidate

    return ""

def parse_core_fields(text, filename, page):
    data = {
        "Source File": filename,
        "Page": page,
        "Tracking Number": "",
        "PO Number": "",
        "Recipient Company": "",
        "INV Number": "",
        "Reference": "",
        "CAD": "",
        "Weight": "",
        "Ship Date": "",
        "Full Extracted Text": text,
        "Application Run Date and time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # Preserve PO / Invoice / Memo exactly as printed.
    data["PO Number"] = _extract_document_number(
        text,
        (r"P[\s.]*[O0]", r"PURCHASE\s+ORDER"),
    )

    data["INV Number"] = _extract_document_number(
        text,
        (r"I[\s.]*N[\s.]*V(?:OICE)?", r"INVOICE", r"MEMO"),
    )

    # Indexed UPS reference, e.g. REF 1:47122997
    m = re.search(
        r'\bR[\s.]*E[\s.]*F(?:ERENCE)?\s*#?\s*1\s*[:\-]\s*'
        r'([A-Z0-9\-]{4,})',
        text,
        re.IGNORECASE,
    )
    if m:
        data["Reference"] = m.group(1).strip()

    # Generic reference fallback
    if not data["Reference"]:
        m = re.search(
            r'\bR[\s.]*E[\s.]*F(?:ERENCE)?[\s:#\-]+'
            r'([A-Z0-9\-]{4,})',
            text,
            re.IGNORECASE,
        )
        if m and m.group(1).strip() != "1":
            data["Reference"] = m.group(1).strip()

    if not data["Reference"]:
        for line in text.splitlines():
            if re.search(r'\bREF(?:ERENCE)?\b', line, re.IGNORECASE):
                indexed = re.search(
                    r'\bREF(?:ERENCE)?\s*1\s*[:\-]\s*'
                    r'([A-Z0-9\-]{4,})',
                    line,
                    re.IGNORECASE,
                )
                if indexed:
                    data["Reference"] = indexed.group(1).strip()
                    break

                values = re.findall(r'[A-Z0-9\-]{4,}', line, re.IGNORECASE)
                if values:
                    data["Reference"] = values[-1]
                    break

    # CAD
    m = re.search(r'\bC[\s.]*A[\s.]*D[\s:]*([\w\d/]+)', text, re.IGNORECASE)
    if m:
        data["CAD"] = m.group(1).strip()

    # Weight
    m = re.search(r'\bACTWGT[\s:]*(.*?LB)', text, re.IGNORECASE)
    if m:
        data["Weight"] = m.group(1).strip()

    # Ship Date
    m = re.search(r'\bDATE[\s:]*([\w\d]+)', text, re.IGNORECASE)
    if m:
        data["Ship Date"] = m.group(1).strip()

    # Recipient Company
    recipient = _extract_ship_to_company(text)

    # Fallback for inline layouts such as: TO REEDS GWLERS
    if not recipient:
        lines = text.splitlines()

        for i, line in enumerate(lines):
            inline = re.match(
                r'^\s*TO\s*[:\-]?\s+(.+?)\s*$',
                line,
                re.IGNORECASE,
            )

            if not inline:
                continue

            candidate = inline.group(1).strip(" .,:;-")

            if (
                len(re.findall(r'[A-Za-z]', candidate)) >= 3
                and not re.search(
                    r'\b(?:ST|STREET|AVE|AVENUE|RD|ROAD|BLVD|'
                    r'DR|DRIVE|LANE|LN)\b',
                    candidate,
                    re.IGNORECASE,
                )
            ):
                recipient = candidate
                break

            for following in lines[i + 1:]:
                following = following.strip()
                if following:
                    recipient = following
                    break

            break

    # Final fallback: formal all-caps company name.
    if not recipient:
        for line in text.splitlines()[:20]:
            candidate = line.strip()
            if (
                re.match(r'^[A-Z][A-Z\s&.,\-]{4,}$', candidate)
                and not re.search(r'\d', candidate)
                and len(candidate.split()) >= 2
                and _COMPANY_SUFFIX_PATTERN.search(candidate)
            ):
                recipient = candidate.rstrip(" .")
                break

    data["Recipient Company"] = recipient
    return data


def _rotate_image(img, angle):
    if angle == 90:
        return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    if angle == 270:
        return cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    return img


def process_single_image(img_array, filename, page):
    # Some UPS print pages contain a landscape label rotated inside a
    # portrait instruction page. OCR all useful orientations and merge.
    all_text = []
    for angle in (0, 90, 270):
        rotated = _rotate_image(img_array, angle)
        processed = deep_scan_preprocess(rotated)
        all_text.append(
            pytesseract.image_to_string(
                processed,
                config="--oem 3 --psm 6",
            )
        )
        all_text.append(
            extract_confident_lines(
                processed,
                "--oem 3 --psm 11",
            )
        )

    clean_text = process_extracted_text("\n".join(all_text))
    return parse_core_fields(clean_text, filename, page)

def extract_data_from_file(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Cannot find file at: {file_path}")
    ext = file_path.lower().split('.')[-1]
    filename = os.path.basename(file_path)
    extracted_records = []
    if ext == 'pdf':
        doc = fitz.open(file_path)
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            pix = page.get_pixmap(dpi=300)
            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.h, pix.w, pix.n)
            if pix.n == 4:
                img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGBA2BGR)
            elif pix.n == 3:
                img_cv = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
            else:
                img_cv = img_array
            extracted_records.append(process_single_image(img_cv, filename, page_num + 1))
        doc.close()
    elif ext in ['jpg', 'jpeg', 'png', 'bmp', 'tiff']:
        img_cv = cv2.imread(file_path)
        if img_cv is None:
            raise ValueError("Failed to load image file.")
        extracted_records.append(process_single_image(img_cv, filename, 1))
    else:
        raise ValueError(f"Unsupported file format: {ext}")
    return extracted_records


def save_to_excel(records, output_filename):
    df = pd.DataFrame(records)
    col_order = [
        "Source File", "Page", "Tracking Number", "PO Number",
        "Recipient Company", "INV Number", "Reference", "CAD",
        "Weight", "Ship Date", "Full Extracted Text",
        "Application Run Date and time"
    ]
    df = df[[c for c in col_order if c in df.columns]]
    if os.path.exists(output_filename):
        try:
            existing_df = pd.read_excel(output_filename)
            df = pd.concat([existing_df, df], ignore_index=True)
        except Exception:
            pass
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Extracted Data')
        worksheet = writer.sheets['Extracted Data']
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            worksheet.column_dimensions[col_letter].width = min(max_length + 3, 60)