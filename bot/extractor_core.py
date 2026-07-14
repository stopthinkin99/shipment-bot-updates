import cv2
import pytesseract
import re
import os
import numpy as np
import fitz
import pandas as pd
from openpyxl.styles import Font
from datetime import datetime

pytesseract.pytesseract.tesseract_cmd = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
os.environ["TESSDATA_PREFIX"] = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata"


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

    # INV Number
    m = re.search(r'\bI[\s.]*N[\s.]*V[\s:#]*([0-9]{4,})', text, re.IGNORECASE)
    if m:
        data["INV Number"] = m.group(1).strip()
    if not data["INV Number"]:
        for line in text.split('\n'):
            if re.search(r'\bI[\s.]*N[\s.]*V\b', line, re.IGNORECASE):
                nums = re.findall(r'\d{4,}', line)
                if nums:
                    data["INV Number"] = nums[0]
                    break

    # Reference
    m = re.search(r'\bR[\s.]*E[\s.]*F[\s:#]*([0-9]{4,})', text, re.IGNORECASE)
    if m:
        data["Reference"] = m.group(1).strip()
    if not data["Reference"]:
        for line in text.split('\n'):
            if re.search(r'\bR[\s.]*E[\s.]*F\b', line, re.IGNORECASE):
                nums = re.findall(r'\d{4,}', line)
                if nums:
                    data["Reference"] = nums[0]
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
    lines = text.split('\n')
    recipient = ""

    # Strategy 1: "TO <name>" — skip that line, take next non-empty
    for i, line in enumerate(lines):
        if re.match(r'^\s*TO\s+\S+', line, re.IGNORECASE):
            for j in range(i + 1, len(lines)):
                candidate = lines[j].strip()
                if candidate:
                    recipient = candidate
                    break
            break

    # Strategy 2: "SHIP TO:" block
    if not recipient:
        for i, line in enumerate(lines):
            if re.match(r'^\s*SHIP\s*TO\s*:', line, re.IGNORECASE):
                for j in range(i + 1, len(lines)):
                    candidate = lines[j].strip()
                    if candidate:
                        recipient = candidate
                        break
                break

    # Strategy 3: all-caps company name in first 15 lines, no digits, 2+ words
    if not recipient:
        for line in lines[:15]:
            line = line.strip()
            if (re.match(r'^[A-Z][A-Z\s&.,]{4,}$', line)
                    and not re.search(r'\d', line)
                    and len(line.split()) >= 2):
                recipient = line
                break

    data["Recipient Company"] = recipient
    return data


def process_single_image(img_array, filename, page):
    processed_img = deep_scan_preprocess(img_array)
    raw_text = extract_confident_lines(processed_img, r'--oem 3 --psm 11')
    clean_text = process_extracted_text(raw_text)
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