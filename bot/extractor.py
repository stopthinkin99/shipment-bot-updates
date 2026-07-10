"""
extractor.py  —  Local-only OCR, zero external APIs, zero cloud.
All label data stays on this machine.

Primary engine : RapidOCR  (pip install rapidocr-onnxruntime)
                 - Works on Python 3.14
                 - PaddleOCR-level accuracy via ONNX models
                 - ~80 MB install, no GPU needed
Fallback engine: Tesseract (already installed)

To install RapidOCR (one time, in PowerShell):
    py -m pip install rapidocr-onnxruntime
"""

import os
import numpy as np
from pdf2image import convert_from_path
from PIL import Image, ImageOps, ImageEnhance

# ------------------------------------------------------------------ #
#  PATHS  — update if Tesseract/Poppler move
# ------------------------------------------------------------------ #
TESSERACT_CMD = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
POPPLER_PATH  = r"C:\Users\aayan.boradia\Downloads\poppler-26.02.0\Library\bin"

OCR_DPI        = 400   # render DPI — higher = more detail per character
BORDER_PX      = 40    # white padding around each page (stops edge clipping)
CONTRAST_BOOST = 1.4   # mild contrast increase (makes thin strokes bolder)

# ------------------------------------------------------------------ #
#  RAPIDOCR — lazy load so Tesseract fallback still works if not installed
# ------------------------------------------------------------------ #
_rapid = None

def _get_rapid():
    global _rapid
    if _rapid is not None:
        return _rapid
    try:
        from rapidocr_onnxruntime import RapidOCR
        _rapid = RapidOCR()
        print("[OCR] RapidOCR engine loaded (primary)")
    except ImportError:
        _rapid = False
        print("[OCR] RapidOCR not found — run: py -m pip install rapidocr-onnxruntime")
        print("[OCR] Falling back to Tesseract")
    return _rapid


# ------------------------------------------------------------------ #
#  IMAGE PRE-PROCESSING
# ------------------------------------------------------------------ #
def _preprocess(img: Image.Image) -> Image.Image:
    img = img.convert("L")                                     # greyscale
    img = ImageOps.expand(img, border=BORDER_PX, fill=255)    # white border
    img = ImageEnhance.Contrast(img).enhance(CONTRAST_BOOST)  # contrast
    return img


# ------------------------------------------------------------------ #
#  PER-ENGINE EXTRACTION
# ------------------------------------------------------------------ #
def _ocr_rapid(img: Image.Image) -> str | None:
    rapid = _get_rapid()
    if not rapid:
        return None

    arr = np.array(img)
    result, _ = rapid(arr)   # returns (list_of_results, elapse)

    if not result:
        return ""

    # Each item: [box, (text, score)]
    lines = [item[1][0] for item in result if item[1][0].strip()]
    return "\n".join(lines)


def _ocr_tesseract(img: Image.Image) -> str:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
    return pytesseract.image_to_string(img, config=r"--oem 3 --psm 6")


# ------------------------------------------------------------------ #
#  PUBLIC
# ------------------------------------------------------------------ #
def extract_text(pdf_path: str) -> str:
    images = convert_from_path(
        pdf_path,
        dpi=OCR_DPI,
        poppler_path=POPPLER_PATH,
    )

    full_text = ""
    for img in images:
        img = _preprocess(img)

        text = _ocr_rapid(img)
        if text is None:                  # RapidOCR not installed
            text = _ocr_tesseract(img)

        full_text += text + "\n"

    print("\nOCR OUTPUT:")
    print(full_text)
    return full_text