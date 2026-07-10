"""
extractor.py  —  Local-only OCR. Zero external APIs. Zero cloud.
Uses Tesseract with image preprocessing to maximize accuracy.

Preprocessing applied before OCR:
  1. Greyscale conversion
  2. 40px white border on all sides  (stops edge-character clipping like EBONY -> BONY)
  3. 2x upscale                      (more pixels per character = better recognition)
  4. Contrast boost                  (makes thin strokes bolder)
"""

import os
import sys
from pathlib import Path
from pdf2image import convert_from_path
from PIL import Image, ImageOps, ImageEnhance
import pytesseract

# ------------------------------------------------------------------ #
#  PATHS — resolved relative to bundle so works on any PC
# ------------------------------------------------------------------ #
_BASE = Path(sys.executable).parent if getattr(sys, "frozen", False) \
        else Path(__file__).parent

TESSERACT_CMD = str(_BASE / "Tesseract-OCR" / "tesseract.exe")
POPPLER_PATH  = str(_BASE / "poppler" / "bin")

# Fallback to original paths if running directly on dev machine
if not os.path.exists(TESSERACT_CMD):
    TESSERACT_CMD = r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
if not os.path.exists(POPPLER_PATH):
    POPPLER_PATH = r"C:\Users\aayan.boradia\Downloads\poppler-26.02.0\Library\bin"

pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD

OCR_DPI        = 300   # render DPI — 300 is the sweet spot for Tesseract
UPSCALE        = 2     # multiply resolution before OCR
BORDER_PX      = 40    # white padding to stop edge-character clipping
CONTRAST_BOOST = 1.5   # mild contrast increase


# ------------------------------------------------------------------ #
#  PRE-PROCESSING
# ------------------------------------------------------------------ #
def _preprocess(img: Image.Image) -> Image.Image:
    # 1. Greyscale
    img = img.convert("L")

    # 2. White border — prevents characters at label edges being clipped
    img = ImageOps.expand(img, border=BORDER_PX, fill=255)

    # 3. Upscale — gives Tesseract more pixels per character
    w, h = img.size
    img = img.resize((w * UPSCALE, h * UPSCALE), Image.LANCZOS)

    # 4. Contrast boost — makes thin FedEx/UPS label text bolder
    img = ImageEnhance.Contrast(img).enhance(CONTRAST_BOOST)

    return img


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
        # OEM 3 = best LSTM engine, PSM 6 = uniform block of text
        text = pytesseract.image_to_string(img, config=r"--oem 3 --psm 6")
        full_text += text + "\n"

    print("\nOCR OUTPUT:")
    print(full_text)
    return full_text