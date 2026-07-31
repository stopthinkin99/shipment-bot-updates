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

    full_text_parts = []
    for source_image in images:
        try:
            for angle in (0, 90, 270):
                rotated = (
                    source_image
                    if angle == 0
                    else source_image.rotate(angle, expand=True, fillcolor="white")
                )
                image = _preprocess(rotated)
                for psm in (6, 11):
                    full_text_parts.append(
                        pytesseract.image_to_string(
                            image,
                            config=f"--oem 3 --psm {psm}",
                        )
                    )
        finally:
            source_image.close()

    full_text = "\n".join(full_text_parts)
    print("\nOCR OUTPUT:")
    print(full_text)
    return full_text
