from __future__ import annotations

import os
import sys
from pathlib import Path


def _unique_paths(paths: list[Path]) -> list[Path]:
    results: list[Path] = []
    seen: set[str] = set()

    for path in paths:
        try:
            normalized = str(path.resolve()).lower()
        except Exception:
            normalized = str(path).lower()

        if normalized not in seen:
            seen.add(normalized)
            results.append(path)

    return results


def application_directories() -> list[Path]:
    """
    Return possible application locations for development, PyInstaller,
    installed builds, and the generated bundle folder.
    """
    source_dir = Path(__file__).resolve().parent
    current_dir = Path.cwd()

    candidates = [
        source_dir,
        source_dir.parent,
        current_dir,
        current_dir / "bundle",
    ]

    if getattr(sys, "frozen", False):
        candidates.insert(
            0,
            Path(sys.executable).resolve().parent,
        )

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.insert(0, Path(meipass))

    return _unique_paths(candidates)


def find_tesseract() -> Path:
    override = os.environ.get("SHIPMENT_BOT_TESSERACT", "").strip()

    candidates: list[Path] = []

    if override:
        candidates.append(Path(override))

    for base in application_directories():
        candidates.extend(
            [
                base / "Tesseract-OCR" / "tesseract.exe",
                base / "bundle" / "Tesseract-OCR" / "tesseract.exe",
            ]
        )

    candidates.extend(
        [
            Path("C:/Program Files/Tesseract-OCR/tesseract.exe"),
            Path("C:/Program Files (x86)/Tesseract-OCR/tesseract.exe"),
        ]
    )

    for candidate in _unique_paths(candidates):
        if candidate.is_file():
            return candidate

    checked = "\n".join(f" - {path}" for path in candidates)

    raise FileNotFoundError(
        "Tesseract OCR could not be found.\n"
        "Checked these locations:\n"
        f"{checked}"
    )


def find_tessdata() -> Path:
    tesseract = find_tesseract()
    tessdata = tesseract.parent / "tessdata"

    if tessdata.is_dir():
        return tessdata

    raise FileNotFoundError(
        f"Tesseract was found, but tessdata is missing: {tessdata}"
    )


def find_poppler() -> Path:
    override = os.environ.get("SHIPMENT_BOT_POPPLER", "").strip()

    candidates: list[Path] = []

    if override:
        candidates.append(Path(override))

    for base in application_directories():
        candidates.extend(
            [
                base / "poppler" / "bin",
                base / "bundle" / "poppler" / "bin",
                base / "poppler",
            ]
        )

    for candidate in _unique_paths(candidates):
        if (
            candidate.is_dir()
            and (
                (candidate / "pdftoppm.exe").is_file()
                or (candidate / "pdftocairo.exe").is_file()
            )
        ):
            return candidate

    checked = "\n".join(f" - {path}" for path in candidates)

    raise FileNotFoundError(
        "Poppler could not be found.\n"
        "Checked these locations:\n"
        f"{checked}"
    )