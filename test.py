import os
import re
import pytesseract

from pdf2image import convert_from_path
from PIL import Image


pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tesseract.exe"
)

os.environ["TESSDATA_PREFIX"] = (
    r"C:\Users\aayan.boradia\Downloads\Tesseract-OCR\tessdata"
)


POPPLER_PATH = (
    r"C:\Users\aayan.boradia\Downloads\poppler-26.02.0\Library\bin"
)


def _ocr_image(image_path):
    img = Image.open(image_path)

    # improve OCR for small shipping labels
    img = img.convert("L")
    img = img.resize(
        (img.width * 3, img.height * 3)
    )

    return pytesseract.image_to_string(
        img,
        config="--oem 3 --psm 6"
    )


def extract_sender_company(text):
    """
    Used only by extractor_core.py to decide excel sheet
    """

    text_upper = text.upper()

    if "EMBY INTERNATIONAL" in text_upper:
        return "EMBY"

    if "FENIX" in text_upper:
        return "FENIX"

    if (
        "UNI" in text_upper
        or "UNIVERSAL" in text_upper
    ):
        return "UNI"

    return ""



def extract_recipient_company(text):

    lines = [
        x.strip()
        for x in text.split("\n")
        if x.strip()
    ]


    # Look after "To:"
    for i, line in enumerate(lines):

        if re.match(
            r"TO\s*:",
            line,
            re.IGNORECASE
        ):

            # Skip location line
            for next_line in lines[i+1:]:

                if (
                    next_line.upper()
                    not in [
                        "LAFAYETTE",
                        "NEW YORK",
                    ]
                    and not re.search(
                        r"\d",
                        next_line
                    )
                ):
                    return next_line.strip()


    # fallback: look after sender address
    found_sender = False

    for line in lines:

        if "INC" in line.upper():
            found_sender = True
            continue

        if found_sender:

            if (
                len(line) > 3
                and not re.search(
                    r"\d",
                    line
                )
            ):
                return line


    return ""



def extract_date(text):

    patterns = [

        r"Date:\s*\n?([A-Za-z]{3}\s+\d{1,2},\s+\d{4})",

        r"Req\.?\s*Pickup\s*Date:\s*\n?([A-Za-z]{3}\s+\d{1,2},\s+\d{4})",

        r"Pickup\s*Date:\s*\n?([A-Za-z]{3}\s+\d{1,2},\s+\d{4})"

    ]


    for p in patterns:

        m = re.search(
            p,
            text,
            re.IGNORECASE
        )

        if m:
            return m.group(1).strip()


    return ""



def extract_tracking_number(text):

    # Parcel number at bottom
    # Example:
    # MALCA-AMIT
    #
    # Parcel
    #
    # 73223275 1 of 1

    matches = re.findall(
        r"\b\d{7,12}\b",
        text
    )


    for number in matches:

        # avoid phone numbers
        if number not in text.replace(
            " ",
            ""
        ):

            return number


    if matches:
        return matches[-1]


    return ""



def extract_invoice_po(text):

    result = {
        "PO Number": "",
        "INV Number": "",
        "Memo Number": ""
    }


    patterns = {

        "PO Number":
            r"\bP\.?\s*O\.?\s*[:#]?\s*([A-Z0-9-]+)",

        "INV Number":
            r"\bINV(?:OICE)?[:#]?\s*([A-Z0-9-]+)",

        "Memo Number":
            r"\bMEMO[:#]?\s*([A-Z0-9-]+)"
    }


    for key, pattern in patterns.items():

        m = re.search(
            pattern,
            text,
            re.IGNORECASE
        )

        if m:
            result[key] = m.group(1)


    return result



def extract_fields(text):

    invoice_data = extract_invoice_po(text)


    data = {

        "Sender Company":
            extract_sender_company(text),

        "Recipient Company":
            extract_recipient_company(text),

        "Shipment Date":
            extract_date(text),

        "Tracking Number":
            extract_tracking_number(text),

        "PO Number":
            invoice_data["PO Number"],

        "INV Number":
            invoice_data["INV Number"],

        "Memo Number":
            invoice_data["Memo Number"],

    }


    return data



def get_malka_brinx_data(pdf_path):

    images = convert_from_path(
        pdf_path,
        poppler_path=POPPLER_PATH
    )


    results = []


    for i, img in enumerate(images):

        img_path = (
            f"{pdf_path}_page_{i}.png"
        )

        img.save(img_path)


        text = _ocr_image(
            img_path
        )


        results.append(
            extract_fields(text)
        )


        os.remove(img_path)


    return results