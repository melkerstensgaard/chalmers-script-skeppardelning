import re
import fitz
from PIL import Image

PNR = re.compile(r"(?<!\d)(\d{6}|\d{8})[-+ ]?(\d{4})(?!\d)")

def analyze(text):
    match = PNR.search(text)
    pnr = match.group(1) + "-" + match.group(2) if match else ""
    score = 45 if pnr else 0
    reasons = [f"Personnummer hittat: {pnr}" if pnr else "Inget personnummer hittades"]
    for word, value in {"bevis":20,"intyg":15,"prov":-8,"bilaga":-5}.items():
        if word in text.lower():
            score += value; reasons.append(f"{word}: {value:+d} poäng")
    return max(0,min(100,score)), pnr, reasons

def extract_text(page):
    text = page.get_text("text") or ""
    if len(text.strip()) > 30:
        return text
    try:
        import pytesseract
        pix = page.get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
        image = Image.frombytes("RGB", [pix.width,pix.height], pix.samples)
        return pytesseract.image_to_string(image, lang="swe+eng")
    except Exception:
        return text
