import re
from datetime import date

PATTERNS = [
    re.compile(r"(?<!\d)(\d{8})\s*[-–—+ ]?\s*(\d{4})(?!\d)"),
    re.compile(r"(?<!\d)(\d{6})\s*[-–—+ ]?\s*(\d{4})(?!\d)"),
]
TRANSLATION = str.maketrans({"O":"0", "o":"0", "I":"1", "l":"1", "|":"1", "S":"5", "B":"8", "Z":"2"})

def find_candidates(text):
    normalized_text = (text or "").translate(TRANSLATION)
    found, seen = [], set()
    for pattern in PATTERNS:
        for match in pattern.finditer(normalized_text):
            value = f"{match.group(1)}-{match.group(2)}"
            if value not in seen:
                seen.add(value)
                found.append((match.group(0), value))
    return found

def valid_date(value):
    digits = re.sub(r"\D", "", value or "")
    try:
        if len(digits) == 12:
            year, month, day = int(digits[:4]), int(digits[4:6]), int(digits[6:8])
        elif len(digits) == 10:
            yy, month, day = int(digits[:2]), int(digits[2:4]), int(digits[4:6])
            year = 2000 + yy if yy <= date.today().year % 100 else 1900 + yy
        else:
            return False
        if day > 60:
            day -= 60
        date(year, month, day)
        return True
    except ValueError:
        return False

def valid_luhn(value):
    digits = re.sub(r"\D", "", value or "")
    if len(digits) == 12:
        digits = digits[2:]
    if len(digits) != 10:
        return False
    total = 0
    for index, char in enumerate(digits):
        number = int(char) * (2 if index % 2 == 0 else 1)
        total += number // 10 + number % 10
    return total % 10 == 0

def score_candidate(value, source_count=1, average_ocr_confidence=0.0):
    score, reasons = 0, []
    digits = re.sub(r"\D", "", value or "")
    if len(digits) in (10, 12):
        score += 15; reasons.append("Rätt antal siffror +15")
    if valid_date(value):
        score += 15; reasons.append("Giltigt datum +15")
    if valid_luhn(value):
        score += 25; reasons.append("Giltig kontrollsiffra +25")
    if source_count >= 2:
        score += 15; reasons.append("Samma kandidat från minst två analysvägar +15")
    if source_count >= 3:
        score += 10; reasons.append("Samma kandidat från minst tre analysvägar +10")
    if average_ocr_confidence >= 80:
        score += 10; reasons.append("Hög OCR-konfidens +10")
    elif average_ocr_confidence and average_ocr_confidence < 40:
        score -= 10; reasons.append("Låg OCR-konfidens -10")
    return max(0, min(100, score)), reasons
