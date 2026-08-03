"""
Validerar OCR-resultat mot det auktoritativa registret.

Input:
    index.xlsx

    Skeppare B - fartygsbefäl klass VIII 1980-2000.xlsx

Output:
    validation_output.xlsx
"""

import re
import pandas as pd
from rapidfuzz import fuzz
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

DATA_DIR = PROJECT_ROOT / "data"

OCR_FILE = DATA_DIR / "output" / "index.xlsx"

REGISTER_FILE = (
    DATA_DIR
    / "authoritative_register"
    / "Skeppare B - fartygsbefäl klass VIII 1980-2000.xlsx"
)

OUTPUT_FILE = (
    DATA_DIR
    / "output"
    / "validation_output.xlsx"
)


# --------------------------------------------------
# HJÄLPFUNKTIONER
# --------------------------------------------------

def clean_text(value):

    if pd.isna(value):
        return ""

    return (
        str(value)
        .strip()
        .upper()
    )


def clean_bevisnummer(value):

    if value is None:
        return ""

    if pd.isna(value):
        return ""

    text = str(value).strip()

    if not text:
        return ""

    if text.lower() in ["nan", "none"]:
        return ""

    # Hanterar Excel-värden som 17853.0
    try:
        return str(int(float(text)))
    except Exception:
        pass

    # Fallback: behåll bara siffror
    digits = re.sub(
        r"[^0-9]",
        "",
        text
    )

    return digits


def clean_personnummer(value):

    if pd.isna(value):
        return ""

    text = str(value)

    return (
        text.replace(" ", "")
            .strip()
    )


# --------------------------------------------------
# LÄS FILER
# --------------------------------------------------

print("Reading OCR file...")
ocr_df = pd.read_excel(
    OCR_FILE,
    dtype=str
)

print("Reading register file...")
register_df = pd.read_excel(
    REGISTER_FILE,
    dtype=str
)

print("\nREGISTER COLUMNS:")
print(register_df.columns.tolist())
print("OCR COLUMNS:")
print(ocr_df.columns.tolist())
print()
print()

# --------------------------------------------------
# HÄR MÅSTE DU EVENTUELLT JUSTERA
# OM KOLUMNNAMNEN ÄR ANNORLUNDA
# --------------------------------------------------

REGISTER_BEVISNR_COLUMN = "Intygsnummer"
REGISTER_LASTNAME_COLUMN = "Efternamn"
REGISTER_FIRSTNAME_COLUMN = "Förnamn"
REGISTER_PNR_COLUMN = "Personnummer"

OCR_RAW_BEVISNR_COLUMN = "ocr_bevisnummer"
OCR_RESOLVED_BEVISNR_COLUMN = "resolved_bevisnummer"
OCR_MATCH_METHOD_COLUMN = "match_method"

OCR_NAME_COLUMN = "ocr_name"
OCR_SPLIT_REGISTER_NAME_COLUMN = "register_name"
OCR_NAME_SCORE_COLUMN = "name_match_score"

OCR_PNR_COLUMN = "ocr_personnummer"
OCR_SPLIT_REGISTER_PNR_COLUMN = "register_personnummer"

# --------------------------------------------------
# LOOKUP PÅ BEVISNUMMER
# --------------------------------------------------

register_lookup = {}

for _, row in register_df.iterrows():

    bevisnr = clean_bevisnummer(
        row.get(REGISTER_BEVISNR_COLUMN, "")
    )

    if not bevisnr:
        continue

    register_lookup[bevisnr] = row

print(
    f"Loaded {len(register_lookup)} register records"
)

print(
    "First 10 register keys:",
    list(register_lookup.keys())[:10]
)

# --------------------------------------------------
# VALIDERING
# --------------------------------------------------

results = []

matched = 0
not_found = 0
pnr_mismatch = 0
name_mismatch = 0

for _, row in ocr_df.iterrows():

    ocr_bevisnr = clean_bevisnummer(
        row.get(OCR_RAW_BEVISNR_COLUMN, "")
    )

    resolved_bevisnr = clean_bevisnummer(
        row.get(OCR_RESOLVED_BEVISNR_COLUMN, "")
    )

    match_method = clean_text(
        row.get(OCR_MATCH_METHOD_COLUMN, "")
    )

    bevisnr = resolved_bevisnr

    ocr_name = clean_text(
        row.get(OCR_NAME_COLUMN, "")
    )

    split_register_name = clean_text(
        row.get(OCR_SPLIT_REGISTER_NAME_COLUMN, "")
    )

    name_score_from_split = row.get(
        OCR_NAME_SCORE_COLUMN,
        ""
    )

    ocr_pnr = clean_personnummer(
        row.get(OCR_PNR_COLUMN, "")
    )

    split_register_pnr = clean_personnummer(
        row.get(OCR_SPLIT_REGISTER_PNR_COLUMN, "")
    )

    if not bevisnr:
        results.append({

            "ocr_bevisnummer": ocr_bevisnr,
            "resolved_bevisnummer": resolved_bevisnr,
            "match_method": match_method,

            "ocr_name": ocr_name,
            "split_register_name": split_register_name,
            "validation_register_name": "",
            "name_score_from_split": name_score_from_split,

            "ocr_personnummer": ocr_pnr,
            "split_register_personnummer": split_register_pnr,
            "validation_register_personnummer": "",

            "ocr_pnr_match": False,
            "split_register_pnr_match": False,

            "status": "NO RESOLVED BEVISNUMMER"
        })

        not_found += 1
        continue

    if bevisnr not in register_lookup:
        results.append({

            "ocr_bevisnummer": ocr_bevisnr,
            "resolved_bevisnummer": resolved_bevisnr,
            "match_method": match_method,

            "ocr_name": ocr_name,
            "split_register_name": split_register_name,
            "validation_register_name": "",
            "name_score_from_split": name_score_from_split,

            "ocr_personnummer": ocr_pnr,
            "split_register_personnummer": split_register_pnr,
            "validation_register_personnummer": "",

            "ocr_pnr_match": False,
            "split_register_pnr_match": False,

            "status": "NOT FOUND IN REGISTER"
        })

        not_found += 1
        continue

    reg = register_lookup[bevisnr]

    validation_register_name = clean_text(
        f"{reg[REGISTER_LASTNAME_COLUMN]} {reg[REGISTER_FIRSTNAME_COLUMN]}"
    )

    validation_register_pnr = clean_personnummer(
        reg[REGISTER_PNR_COLUMN]
    )

    ocr_name_similarity = fuzz.token_sort_ratio(
        ocr_name,
        validation_register_name
    )

    split_register_name_match = (
            split_register_name == validation_register_name
    )

    ocr_pnr_match = (
            ocr_pnr == validation_register_pnr
    )

    split_register_pnr_match = (
            split_register_pnr == validation_register_pnr
    )

    if split_register_pnr_match:

        status = "VALIDATED"
        matched += 1

    elif ocr_pnr_match:

        status = "OCR PNR MATCHES REGISTER"
        matched += 1

    elif ocr_name_similarity >= 90:

        status = "NAME MATCH ONLY"
        name_mismatch += 1

    else:

        status = "PNR MISMATCH"
        pnr_mismatch += 1

    results.append({

        "ocr_bevisnummer": ocr_bevisnr,
        "resolved_bevisnummer": resolved_bevisnr,
        "match_method": match_method,

        "ocr_name": ocr_name,
        "split_register_name": split_register_name,
        "validation_register_name": validation_register_name,
        "split_register_name_match": split_register_name_match,
        "ocr_name_similarity": ocr_name_similarity,
        "name_score_from_split": name_score_from_split,

        "ocr_personnummer": ocr_pnr,
        "split_register_personnummer": split_register_pnr,
        "validation_register_personnummer": validation_register_pnr,

        "ocr_pnr_match": ocr_pnr_match,
        "split_register_pnr_match": split_register_pnr_match,

        "status": status
    })

# --------------------------------------------------
# SKRIV RESULTAT
# --------------------------------------------------

result_df = pd.DataFrame(results)

result_df.to_excel(
    OUTPUT_FILE,
    index=False
)

# --------------------------------------------------
# SAMMANFATTNING
# --------------------------------------------------

print()
print("========== RESULT ==========")
print(f"Total OCR rows:       {len(results)}")
print(f"Matches:              {matched}")
print(f"PNR mismatches:       {pnr_mismatch}")
print(f"Name mismatches:      {name_mismatch}")
print(f"Not in register:      {not_found}")
print("============================")
print()
print(
    f"Validation written to {OUTPUT_FILE}"
)