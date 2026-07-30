"""
Validerar OCR-resultat mot det auktoritativa registret.

Input:
    index.xlsx

    Skeppare B - fartygsbefäl klass VIII 1980-2000.xlsx

Output:
    validation_output.xlsx
"""

import pandas as pd
from rapidfuzz import fuzz


OCR_FILE = "index.xlsx"
REGISTER_FILE = "Skeppare B - fartygsbefäl klass VIII 1980-2000.xlsx"

OUTPUT_FILE = "validation_output.xlsx"


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

    if pd.isna(value):
        return ""

    try:
        return str(int(float(value)))
    except Exception:
        return str(value).strip()


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
ocr_df = pd.read_excel(OCR_FILE)

print("Reading register file...")
register_df = pd.read_excel(REGISTER_FILE)

print("\nREGISTER COLUMNS:")
print(register_df.columns.tolist())
print()

# --------------------------------------------------
# HÄR MÅSTE DU EVENTUELLT JUSTERA
# OM KOLUMNNAMNEN ÄR ANNORLUNDA
# --------------------------------------------------

REGISTER_BEVISNR_COLUMN = "bevisnummer"
REGISTER_NAME_COLUMN = "namn"
REGISTER_PNR_COLUMN = "personnummer"

OCR_BEVISNR_COLUMN = "bevisnummer"
OCR_NAME_COLUMN = "ocr_name"
OCR_PNR_COLUMN = "ocr_personnummer"


# --------------------------------------------------
# LOOKUP PÅ BEVISNUMMER
# --------------------------------------------------

register_lookup = {}

for _, row in register_df.iterrows():

    bevisnr = clean_bevisnummer(
        row[REGISTER_BEVISNR_COLUMN]
    )

    if bevisnr:
        register_lookup[bevisnr] = row

print(
    f"Loaded {len(register_lookup)} register records"
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

    bevisnr = clean_bevisnummer(
        row[OCR_BEVISNR_COLUMN]
    )

    ocr_name = clean_text(
        row[OCR_NAME_COLUMN]
    )

    ocr_pnr = clean_personnummer(
        row[OCR_PNR_COLUMN]
    )

    if bevisnr not in register_lookup:

        results.append({

            "bevisnummer": bevisnr,

            "ocr_name": ocr_name,

            "ocr_personnummer": ocr_pnr,

            "register_name": "",

            "register_personnummer": "",

            "name_similarity": 0,

            "pnr_match": False,

            "status": "NOT FOUND IN REGISTER"
        })

        not_found += 1

        continue

    reg = register_lookup[bevisnr]

    reg_name = clean_text(
        reg[REGISTER_NAME_COLUMN]
    )

    reg_pnr = clean_personnummer(
        reg[REGISTER_PNR_COLUMN]
    )

    # fuzzy match namn
    similarity = fuzz.token_sort_ratio(
        ocr_name,
        reg_name
    )

    pnr_ok = (
        ocr_pnr == reg_pnr
    )

    if similarity >= 90 and pnr_ok:

        status = "MATCH"
        matched += 1

    elif not pnr_ok:

        status = "PNR MISMATCH"
        pnr_mismatch += 1

    else:

        status = "NAME MISMATCH"
        name_mismatch += 1

    results.append({

        "bevisnummer": bevisnr,

        "ocr_name": ocr_name,

        "register_name": reg_name,

        "name_similarity": similarity,

        "ocr_personnummer": ocr_pnr,

        "register_personnummer": reg_pnr,

        "pnr_match": pnr_ok,

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