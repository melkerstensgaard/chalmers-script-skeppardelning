# Detta är en anpassad version av examensbevis-skriptet
# för serien med skepparexamen / skepparintyg.
#
# Syfte:
# - Identifiera startsidor för varje skepparexamensbevis
# - Splitta PDF:er till ett bevis per PDF
# - Extrahera bevisnummer
# - Extrahera personnummer så långt OCR-kvaliteten tillåter
# - Skapa index.xlsx och validation_report.txt
#
# Körning:
# python split_skepparexamen_certificates.py <input_root> <output_root>

import os
import re
import sys
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed

from tqdm import tqdm
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook


EXCEL_FILENAME = "index.xlsx"
VALIDATION_FILENAME = "validation_report.txt"
UNREADABLE_LOG = "unreadable_volumes.log"


# --------------------------------------------------
# LOGGNING
# --------------------------------------------------

def emit_log(message, log_callback=None):
    if log_callback:
        log_callback(message)
    else:
        print(message)


# --------------------------------------------------
# CONFIG
# --------------------------------------------------

def get_default_config():
    return {
        # Startsidesmarkörer för nya serien
        "start_page_markers": [
            "BETYG",
            "ÖVER AVLAGD SKEPPAREXAMEN",
            "OVER AVLAGD SKEPPAREXAMEN",
            "BEVIS ÖVER SKEPPAREXAMEN",
            "BEVIS OVER SKEPPAREXAMEN",
            "SKEPPAREXAMEN HAR AVLAGTS AV",
        ],

        # Starka icke-startsidor
        "non_start_page_patterns": [
            "PRÖVNING FÖR SKEPPAREXAMEN",
            "PROVNING FOR SKEPPAREXAMEN",
            "IFYLLES AV EXAMINANDEN",
            "IFYLLES AV EXAMINATORN",
            "IFYLLES AV LÄRAREN",
            "IFYLLES AV LARAREN",
            "EXAMINANDENS NAMNTECKNING",
            "AVGIFT ERLAGD",
            "OMRÅDE SKRIFTL",
            "OMRADE SKRIFTL",
        ],

        # Om True hoppar hela volymer över om de inte verkar läsbara
        "skip_unreadable_volumes": False,

        # Tillåt om en startsida innehåller personnummer.
        # I denna serie SKA startsidan ofta innehålla personnummer,
        # så detta måste vara True.
        "allow_personnummer_on_start_page": True,

        # Hur många workers som används parallellt.
        # Sänk om minnet tar slut.
        "max_workers": 2,

        # Hur många PDF:er som skickas till executor åt gången.
        "batch_size": 5,
    }


# --------------------------------------------------
# OCR-NORMALISERING
# --------------------------------------------------

def normalize_ocr_text(text):
    """
    Normaliserar OCR-text för markörsökning.
    Detta är en generell textnormalisering, inte personnummernormalisering.
    """

    if not text:
        return ""

    t = text.upper()

    t = t.replace("\r", " ")
    t = t.replace("\n", " ")

    # Normalisera bindestreck
    t = t.replace("–", "-")
    t = t.replace("—", "-")
    t = t.replace("−", "-")

    # Förenkla vissa svenska tecken för robustare jämförelser
    t = t.replace("Ö", "O")
    t = t.replace("Ä", "A")
    t = t.replace("Å", "A")

    # Ta bort störande interpunktion
    t = re.sub(r"[;:,.]+", " ", t)
    t = re.sub(r"\s+", " ", t)

    return t.strip()


def compact_ocr_text(text):
    """
    Tar bort allt utom bokstäver och siffror.
    Används för att hitta markörer även när OCR har tappat mellanslag.
    """

    t = normalize_ocr_text(text)

    return re.sub(
        r"[^A-Z0-9]+",
        "",
        t
    )


# --------------------------------------------------
# PERSONNUMMER-NORMALISERING
# --------------------------------------------------

OCR_DIGIT_REPLACEMENTS = {
    "O": "0",
    "Q": "0",
    "D": "0",
    "Ö": "0",

    "I": "1",
    "L": "1",
    "|": "1",
    "!": "1",

    "Z": "2",

    "S": "5",

    "G": "6",
    "É": "6",

    "T": "7",

    "B": "8",
}

BLACKLIST_NAMES = [

    "STAFFAN WILSKE",
    "KURT JONASSON",
    "STIG HOLMSTRÖM",
    "ANDERS HASSLING",
    "BJÖRN BORG",
    "GÖRAN LINDHOLM",
    "LENNART JOHNSSON",

]

def normalize_ocr_digit_text(text):
    """
    Normaliserar OCR-skadade kandidatsträngar där tecken troligen ska vara siffror.

    Viktigt:
    Kör detta på korta kandidatfält, inte på hela sidtexten.
    """

    if not text:
        return ""

    t = text.upper()

    t = t.replace("–", "-")
    t = t.replace("—", "-")
    t = t.replace("−", "-")

    for wrong, correct in OCR_DIGIT_REPLACEMENTS.items():
        t = t.replace(wrong, correct)

    return t


def valid_date_yymmdd(digits6):
    """
    Enkel datumkontroll för YYMMDD.
    Används bara för att minska falska träffar.
    """

    if not re.fullmatch(r"\d{6}", digits6):
        return False

    yy = int(digits6[:2])
    mm = int(digits6[2:4])
    dd = int(digits6[4:6])

    if mm < 1 or mm > 12:
        return False

    if dd < 1 or dd > 31:
        return False

    return True


def format_pnr_from_digits(digits):
    """
    Tar 10 siffror och returnerar YYMMDD-NNNN.
    """

    if len(digits) != 10:
        return ""

    return f"{digits[:6]}-{digits[6:]}"


def extract_10_digit_windows(digits):
    """
    Tar fram möjliga 10-siffriga personnummer ur en längre siffersträng.

    Exempel:
    06308060773 kan innehålla 630806-0773 om OCR har lagt till en extra nolla.
    """

    results = []

    if len(digits) == 10:
        if valid_date_yymmdd(digits[:6]):
            results.append(format_pnr_from_digits(digits))

    if len(digits) > 10:
        for i in range(0, len(digits) - 9):
            window = digits[i:i + 10]

            if valid_date_yymmdd(window[:6]):
                results.append(format_pnr_from_digits(window))

    return results


def find_personnummer_snippets(text, window_lines=3):
    """
    Hittar korta textsnuttar runt raden där PERSONNUMMER förekommer.
    Detta är viktigt för handskrivna och OCR-skadade blanketter.
    """

    if not text:
        return []

    lines = text.splitlines()

    snippets = []

    for i, line in enumerate(lines):
        upper = line.upper()

        if (
            "PERSONNUMMER" in upper
            or "PERSONNR" in upper
            or "PERSNR" in upper
        ):
            snippet = " ".join(
                lines[i:i + window_lines]
            )

            snippets.append(snippet.strip())

    return snippets


def extract_personnummer_candidates(text):
    """
    Letar ENDAST i rader som innehåller PERSONNUMMER.

    Returnerar kandidater i prioriterad ordning.
    """

    if not text:
        return []

    candidates = []

    lines = text.splitlines()

    for line in lines:

        upper = line.upper()

        if not re.search(
                r"P[AE]R?SON",
                upper
        ):
            continue

        normalized = normalize_ocr_digit_text(line)

        # försök hitta klassiskt format
        direct_match = re.search(
            r"(\d{6})[- ]?(\d{4})",
            normalized
        )

        if direct_match:

            candidates.append(
                f"{direct_match.group(1)}-{direct_match.group(2)}"
            )

            continue

        # fallback:
        # plocka endast siffror från raden
        digits = re.sub(
            r"[^0-9]",
            "",
            normalized
        )

        if len(digits) == 10:

            candidates.append(
                f"{digits[:6]}-{digits[6:]}"
            )

        elif len(digits) == 11:

            # vanligt OCR-fel:
            # extra nolla först

            if digits.startswith("0"):

                d = digits[1:]

                candidates.append(
                    f"{d[:6]}-{d[6:]}"
                )

    return candidates
# --------------------------------------------------
# BEVISNUMMER
# --------------------------------------------------

def extract_bevisnummer_candidates(text):
    """
    Extraherar bevisnummer från sidor.

    Exempel som ska matcha:
    Bevisnr 17798
    Nr 17818
    Nr 25021
    NR 25051

    Exempel som INTE ska matcha:
    SÖ 85-021
    85-021
    """

    if not text:
        return []

    t = normalize_ocr_text(text)

    candidates = []

    patterns = [
        r"\bBEVIS\s*NR\s*\.?\s*(\d{4,6})\b",
        r"\bBEVISNR\s*\.?\s*(\d{4,6})\b",
        r"\bNR\s*\.?\s*(\d{4,6})\b",
        r"\bNR\s+(\d{4,6})\b",
    ]

    for pattern in patterns:
        candidates.extend(
            re.findall(
                pattern,
                t,
                flags=re.IGNORECASE
            )
        )

    cleaned = []

    for value in candidates:
        digits = re.sub(r"[^0-9]", "", value)

        if not digits:
            continue

        # Undvik mallnummer som 85-021.
        # Bevisnummer i denna serie verkar ligga som 5-siffriga nummer
        # eller äldre 4-5-siffriga nummer.
        if 1000 <= int(digits) <= 99999:
            cleaned.append(digits)

    return sorted(set(cleaned))

# ---------
# FÖR OCH EFTERNAMN
# -------------

def extract_name_candidates(text):

    if not text:
        return []

    candidates = []

    lines = text.splitlines()

    markers = [
        "PERSONNR",
        "PERSONNUMMER",
        "PARSONNR",
        "PARSONNUMMER"
    ]

    for i, line in enumerate(lines):

        upper = line.upper()

        if not any(m in upper for m in markers):
            continue

        # mycket större sökfönster
        window = lines[max(0, i - 5): min(len(lines), i + 5)]

        for candidate in window:

            candidate = candidate.strip()

            if len(candidate) < 5:
                continue

            # ta bort personnummer om de ligger på raden
            candidate = re.sub(
                r"\d{6}[- ]?\d{4}",
                "",
                candidate
            )

            candidate = candidate.strip()

            words = candidate.split()

            if len(words) < 2:
                if is_blacklisted_name(candidate): 4
                continue

            # alla ord måste vara bokstavsliknande
            if all(
                re.fullmatch(
                    r"[A-Za-zÅÄÖåäö\-]+",
                    w
                )
                for w in words
            ):

                bad_words = [

                    "ADJUNKT",
                    "LEKTOR",
                    "TIMLÄRARE",
                    "SJÖKAPTEN",
                    "EXAMENSFÖRRÄTTARE",
                    "LÄRARE"

                ]

                if any(
                        word in candidate.upper()
                        for word in bad_words
                ):
                    continue
                candidates.append(candidate)

    return list(dict.fromkeys(candidates))


# --------------------------------------------------
# SIDTYPER
# --------------------------------------------------

def is_application_form(text, config):
    """
    Identifierar ansöknings- eller prövningsblanketter.
    Dessa ska normalt inte vara startsida för ett bevis.
    """

    t = normalize_ocr_text(text)

    for marker in config["non_start_page_patterns"]:
        marker_norm = normalize_ocr_text(marker)

        if marker_norm in t:
            return True

    compact = compact_ocr_text(text)

    compact_markers = [
        "PROVNINGFORSKEPPAREXAMEN",
        "IFYLLESAVEXAMINANDEN",
        "IFYLLESAVLARAREN",
        "EXAMINANDENSNAMNTECKNING",
    ]

    return any(marker in compact for marker in compact_markers)


def is_start_page(text, config):
    """
    Avgör om sidan är startsida för ett skepparexamensbevis.

    Den nya serien använder främst:
    - BETYG över avlagd skepparexamen
    - BEVIS ÖVER SKEPPAREXAMEN
    - Skepparexamen har avlagts av
    - Nr / Bevisnr
    """

    if not text:
        return False

    t = normalize_ocr_text(text)
    compact = compact_ocr_text(text)

    # Ansökningsblanketter är inte bevisstartsidor
    if is_application_form(text, config):
        return False

    score = 0

    # Starka startmarkörer
    if "SKEPPAREXAMEN HAR AVLAGTS AV" in t:
        score += 4

    if "BEVIS OVER SKEPPAREXAMEN" in t:
        score += 4

    if "BETYG" in t and "SKEPPAREXAMEN" in t:
        score += 3

    if "OVER AVLAGD SKEPPAREXAMEN" in t:
        score += 2

    # Kompakt fallback för OCR utan mellanslag
    if "SKEPPAREXAMENHARAVLAGTSAV" in compact:
        score += 4

    if "BEVISOVERSKEPPAREXAMEN" in compact:
        score += 4

    if "BETYG" in compact and "SKEPPAREXAMEN" in compact:
        score += 3

    # Bevisnummer stärker startsidesbedömningen
    bevis_candidates = extract_bevisnummer_candidates(text)

    if bevis_candidates:
        score += 2

    # Personnr/personnummerlabel förekommer ofta på bevisstartsidor
    if "PERSONNR" in t or "PERSONNUMMER" in t:
        score += 1

    return score >= 4


# --------------------------------------------------
# OCR CHECK
# --------------------------------------------------

def is_pdf_readable(reader):
    """
    Enkel kontroll att PDF:en har textlager.
    """

    try:
        for i in range(min(3, len(reader.pages))):
            text = reader.pages[i].extract_text() or ""

            if text and len(text.strip()) > 20:
                return True

    except Exception:
        return False

    return False


def has_certificates(reader, config):
    """
    Kontrollerar om PDF:en verkar innehålla minst en startsida.
    """

    try:
        for page in reader.pages:
            text = page.extract_text() or ""

            if is_start_page(text, config):
                return True

    except Exception:
        return False

    return False


# --------------------------------------------------
# HJÄLPFUNKTIONER
# --------------------------------------------------

def safe_filename_part(value):
    """
    Gör sträng säker för filnamn.
    """

    if not value:
        return ""

    value = re.sub(r"[^A-Za-z0-9ÅÄÖåäö_-]+", "_", value)

    return value.strip("_")

def is_blacklisted_name(candidate):

    candidate_upper = candidate.upper()

    return any(
        blocked in candidate_upper
        for blocked in BLACKLIST_NAMES
    )

# --------------------------------------------------
# PROCESS ONE PDF
# --------------------------------------------------

def process_pdf(args):
    pdf_path, input_root, output_root, config = args

    rows = []
    validation = []

    try:
        reader = PdfReader(pdf_path)

    except Exception as e:
        return [], [
            f"❌ Cannot open: {pdf_path} | {type(e).__name__}: {e}"
        ]

    start_pages = []

    for i, page in enumerate(reader.pages):
        text = page.extract_text() or ""

        if is_start_page(text, config):
            start_pages.append(i)

            validation.append(
                f"START PAGE detected on page {i + 1} in {pdf_path}"
            )

    if not start_pages:
        validation.append(
            f"⚠️ No start pages found: {pdf_path}"
        )

        return [], validation

    start_pages.append(len(reader.pages))

    if len(start_pages) > 2:
        validation.append(
            f"Split into {len(start_pages) - 1} certificates: {pdf_path}"
        )

    rel_dir = os.path.relpath(
        os.path.dirname(pdf_path),
        input_root
    )

    output_dir = os.path.join(
        output_root,
        rel_dir
    )

    os.makedirs(
        output_dir,
        exist_ok=True
    )

    base = os.path.splitext(
        os.path.basename(pdf_path)
    )[0]

    volym = rel_dir.split(os.sep)[0]

    for i in range(len(start_pages) - 1):

        start = start_pages[i]
        end = start_pages[i + 1]

        pnr = ""

        all_pnrs = []
        all_names = []
        all_bevisnummer = []


        for p in range(start, end):

            try:
                text = reader.pages[p].extract_text() or ""

            except Exception:
                continue

            page_pnrs = extract_personnummer_candidates(text)

            all_pnrs.extend(page_pnrs)

            page_names = extract_name_candidates(text)

            all_names.extend(page_names)

            # Ta första träffen istället för consensus
            if not pnr and page_pnrs:
                pnr = page_pnrs[0]

                validation.append(
                    f"Selected PNR {pnr} from page {p + 1}"
                )

            all_bevisnummer.extend(
                extract_bevisnummer_candidates(text)
            )

        ocr_name = ""

        if all_names:
            ocr_name = all_names[0]

        bevisnummer = ""

        if all_bevisnummer:
            bevisnummer = all_bevisnummer[0]

        if not pnr:
            validation.append(
                f"⚠️ Missing personnummer: {pdf_path} pages {start + 1}-{end}"
            )

        if not bevisnummer:
            validation.append(
                f"⚠️ Missing bevisnummer: {pdf_path} pages {start + 1}-{end}"
            )

        unique_pnrs = sorted(set(all_pnrs))
        unique_bevisnummer = sorted(set(all_bevisnummer))

        if len(unique_pnrs) > 1:
            validation.append(
                f"⚠️ Multiple personnummer candidates in split: {pdf_path} pages {start + 1}-{end} | {unique_pnrs}"
            )

        if len(unique_bevisnummer) > 1:
            validation.append(
                f"⚠️ Multiple bevisnummer candidates in split: {pdf_path} pages {start + 1}-{end} | {unique_bevisnummer}"
            )

        writer = PdfWriter()

        for p in range(start, end):
            writer.add_page(reader.pages[p])

        bnr_part = safe_filename_part(bevisnummer)

        if bnr_part:
            out_name = f"{base}_{i + 1:03d}_bevis_{bnr_part}.pdf"
        else:
            out_name = f"{base}_{i + 1:03d}.pdf"

        out_path = os.path.join(
            output_dir,
            out_name
        )

        try:
            with open(out_path, "wb") as f:
                writer.write(f)

        except Exception as e:
            validation.append(
                f"❌ Write failed: {out_path} | {type(e).__name__}: {e}"
            )

            continue

        # Validera att splitten innehåller exakt en startsida
        start_count = 0

        for p in range(start, end):
            try:
                text = reader.pages[p].extract_text() or ""

            except Exception:
                continue

            if is_start_page(text, config):
                start_count += 1

        if start_count != 1:
            validation.append(
                f"⚠️ Suspicious split start pages={start_count}: {out_path}"
            )

        rows.append(
            [
                out_name,
                volym,
                bevisnummer,
                ocr_name,
                pnr,
                ";".join(unique_bevisnummer),
                ";".join(unique_pnrs),
                ";".join(all_names),
                start + 1,
                end,
            ]
        )

    return rows, validation


# --------------------------------------------------
# MAIN
# --------------------------------------------------

def process_all(
    input_root,
    output_root,
    config,
    log_callback=None,
    progress_callback=None
):

    volumes = {}

    for root, _, files in os.walk(input_root):

        pdfs = [
            os.path.join(root, f)
            for f in files
            if f.lower().endswith(".pdf")
        ]

        if pdfs:
            volumes[root] = pdfs

    emit_log(
        f"📦 Found {len(volumes)} volumes",
        log_callback
    )

    valid_pdfs = []
    skipped_volumes = []

    for volume, pdf_list in volumes.items():

        volume_ok = True

        for pdf in pdf_list:

            try:
                reader = PdfReader(pdf)

            except Exception:
                volume_ok = False
                break

            if config.get("skip_unreadable_volumes", False):

                if not is_pdf_readable(reader):
                    volume_ok = False
                    break

                if not has_certificates(reader, config):
                    volume_ok = False
                    break

        if volume_ok:
            valid_pdfs.extend(pdf_list)

        else:
            skipped_volumes.append(volume)

    emit_log(
        f"✅ Valid PDFs: {len(valid_pdfs)}",
        log_callback
    )

    emit_log(
        f"⛔ Skipped volumes: {len(skipped_volumes)}",
        log_callback
    )

    total_pdfs = len(valid_pdfs)
    completed_pdfs = 0

    if progress_callback:
        progress_callback(
            0,
            total_pdfs
        )

    all_rows = []
    validation = []
    all_futures = []

    batch_size = config.get(
        "batch_size",
        5
    )

    max_workers = config.get(
        "max_workers",
        2
    )

    with ProcessPoolExecutor(max_workers=max_workers) as executor:

        batch = []

        for pdf_path in valid_pdfs:

            batch.append(
                (
                    pdf_path,
                    input_root,
                    output_root,
                    config
                )
            )

            if len(batch) == batch_size:

                futures = [
                    executor.submit(
                        process_pdf,
                        arg
                    )
                    for arg in batch
                ]

                all_futures.extend(futures)

                batch = []

        if batch:

            futures = [
                executor.submit(
                    process_pdf,
                    arg
                )
                for arg in batch
            ]

            all_futures.extend(futures)

        for future in tqdm(
            as_completed(all_futures),
            total=len(all_futures),
            desc="Processing"
        ):

            rows, val = future.result()

            all_rows.extend(rows)
            validation.extend(val)

            completed_pdfs += 1

            if progress_callback:
                progress_callback(
                    completed_pdfs,
                    total_pdfs
                )

            emit_log(
                f"Processed {completed_pdfs}/{total_pdfs}",
                log_callback
            )

    os.makedirs(
        output_root,
        exist_ok=True
    )

    # Excel
    wb = Workbook()
    ws = wb.active

    ws.append(
        [
            "file name",
            "volym",
            "bevisnummer",
            "ocr_name",
            "ocr_personnummer",
            "all_bevisnummer_candidates",
            "all_personnummer_candidates",
            "all_name_candidates",
            "start_page",
            "end_page_exclusive",
        ]
    )

    for row in all_rows:
        ws.append(row)

    excel_path = os.path.join(
        output_root,
        EXCEL_FILENAME
    )

    wb.save(excel_path)

    validation_path = os.path.join(
        output_root,
        VALIDATION_FILENAME
    )

    with open(
        validation_path,
        "w",
        encoding="utf-8"
    ) as f:

        for item in validation:
            f.write(item + "\n")

    unreadable_path = os.path.join(
        output_root,
        UNREADABLE_LOG
    )

    with open(
        unreadable_path,
        "w",
        encoding="utf-8"
    ) as f:

        for item in skipped_volumes:
            f.write(item + "\n")

    emit_log(
        f"\n📊 Excel: {excel_path}",
        log_callback
    )

    emit_log(
        f"📋 Validation: {validation_path}",
        log_callback
    )

    emit_log(
        f"🚫 Skipped volumes: {unreadable_path}",
        log_callback
    )

    emit_log(
        "🎉 Done!",
        log_callback
    )


# --------------------------------------------------
# ENTRY
# --------------------------------------------------

if __name__ == "__main__":

    if len(sys.argv) != 3:
        print(
            "Usage: python split_skepparexamen_certificates.py <input_root> <output_root>"
        )

        sys.exit(1)

    config = get_default_config()

    process_all(
        sys.argv[1],
        sys.argv[2],
        config
    )