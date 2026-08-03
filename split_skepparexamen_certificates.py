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
from pathlib import Path
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd
from rapidfuzz import fuzz
from rapidfuzz.distance import Levenshtein

from tqdm import tqdm
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook


EXCEL_FILENAME = "index.xlsx"
VALIDATION_FILENAME = "validation_report.txt"
UNREADABLE_LOG = "unreadable_volumes.log"

PROJECT_ROOT = Path(__file__).resolve().parent

REGISTER_FILE = (
    PROJECT_ROOT
    / "data"
    / "authoritative_register"
    / "Skeppare B - fartygsbefäl klass VIII 1980-2000.xlsx"
)

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

BAD_NAME_WORDS = [
    # School / institution text
    "SJÖBEFÄLSSKOLAN",
    "SJÖBEFALSSKOLAN",
    "SJÖBETÄLSSKOLAN",
    "SJÖBOFÄLSSKOLAN",
    "SJÖBÖFÄLSSKOLAN",
    "SJÖLETALSSKULAN",
    "SJÖLETALSSKOLAN",
    "SJÖBAEFÄLSSKOLAN",
    "CHALMERS",
    "TEKNISKA",
    "HÖGSKOLA",
    "HOGSKOLA",
    "NAUTISKA",
    "INSTITUTIONEN",
    "EFTERNAMN",
    "EFTERNAM",
    "EFTERMNAMN",
    "EFTER RAMN",
    "EFTERAMN",
    "EFTERHMAN",
    "EFTERNAMN TILLTALSNAMN",
    "EFTERNAMN TILLTALSNHAMN",
    "TILLTALSNAMN",
    "TILLTALSNHAMN",
    "TILLTALSNARNN",
    "TILLTALSNAM",
    "NAMN",
    "NAMNET",

    "SJÖBEFÖÄLSSKOLAN",
    "SJÖBEFÄLSSKOLTAN",
    "SJÖBEFÄLSSKOLTAN I",
    "SIJÖBEFÄLSSKOLAN",
    "SIJÖBEFALSSKOLAN",
    "SJÖBEFÄLSSKOLTAN",

    # Certificate headings
    "SKEPPAREXAMEN",
    "BETYG",
    "AVLAGD",
    "DATUM",
    "BEVIS",
    "INTYG",

    # Form labels
    "PERSONNR",
    "PERSONNUMMER",
    "PERSNR",
    "POST NR",
    "POSTADRESS",
    "ADRESS",
    "NAMNTECKNING",
    "EXAMINANDENS",
    "EXAMINATOR",
    "EXAMINA",
    "UNDERSKRIFT",
    "LÄRARENS",
    "LARARENS",
    "REKTOR",
    "STUDIEREKTOR",
    "EXAMENSFÖRRÄTTARE",
    "EXAMENSFORRATTARE",
    "FÖRRÄTTARE",
    "FORRATTARE",

    # Roles / titles
    "ADJUNKT",
    "LEKTOR",
    "TIMLÄRARE",
    "TIMLARARE",
    "SJÖKAPTEN",
    "SJOKAPTEN",
    "LÄRARE",
    "LARARE",

    # Instruction text
    "SKRIV",
    "SAMMA",
    "NUMMER",
    "UNDERLIGGANDE",
    "PAPPER",
    "OMRÅDE",
    "OMRADE",
    "PROV",
    "PRÖVNING",
    "PROVNING",
]

def is_bad_name_candidate(candidate):

    if not candidate:
        return True

    upper = str(candidate).upper().strip()

    if is_blacklisted_name(upper):
        return True

    normalized = normalize_name_for_matching(upper)

    if not normalized:
        return True

    if any(word in upper for word in BAD_NAME_WORDS):
        return True

    if any(
            normalize_name_for_matching(word) in normalized
            for word in BAD_NAME_WORDS
    ):
        return True

    if looks_like_bad_name_word(upper):
        return True

    words = normalized.split()

    # Need at least surname + first name
    if len(words) < 2:
        return True

    # Long strings are usually OCR garbage or whole form lines
    if len(words) > 5:
        return True

    # Reject one-letter-heavy garbage like "N", "RR", "EE", "JE"
    short_words = [
        w for w in words
        if len(w) <= 2
    ]

    if len(short_words) >= 2:
        return True

    letters = re.sub(
        r"[^A-ZÅÄÖ]",
        "",
        upper
    )

    if len(letters) < 5:
        return True

    # Too much repeated OCR-noise text
    noise_tokens = {
        "RR", "RDR", "RE", "EE", "JE", "VR", "TR", "FR", "AR",
        "SAS", "PTS", "SKR", "RNUN", "NNPOST", "FORS"
    }

    if any(token in words for token in noise_tokens):
        return True

    return False

def looks_like_bad_name_word(candidate):
    """
    Fuzzy check for OCR-distorted form labels and institution text.
    """

    if not candidate:
        return False

    normalized_candidate = normalize_name_for_matching(candidate)

    if not normalized_candidate:
        return False

    for bad_word in BAD_NAME_WORDS:

        normalized_bad = normalize_name_for_matching(bad_word)

        if not normalized_bad:
            continue

        if normalized_bad in normalized_candidate:
            return True

        score = fuzz.partial_ratio(
            normalized_bad,
            normalized_candidate
        )

        if score >= 88:
            return True

    return False

def clean_register_bevisnummer(value):
    """
    Normaliserar intygsnummer/bevisnummer från registret.
    Hindrar t.ex. 17853.0 från att bli felaktigt.
    """

    if value is None:
        return ""

    text = str(value).strip()

    if not text or text.lower() == "nan":
        return ""

    try:
        return str(int(float(text)))
    except Exception:
        return re.sub(r"[^0-9]", "", text)


def clean_register_personnummer(value):
    """
    Normaliserar personnummer från registret.
    """

    if value is None:
        return ""

    text = str(value).strip()

    if not text or text.lower() == "nan":
        return ""

    text = text.replace(" ", "")

    return text


def normalize_name_for_matching(text):
    """
    OCR-tolerant namnnormalisering för fuzzy matchning mot registret.

    Exempel:
    KINDSTRÖM -> KINDSTROM
    K1NDSTR0M -> KINDSTROM
    """

    if not text:
        return ""

    text = str(text).upper()

    text = text.replace("Å", "A")
    text = text.replace("Ä", "A")
    text = text.replace("Ö", "O")

    # OCR-fel i namn
    text = text.replace("0", "O")
    text = text.replace("1", "I")
    text = text.replace("|", "I")
    text = text.replace("5", "S")

    text = re.sub(
        r"[^A-Z\- ]",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def load_register():
    """
    Läser det auktoritativa registret.

    Returnerar:
    - register_by_bevisnummer
    - register_by_personnummer
    - register_name_records
    """

    if not REGISTER_FILE.exists():
        raise FileNotFoundError(
            f"Register file not found: {REGISTER_FILE}"
        )

    df = pd.read_excel(
        REGISTER_FILE,
        dtype=str
    )

    required_columns = [
        "Efternamn",
        "Förnamn",
        "Personnummer",
        "Intygsnummer",
    ]

    missing_columns = [
        col
        for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            f"Register is missing columns: {missing_columns}. "
            f"Found columns: {list(df.columns)}"
        )

    register_by_bevisnummer = {}
    register_by_personnummer = {}
    register_name_records = []

    for _, row in df.iterrows():

        bevisnummer = clean_register_bevisnummer(
            row["Intygsnummer"]
        )

        if not bevisnummer:
            continue

        efternamn = str(
            row["Efternamn"]
        ).strip()

        fornamn = str(
            row["Förnamn"]
        ).strip()

        full_name = (
            f"{efternamn} {fornamn}"
            .strip()
            .upper()
        )

        personnummer = clean_register_personnummer(
            row["Personnummer"]
        )

        record = {
            "bevisnummer": bevisnummer,
            "name": full_name,
            "normalized_name": normalize_name_for_matching(
                full_name
            ),
            "personnummer": personnummer,
        }

        register_by_bevisnummer[bevisnummer] = record

        if personnummer:
            register_by_personnummer[personnummer] = record

        register_name_records.append(record)

    print(
        f"Loaded {len(register_by_bevisnummer)} register records by bevisnummer"
    )

    print(
        f"Loaded {len(register_by_personnummer)} register records by personnummer"
    )

    print(
        f"Loaded {len(register_name_records)} register name records"
    )

    return (
        register_by_bevisnummer,
        register_by_personnummer,
        register_name_records
    )


def match_ocr_names_to_register(
        ocr_names,
        register_name_records,
        min_score=92
):
    """
    Försöker matcha OCR-namn mot registret.

    Används bara när bevisnummer saknas eller inte finns i registret.

    Returnerar:
        (record, score)
    eller:
        (None, 0)
    """

    if not ocr_names:
        return None, 0

    best_record = None
    best_score = 0

    normalized_ocr_names = []

    for name in ocr_names:
        normalized = normalize_name_for_matching(name)

        if normalized:
            normalized_ocr_names.append(normalized)

    if not normalized_ocr_names:
        return None, 0

    for ocr_name in normalized_ocr_names:

        for record in register_name_records:

            score = fuzz.token_sort_ratio(
                ocr_name,
                record["normalized_name"]
            )

            if score > best_score:
                best_score = score
                best_record = record

    if best_score >= min_score:
        return best_record, best_score

    return None, best_score

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
    Extracts personnummer candidates from lines near PERSONNUMMER/PERSONNR labels.

    Supports:
    - 650224-4897
    - 650224 4897
    - 6502244897
    - 65-02-24-4897
    - 65 02 24 4897
    """

    if not text:
        return []

    candidates = []

    snippets = find_personnummer_snippets(
        text,
        window_lines=4
    )

    # Fallback if no snippets were found
    if not snippets:
        snippets = text.splitlines()

    for snippet in snippets:

        normalized = normalize_ocr_digit_text(snippet)

        # Direct pattern with optional separators
        for match in re.finditer(
                r"\b(\d{2})[- ]?(\d{2})[- ]?(\d{2})[- ]?(\d{4})\b",
                normalized
        ):

            date_digits = (
                match.group(1)
                + match.group(2)
                + match.group(3)
            )

            suffix = match.group(4)

            if valid_date_yymmdd(date_digits):
                candidates.append(
                    f"{date_digits}-{suffix}"
                )

        digits = re.sub(
            r"[^0-9]",
            "",
            normalized
        )

        for pnr_candidate in extract_10_digit_windows(digits):
            candidates.append(pnr_candidate)

        # Common OCR case: extra leading zero
        if len(digits) == 11 and digits.startswith("0"):

            d = digits[1:]

            for pnr_candidate in extract_10_digit_windows(d):
                candidates.append(pnr_candidate)

    return list(dict.fromkeys(candidates))
# --------------------------------------------------
# BEVISNUMMER
# --------------------------------------------------

def extract_bevisnummer_candidates(text):
    """
    Extraherar bevisnummer/intygsnummer från OCR-text.

    Fångar:
    - Nr 249938
    - NR 249938
    - Bevisnr 17798
    - Bevis nr. 18015
    - Intygsnr 249938

    Undviker:
    - SÖ 85-021
    - 85-021
    - personnummer
    """

    if not text:
        return []

    t = normalize_ocr_text(text)

    candidates = []

    patterns = [
        r"\bBEVIS\s*NR\s*\.?\s*(\d{4,6})\b",
        r"\bBEVISNR\s*\.?\s*(\d{4,6})\b",
        r"\bINTYG\s*NR\s*\.?\s*(\d{4,6})\b",
        r"\bINTYGSNR\s*\.?\s*(\d{4,6})\b",
        r"\bNR\s*\.?\s*(\d{4,6})\b",
    ]

    for pattern in patterns:
        matches = re.findall(
            pattern,
            t,
            flags=re.IGNORECASE
        )

        candidates.extend(matches)

    cleaned = []

    for value in candidates:

        digits = re.sub(
            r"[^0-9]",
            "",
            str(value)
        )

        if not digits:
            continue

        # Bevisnummer/intygsnummer kan vara 4-6 siffror.
        if not (4 <= len(digits) <= 6):
            continue

        number = int(digits)

        if 1000 <= number <= 999999:
            cleaned.append(digits)

    return list(dict.fromkeys(cleaned))

def get_volume_number_range(volume_hint):
    """
    Extracts a numeric range from volume names like:
    F2AB1 17818-18095
    F2AB14 24598-25099

    Returns:
        (low, high)
    or:
        (None, None)
    """

    numbers = re.findall(
        r"\d+",
        str(volume_hint)
    )

    if len(numbers) < 2:
        return None, None

    low = int(numbers[-2])
    high = int(numbers[-1])

    return low, high

def filter_register_records_by_volume(
        register_name_records,
        volume_hint
):
    """
    Restricts register name fallback to the numeric bevisnummer range
    encoded in the volume folder name.

    Example:
    F2AB1 17818-18095 -> only records with bevisnummer 17818-18095.
    """

    low, high = get_volume_number_range(volume_hint)

    if low is None or high is None:
        return register_name_records

    filtered = []

    for record in register_name_records:

        bevisnummer = str(
            record.get("bevisnummer", "")
        ).strip()

        if not bevisnummer.isdigit():
            continue

        number = int(bevisnummer)

        if low <= number <= high:
            filtered.append(record)

    if filtered:
        return filtered

    return register_name_records

def find_similar_register_bevisnummer(
        raw_bevisnummer,
        register_by_bevisnummer,
        volume_hint="",
        max_distance=1
):
    """
    Searches the authoritative register for bevisnummer similar to an OCR candidate.

    This does NOT guess a new number.
    It only returns a corrected number if exactly one nearby number exists
    in the authoritative register.

    Returns:
        resolved_bevisnummer, candidates, reason

    Examples:
        raw 179873 may match register key 17873
        raw 249138 may produce several candidates and should stay unresolved
    """

    if not raw_bevisnummer:
        return "", [], "NO_RAW_BEVISNUMMER"

    digits = re.sub(
        r"[^0-9]",
        "",
        str(raw_bevisnummer)
    )

    if not digits:
        return "", [], "NO_DIGITS"

    # Exact register hit is always accepted.
    if digits in register_by_bevisnummer:
        return digits, [digits], "EXACT_REGISTER_MATCH"

    # Only attempt similarity search for suspicious OCR values.
    # This avoids overcorrecting normal 4-5 digit numbers.
    if len(digits) < 5:
        return "", [], "TOO_SHORT_FOR_SIMILAR_SEARCH"

    low, high = get_volume_number_range(volume_hint)

    register_keys = list(
        register_by_bevisnummer.keys()
    )

    # Prefer searching in the same volume range if we can parse it.
    if low is not None and high is not None:

        register_keys = [
            key
            for key in register_keys
            if key.isdigit()
            and low <= int(key) <= high
        ]

    candidates = []

    for key in register_keys:

        if not key.isdigit():
            continue

        distance = Levenshtein.distance(
            digits,
            key
        )

        if distance <= max_distance:

            ratio = fuzz.ratio(
                digits,
                key
            )

            candidates.append(
                {
                    "key": key,
                    "distance": distance,
                    "ratio": ratio,
                }
            )

    if not candidates:
        return "", [], "NO_SIMILAR_REGISTER_MATCH"

    candidates = sorted(
        candidates,
        key=lambda item: (
            item["distance"],
            -item["ratio"],
            item["key"]
        )
    )

    best_distance = candidates[0]["distance"]

    best_candidates = [
        item
        for item in candidates
        if item["distance"] == best_distance
    ]

    # Critical safety rule:
    # only auto-correct if exactly one best candidate exists.
    if len(best_candidates) == 1:
        return (
            best_candidates[0]["key"],
            candidates,
            "UNIQUE_SIMILAR_REGISTER_MATCH"
        )

    return (
        "",
        candidates,
        "AMBIGUOUS_SIMILAR_REGISTER_MATCH"
    )

def resolve_ocr_bevisnummer_against_register(
        raw_candidates,
        register_by_bevisnummer,
        volume_hint,
        validation
):
    """
    Selects OCR bevisnummer, then tries to resolve suspicious values
    against the authoritative register.

    Returns:
        ocr_bevisnummer, resolved_candidate
    """

    if not raw_candidates:
        return "", ""

    counter = Counter(raw_candidates)

    raw_bevisnummer = counter.most_common(1)[0][0]

    resolved, similar_candidates, reason = find_similar_register_bevisnummer(
        raw_bevisnummer,
        register_by_bevisnummer,
        volume_hint=volume_hint,
        max_distance=1
    )

    if reason == "EXACT_REGISTER_MATCH":

        return raw_bevisnummer, resolved

    if reason == "UNIQUE_SIMILAR_REGISTER_MATCH":

        validation.append(
            f"Resolved OCR bevisnummer by register similarity: "
            f"raw={raw_bevisnummer}, resolved={resolved}, "
            f"reason={reason}"
        )

        return raw_bevisnummer, resolved

    if reason == "AMBIGUOUS_SIMILAR_REGISTER_MATCH":

        validation.append(
            f"Ambiguous OCR bevisnummer similarity search: "
            f"raw={raw_bevisnummer}, "
            f"candidates={similar_candidates}"
        )

        # Do not guess. Return raw OCR candidate and let PNR/name fallback handle it.
        return raw_bevisnummer, raw_bevisnummer

    if reason == "NO_SIMILAR_REGISTER_MATCH":

        validation.append(
            f"No similar register bevisnummer found for OCR candidate: "
            f"raw={raw_bevisnummer}"
        )

    return raw_bevisnummer, raw_bevisnummer
# ---------
# FÖR OCH EFTERNAMN
# -------------

def extract_name_candidates(text):

    if not text:
        return []

    candidates = []
    priority_candidates = []

    upper = text.upper()

    #
    # Priority 1:
    # Structured fields: EFTERNAMN / TILLTALSNAMN
    #

    pattern = re.compile(
        r"(?:EFTERNAMN|FTERNAMN|EFTENAMN|EFTERNAM)\s*([A-ZÅÄÖ\-]+)"
        r".{0,100}?"
        r"(?:TILLTALSNAMN|TILLTALSNAM|TILLTALSNARNN)\s*([A-ZÅÄÖ ]+?)"
        r"(?:PERSONNR|PERSONNUMMER|PARSONNR|PERSNR|$)",
        re.DOTALL
    )

    for match in pattern.finditer(upper):

        surname = match.group(1).strip()
        firstname = match.group(2).strip()

        candidate = f"{surname} {firstname}"

        candidate = clean_name_candidate(candidate)

        if not candidate:
            continue

        if not is_bad_name_candidate(candidate):
            priority_candidates.append(candidate)

    if priority_candidates:
        return list(dict.fromkeys(priority_candidates))

    #
    # Priority 2:
    # Fallback around PERSONNR / PERSONNUMMER
    #

    lines = text.splitlines()

    markers = [
        "PERSONNR",
        "PERSONNUMMER",
        "PARSONNR",
        "PARSONNUMMER",
        "PERSNR",
    ]

    for i, line in enumerate(lines):

        upper_line = line.upper()

        if not any(m in upper_line for m in markers):
            continue

        window = lines[
            max(0, i - 5): min(len(lines), i + 5)
        ]

        for candidate in window:

            candidate = candidate.strip()

            if len(candidate) < 5:
                continue

            # Remove personnummer from candidate line
            candidate = re.sub(
                r"\d{6}[- ]?\d{4}",
                "",
                candidate
            )

            candidate = clean_name_candidate(candidate)

            if not candidate:
                continue

            words = candidate.split()

            if len(words) < 2:
                continue

            if is_bad_name_candidate(candidate):
                continue

            # All words must be name-like
            if all(
                re.fullmatch(
                    r"[A-Za-zÅÄÖåäö\-]+",
                    w
                )
                for w in words
            ):
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

def validate_bevisnummer_against_ocr(
        bevisnummer,
        ocr_name,
        ocr_pnr,
        register_by_bevisnummer,
        register_by_personnummer=None,
        min_name_score=85
):
    """
    Kontrollerar om ett OCR-bevisnummer verkar stämma med OCR-namn/personnummer.

    Princip:
    - Bevisnummer som finns i registret är stark evidens.
    - Dåligt OCR-personnummer ska inte automatiskt avvisa bevisnumret.
    - Dåligt OCR-namn ska ignoreras.
    - Bevisnummer avvisas främst om OCR-personnummer tydligt pekar på en annan registerpost.
    """

    if not bevisnummer:
        return None, False, "NO_BEVISNUMMER", 0

    if bevisnummer not in register_by_bevisnummer:
        return None, False, "BEVISNUMMER_NOT_IN_REGISTER", 0

    record = register_by_bevisnummer[bevisnummer]

    register_pnr = record.get("personnummer", "")
    register_name = record.get("name", "")

    usable_ocr_name = ""

    if ocr_name and not is_bad_name_candidate(ocr_name):
        usable_ocr_name = ocr_name

    # 1. Strongest case: OCR personnummer matches this register record.
    if ocr_pnr and register_pnr and ocr_pnr == register_pnr:
        return record, True, "BEVISNUMMER_VALIDATED_BY_PNR", 100

    # 2. OCR personnummer differs from this bevisnummer's register record.
    if ocr_pnr and register_pnr and ocr_pnr != register_pnr:

        other_record = None

        if register_by_personnummer:
            other_record = register_by_personnummer.get(ocr_pnr)

        if (
            other_record
            and other_record.get("bevisnummer") != bevisnummer
        ):
            return (
                record,
                False,
                "BEVISNUMMER_REJECTED_PNR_POINTS_TO_OTHER_RECORD",
                0
            )

        if usable_ocr_name:

            score = fuzz.token_sort_ratio(
                normalize_name_for_matching(usable_ocr_name),
                normalize_name_for_matching(register_name)
            )

            if score >= min_name_score:
                return (
                    record,
                    True,
                    "BEVISNUMMER_VALIDATED_BY_NAME_PNR_WARNING",
                    score
                )

            return (
                record,
                False,
                "BEVISNUMMER_REJECTED_NAME_AND_PNR_MISMATCH",
                score
            )

        return (
            record,
            True,
            "BEVISNUMMER_ACCEPTED_PNR_WARNING",
            0
        )

    # 3. No reliable OCR personnummer. Use name if usable.
    if usable_ocr_name and register_name:

        score = fuzz.token_sort_ratio(
            normalize_name_for_matching(usable_ocr_name),
            normalize_name_for_matching(register_name)
        )

        if score >= min_name_score:
            return record, True, "BEVISNUMMER_VALIDATED_BY_NAME", score

        return (
            record,
            False,
            "BEVISNUMMER_REJECTED_NAME_MISMATCH",
            score
        )

    # 4. Bevisnummer exists in register, but no usable secondary evidence.
    return (
        record,
        True,
        "BEVISNUMMER_ACCEPTED_NO_USABLE_SECONDARY",
        0
    )

def generate_pnr_variants(pnr):

    if not pnr:
        return []

    digits = re.sub(
        r"[^0-9]",
        "",
        pnr
    )

    if len(digits) != 10:
        return []

    variants = []

    exact = f"{digits[:6]}-{digits[6:]}"
    variants.append(exact)

    first_digit_alternatives = {
        "0": ["6"],
        "1": ["7"],
        "2": ["5"],
        "3": ["8"],
        "5": ["2"],
        "6": ["0"],
        "7": ["1"],
        "8": ["3"],
        "9": ["5"],
    }

    first = digits[0]

    for alt in first_digit_alternatives.get(first, []):

        alt_digits = alt + digits[1:]

        variant = f"{alt_digits[:6]}-{alt_digits[6:]}"

        if variant not in variants:
            variants.append(variant)

    return variants

def clean_name_candidate(candidate):

    if not candidate:
        return ""

    text = str(candidate).strip()

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    # Remove leading OCR artifacts like "N "
    text = re.sub(
        r"^(N|M|RN|NR)\s+",
        "",
        text,
        flags=re.IGNORECASE
    )

    # Remove trailing labels accidentally attached
    text = re.sub(
        r"\b(PERSONNR|PERSONNUMMER|PERSNR|POSTADRESS|POST NR)\b.*$",
        "",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    ).strip()

    return text
#--------------------------------------------------
# PROCESS ONE PDF
# --------------------------------------------------

def match_pnr_to_register(
        ocr_pnr,
        register_by_personnummer
):
    """
    Fallback på personnummer.

    Först exakt match.
    Sedan vanliga OCR-varianter på första siffran i YYMMDD.
    """

    if not ocr_pnr:
        return None

    for candidate in generate_pnr_variants(ocr_pnr):

        record = register_by_personnummer.get(candidate)

        if record:
            return record

    return None

def select_best_pnr_candidate(
        all_pnrs,
        register_by_personnummer
):
    """
    Selects the strongest personnummer candidate.

    Priority:
    1. Exact register match
    2. OCR-variant register match
    3. Most common OCR candidate
    """

    if not all_pnrs:
        return ""

    clean_candidates = []

    for pnr in all_pnrs:

        if not pnr:
            continue

        digits = re.sub(
            r"[^0-9]",
            "",
            str(pnr)
        )

        if len(digits) != 10:
            continue

        formatted = f"{digits[:6]}-{digits[6:]}"

        if formatted not in clean_candidates:
            clean_candidates.append(formatted)

    if not clean_candidates:
        return ""

    # 1. Exact register match first
    for candidate in clean_candidates:

        if candidate in register_by_personnummer:
            return candidate

    # 2. OCR-variant register match
    for candidate in clean_candidates:

        record = match_pnr_to_register(
            candidate,
            register_by_personnummer
        )

        if record:
            return record.get(
                "personnummer",
                candidate
            )

    # 3. Most common OCR candidate
    counter = Counter(clean_candidates)

    return counter.most_common(1)[0][0]

def process_pdf(args):
    (
        pdf_path,
        input_root,
        output_root,
        config,
        register_by_bevisnummer,
        register_by_personnummer,
        register_name_records,
    ) = args

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

        match_method = "UNMATCHED"

        register_name = ""
        register_personnummer = ""
        name_match_score = 0
        ocr_bevisnummer = ""

        #
        # Bevisnummer söks endast på de två första sidorna
        # eftersom de nästan alltid finns där
        #

        for p in range(
                start,
                min(start + 2, end)
        ):

            try:
                text = reader.pages[p].extract_text() or ""

            except Exception:
                continue

            all_bevisnummer.extend(
                extract_bevisnummer_candidates(text)
            )

        for p in range(start, end):

            try:
                text = reader.pages[p].extract_text() or ""

            except Exception:
                continue

            page_pnrs = extract_personnummer_candidates(text)

            all_pnrs.extend(page_pnrs)

            page_names = extract_name_candidates(text)

            all_names.extend(page_names)

        pnr = select_best_pnr_candidate(
            all_pnrs,
            register_by_personnummer
        )

        if pnr:
            validation.append(
                f"Selected best PNR {pnr} from candidates {sorted(set(all_pnrs))}"
            )

        ocr_name = ""

        cleaned_names = []

        for name in all_names:

            cleaned = clean_name_candidate(name)

            if not cleaned:
                continue

            if is_bad_name_candidate(cleaned):
                continue

            cleaned_names.append(cleaned)

        filtered_names = list(
            dict.fromkeys(cleaned_names)
        )

        if filtered_names:
            name_counter = Counter(filtered_names)

            ocr_name = (
                name_counter.most_common(1)[0][0]
            )

        ocr_bevisnummer = ""
        bevisnummer = ""

        if all_bevisnummer:
            (
                ocr_bevisnummer,
                bevisnummer
            ) = resolve_ocr_bevisnummer_against_register(
                raw_candidates=all_bevisnummer,
                register_by_bevisnummer=register_by_bevisnummer,
                volume_hint=volym,
                validation=validation
            )

        #
        # PRIORITET 1: Bevisnummer mot register
        #

        #
        # PRIORITET 1:
        # OCR-bevisnummer, men bara om det stämmer med namn/personnummer i registret
        #

        reg_record = None
        name_match_score = 0

        if bevisnummer:

            (
                candidate_record,
                accepted,
                reason,
                score
            ) = validate_bevisnummer_against_ocr(
                bevisnummer=bevisnummer,
                ocr_name=ocr_name,
                ocr_pnr=pnr,
                register_by_bevisnummer=register_by_bevisnummer,
                register_by_personnummer=register_by_personnummer,
                min_name_score=85
            )

            name_match_score = score

            if accepted and candidate_record:

                reg_record = candidate_record

                match_method = reason

                register_name = reg_record["name"]
                register_personnummer = reg_record["personnummer"]

                validation.append(
                    f"Accepted OCR bevisnummer {bevisnummer}: "
                    f"{register_name}, reason={reason}"
                )


            else:
                validation.append(
                    f"Rejected OCR bevisnummer {bevisnummer}: "
                    f"reason={reason}, "
                    f"ocr_name={ocr_name}, "
                    f"ocr_pnr={pnr}"
                )

                strong_rejection_reasons = {
                    "BEVISNUMMER_NOT_IN_REGISTER",
                    "BEVISNUMMER_REJECTED_PNR_POINTS_TO_OTHER_RECORD",
                    "BEVISNUMMER_REJECTED_NAME_AND_PNR_MISMATCH",
                    "BEVISNUMMER_REJECTED_NAME_MISMATCH",
                }

                if reason in strong_rejection_reasons:
                    bevisnummer = ""

                else:
                    reg_record = candidate_record

                    if reg_record:
                        match_method = reason
                        register_name = reg_record["name"]
                        register_personnummer = reg_record["personnummer"]

        #
        # PRIORITET 2:
        # Exakt personnummer mot register
        #

        if not reg_record and pnr:

            pnr_record = match_pnr_to_register(
                pnr,
                register_by_personnummer
            )

            if pnr_record:
                reg_record = pnr_record

                match_method = "PNR_FALLBACK"

                bevisnummer = reg_record["bevisnummer"]
                register_name = reg_record["name"]
                register_personnummer = reg_record["personnummer"]

                validation.append(
                    f"PNR fallback matched {pnr}: "
                    f"{register_name}, bevisnummer={bevisnummer}"
                )

        #
        # PRIORITET 3:
        # Namn-fallback mot register
        #

        if not reg_record:

            volume_register_name_records = filter_register_records_by_volume(
                register_name_records,
                volym
            )

            reg_record, score = match_ocr_names_to_register(
                filtered_names,
                volume_register_name_records,
                min_score=88
            )

            name_match_score = score

            if reg_record:
                match_method = "NAME_FALLBACK"

                bevisnummer = reg_record["bevisnummer"]
                register_name = reg_record["name"]
                register_personnummer = reg_record["personnummer"]

                validation.append(
                    f"Name fallback matched OCR names {filtered_names} "
                    f"to {register_name} "
                    f"with score {score}. "
                    f"Resolved bevisnummer={bevisnummer}"
                )

        #
        # PRIORITET 4:
        # Ingen träff
        #

        if not reg_record:
            match_method = "UNMATCHED"

            validation.append(
                f"No register match. "
                f"ocr_bevisnummer={ocr_bevisnummer}, "
                f"ocr_name={ocr_name}, "
                f"ocr_pnr={pnr}, "
                f"best_name_score={name_match_score}"
            )


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

                ocr_bevisnummer,
                bevisnummer,

                register_name,

                pnr,
                register_personnummer,
            ]
        )

    return rows, validation


def safe_int(value):

    text = str(value).strip()

    if not text.isdigit():
        return None

    return int(text)


def sort_rows_for_sequence_fallback(all_rows):
    """
    Sort rows before sequence fallback.

    Sort key:
    - volym
    - file name
    - start_page
    """

    FILE_COL = 0
    VOLUME_COL = 1
    START_PAGE_COL = 13

    return sorted(
        all_rows,
        key=lambda row: (
            str(row[VOLUME_COL]),
            str(row[FILE_COL]),
            safe_int(row[START_PAGE_COL]) or 0
        )
    )

def sort_rows_for_sequence_fallback(all_rows):
    """
    Sort rows before sequence fallback.

    Sort key:
    - volym
    - file name
    - start_page
    """

    FILE_COL = 0
    VOLUME_COL = 1
    START_PAGE_COL = 13

    return sorted(
        all_rows,
        key=lambda row: (
            str(row[VOLUME_COL]),
            str(row[FILE_COL]),
            safe_int(row[START_PAGE_COL]) or 0
        )
    )

def safe_int(value):

    text = str(value).strip()

    if not text.isdigit():
        return None

    return int(text)

def apply_sequence_fallback(all_rows):
    """
    Conservative sequence repair after sorting.

    Only fills resolved_bevisnummer when:
    - current row is UNMATCHED
    - previous and next rows are in the same volume
    - previous and next resolved bevisnummer imply exactly one missing number
    - OCR bevisnummer is empty or similar to expected number
    """

    VOLUME_COL = 1
    OCR_BEVIS_COL = 2
    RESOLVED_BEVIS_COL = 3
    MATCH_METHOD_COL = 4

    sorted_rows = sort_rows_for_sequence_fallback(all_rows)

    for i in range(1, len(sorted_rows) - 1):

        row = sorted_rows[i]
        prev_row = sorted_rows[i - 1]
        next_row = sorted_rows[i + 1]

        current_volume = row[VOLUME_COL]
        previous_volume = prev_row[VOLUME_COL]
        next_volume = next_row[VOLUME_COL]

        same_volume = (
            current_volume == previous_volume
            and current_volume == next_volume
        )

        if not same_volume:
            continue

        if row[MATCH_METHOD_COL] != "UNMATCHED":
            continue

        prev_bnr = safe_int(
            prev_row[RESOLVED_BEVIS_COL]
        )

        next_bnr = safe_int(
            next_row[RESOLVED_BEVIS_COL]
        )

        if prev_bnr is None:
            continue

        if next_bnr is None:
            continue

        expected_bnr = prev_bnr + 1

        if next_bnr != expected_bnr + 1:
            continue

        ocr_bnr = str(
            row[OCR_BEVIS_COL]
        ).strip()

        if ocr_bnr:

            cleaned_ocr_bnr = re.sub(
                r"[^0-9]",
                "",
                ocr_bnr
            )

            if cleaned_ocr_bnr:

                distance = Levenshtein.distance(
                    cleaned_ocr_bnr,
                    str(expected_bnr)
                )

                if distance > 2:
                    continue

        row[RESOLVED_BEVIS_COL] = str(expected_bnr)
        row[MATCH_METHOD_COL] = "SEQUENCE_FALLBACK_REVIEW"

    return sorted_rows
# --------------------------------------------------
# MAIN
# --------------------------------------------------

def process_all(
    input_root,
    output_root,
    config,
    register_by_bevisnummer,
    register_by_personnummer,
    register_name_records,
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
                    config,
                    register_by_bevisnummer,
                    register_by_personnummer,
                    register_name_records,
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

    all_rows = apply_sequence_fallback(all_rows)

    # Excel
    wb = Workbook()
    ws = wb.active

    ws.append(
        [
            "file name",
            "volym",

            "ocr_bevisnummer",
            "resolved_bevisnummer",

            "register_name",

            "ocr_personnummer",
            "register_personnummer",
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

    (
        register_by_bevisnummer,
        register_by_personnummer,
        register_name_records
    ) = load_register()

    process_all(
        sys.argv[1],
        sys.argv[2],
        config,
        register_by_bevisnummer,
        register_by_personnummer,
        register_name_records
    )