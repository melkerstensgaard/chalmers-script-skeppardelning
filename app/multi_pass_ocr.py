import re
from collections import Counter
from dataclasses import dataclass

import pymupdf
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter, ImageOps


PERSONNUMMER_PATTERN = re.compile(
    r"(?<!\d)((?:18|19|20)?\d{6})[-+ ]?(\d{4})(?!\d)"
)


@dataclass
class OCRPassResult:
    name: str
    text: str
    candidates: list[str]


@dataclass
class MultiPassOCRResult:
    text: str
    personnummer: str
    score: int
    reasons: list[str]
    vote_count: int
    total_sources: int
    checksum_valid: bool
    candidates: dict[str, int]


def normalize_personnummer(date_part: str, suffix: str) -> str:
    digits = re.sub(r"\D", "", date_part + suffix)

    if len(digits) == 12:
        digits = digits[2:]

    if len(digits) != 10:
        return ""

    return f"{digits[:6]}-{digits[6:]}"


def extract_personnummer_candidates(text: str) -> list[str]:
    candidates = set()

    for match in PERSONNUMMER_PATTERN.finditer(text or ""):
        normalized = normalize_personnummer(
            match.group(1),
            match.group(2),
        )

        if normalized:
            candidates.add(normalized)

    return sorted(candidates)


def valid_luhn(personnummer: str) -> bool:
    digits = re.sub(r"\D", "", personnummer)

    if len(digits) == 12:
        digits = digits[2:]

    if len(digits) != 10:
        return False

    total = 0

    for index, character in enumerate(digits):
        value = int(character)

        if index % 2 == 0:
            value *= 2
            if value > 9:
                value -= 9

        total += value

    return total % 10 == 0


def page_to_image(page, dpi: int = 300) -> Image.Image:
    scale = dpi / 72.0

    pixmap = page.get_pixmap(
        matrix=pymupdf.Matrix(scale, scale),
        alpha=False,
    )

    return Image.frombytes(
        "RGB",
        (pixmap.width, pixmap.height),
        pixmap.samples,
    )


def create_image_variants(image: Image.Image) -> dict[str, Image.Image]:
    grayscale = ImageOps.grayscale(image)

    contrast = ImageEnhance.Contrast(grayscale).enhance(1.8)
    contrast = contrast.filter(ImageFilter.SHARPEN)

    threshold = grayscale.point(
        lambda pixel: 255 if pixel > 175 else 0
    )

    return {
        "original": grayscale,
        "contrast": contrast,
        "threshold": threshold,
    }


def run_tesseract_pass(
    image: Image.Image,
    name: str,
    psm: int,
    language: str,
) -> OCRPassResult:
    config = (
        f"--oem 3 --psm {psm} "
        "-c preserve_interword_spaces=1"
    )

    try:
        text = pytesseract.image_to_string(
            image,
            lang=language,
            config=config,
        )
    except pytesseract.TesseractError:
        if language == "eng":
            raise

        text = pytesseract.image_to_string(
            image,
            lang="eng",
            config=config,
        )

    return OCRPassResult(
        name=name,
        text=text,
        candidates=extract_personnummer_candidates(text),
    )


def choose_display_text(
    pass_results: list[OCRPassResult],
    selected_candidate: str,
    embedded_text: str,
) -> str:
    supporting_results = [
        result
        for result in pass_results
        if selected_candidate
        and selected_candidate in result.candidates
        and result.text.strip()
    ]

    if supporting_results:
        return max(
            supporting_results,
            key=lambda result: len(result.text.strip()),
        ).text

    available_texts = [
        result.text
        for result in pass_results
        if result.text.strip()
    ]

    if embedded_text.strip():
        available_texts.append(embedded_text)

    if not available_texts:
        return ""

    return max(
        available_texts,
        key=lambda value: len(value.strip()),
    )


def analyze_page_multi_pass(
    page,
    language: str = "swe+eng",
    dpi: int = 300,
) -> MultiPassOCRResult:
    embedded_text = page.get_text("text") or ""
    image = page_to_image(page, dpi=dpi)
    variants = create_image_variants(image)

    pass_definitions = [
        ("Original, PSM 3", "original", 3),
        ("Original, PSM 6", "original", 6),
        ("Original, PSM 11", "original", 11),
        ("Kontrast, PSM 6", "contrast", 6),
        ("Trosklad, PSM 6", "threshold", 6),
        ("Trosklad, PSM 11", "threshold", 11),
    ]

    pass_results = []
    failed_passes = []

    for pass_name, variant_name, psm in pass_definitions:
        try:
            result = run_tesseract_pass(
                image=variants[variant_name],
                name=pass_name,
                psm=psm,
                language=language,
            )
            pass_results.append(result)
        except Exception as exception:
            failed_passes.append(
                f"{pass_name}: {type(exception).__name__}: {exception}"
            )

    candidate_votes = Counter()

    for result in pass_results:
        for found_candidate in set(result.candidates):
            candidate_votes[found_candidate] += 1

    embedded_candidates = extract_personnummer_candidates(embedded_text)
    embedded_source_used = bool(embedded_text.strip())

    if embedded_source_used:
        for found_candidate in set(embedded_candidates):
            candidate_votes[found_candidate] += 1

    total_sources = len(pass_results) + int(embedded_source_used)

    if not candidate_votes:
        reasons = [
            "Inget personnummer hittades i de slutforda OCR-kallorna"
        ]

        if failed_passes:
            reasons.append(
                "Misslyckade OCR-pass: " + ", ".join(failed_passes)
            )

        return MultiPassOCRResult(
            text=choose_display_text(
                pass_results,
                "",
                embedded_text,
            ),
            personnummer="",
            score=0,
            reasons=reasons,
            vote_count=0,
            total_sources=total_sources,
            checksum_valid=False,
            candidates={},
        )

    def candidate_sort_key(item: tuple[str, int]) -> tuple[int, bool, str]:
        personnummer_candidate, votes = item
        return votes, valid_luhn(personnummer_candidate), personnummer_candidate

    selected_candidate, vote_count = max(
        candidate_votes.items(),
        key=candidate_sort_key,
    )

    checksum_valid = valid_luhn(selected_candidate)
    vote_ratio = vote_count / max(1, total_sources)
    score = round(vote_ratio * 75)

    if checksum_valid:
        score += 20

    if vote_count >= 3:
        score += 5

    score = min(100, score)

    reasons = [
        (
            f"Personnummer {selected_candidate} hittades av "
            f"{vote_count} av {total_sources} OCR-kallor"
        ),
        (
            "Personnumrets kontrollsiffra ar giltig"
            if checksum_valid
            else "Personnumrets kontrollsiffra ar inte giltig"
        ),
    ]

    alternatives = [
        (personnummer_candidate, votes)
        for personnummer_candidate, votes in candidate_votes.most_common()
        if personnummer_candidate != selected_candidate
    ]

    if alternatives:
        reasons.append(
            "Alternativa OCR-kandidater: "
            + ", ".join(
                f"{personnummer_candidate} ({votes} roster)"
                for personnummer_candidate, votes in alternatives[:3]
            )
        )
    else:
        reasons.append("Inga konkurrerande personnummer hittades")

    supporting_sources = [
        result.name
        for result in pass_results
        if selected_candidate in result.candidates
    ]

    if selected_candidate in embedded_candidates:
        supporting_sources.append("PDF-textlager")

    if supporting_sources:
        reasons.append(
            "Kandidaten hittades av: " + ", ".join(supporting_sources)
        )

    if failed_passes:
        reasons.append(
            "Misslyckade OCR-pass: " + ", ".join(failed_passes)
        )

    display_text = choose_display_text(
        pass_results,
        selected_candidate,
        embedded_text,
    )

    return MultiPassOCRResult(
        text=display_text,
        personnummer=selected_candidate,
        score=score,
        reasons=reasons,
        vote_count=vote_count,
        total_sources=total_sources,
        checksum_valid=checksum_valid,
        candidates=dict(candidate_votes),
    )
