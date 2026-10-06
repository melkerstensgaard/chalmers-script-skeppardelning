import json
from pathlib import Path

import pymupdf

from .multi_pass_ocr import (
    analyze_page_multi_pass,
)


def import_folder(
    db,
    root,
    progress=None,
):
    root_path = Path(
        root,
    ).resolve()

    pdf_files = sorted(
        root_path.rglob(
            "*.pdf",
        ),
        key=lambda path: str(
            path,
        ).lower(),
    )

    db.set_root(
        root_path,
    )

    errors = []
    total_pdf_files = len(
        pdf_files,
    )

    for pdf_index, pdf_path in enumerate(
        pdf_files,
        start=1,
    ):
        document = None

        try:
            document = pymupdf.open(
                pdf_path,
            )

            page_count = len(
                document,
            )

            for page_index, page in enumerate(
                document,
            ):
                page_number = (
                    page_index + 1
                )

                if progress is not None:
                    progress(
                        pdf_index,
                        total_pdf_files,
                        (
                            f"{pdf_path.name} - "
                            f"sida {page_number}/"
                            f"{page_count}"
                        ),
                    )

                ocr_result = (
                    analyze_page_multi_pass(
                        page=page,
                        language="swe+eng",
                        dpi=300,
                    )
                )

                personnummer = (
                    ocr_result.personnummer
                )

                page_data = {
                    "source_pdf": str(
                        pdf_path,
                    ),
                    "relative_pdf": str(
                        pdf_path.relative_to(
                            root_path,
                        )
                    ),
                    "page_number": (
                        page_number
                    ),
                    "page_count": (
                        page_count
                    ),
                    "ocr_text": (
                        ocr_result.text
                    ),
                    "personnummer": (
                        personnummer
                    ),
                    "script_personnummer": (
                        personnummer
                    ),
                    "personnummer_source": (
                        "skript"
                        if personnummer
                        else ""
                    ),
                    "score": (
                        ocr_result.score
                    ),
                    "score_reasons": (
                        json.dumps(
                            ocr_result.reasons,
                            ensure_ascii=False,
                        )
                    ),
                }

                db.upsert(
                    page_data,
                )

            db.conn.commit()

        except Exception as exception:
            try:
                db.conn.rollback()
            except Exception:
                pass

            errors.append(
                (
                    f"{pdf_path}: "
                    f"{type(exception).__name__}: "
                    f"{exception}"
                )
            )

        finally:
            if document is not None:
                document.close()

    if progress is not None:
        progress(
            total_pdf_files,
            total_pdf_files,
            (
                "Import och flerpass-OCR "
                "är klar"
            ),
        )

    return (
        total_pdf_files,
        errors,
    )