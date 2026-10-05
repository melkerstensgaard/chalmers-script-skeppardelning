from pathlib import Path
import fitz, json
from .analyzer import extract_text, analyze

def import_folder(db, root, progress=None):
    root = Path(root).resolve()
    pdfs = sorted(root.rglob("*.pdf"), key=lambda p: str(p).lower())
    db.set_root(root); errors=[]
    for number, pdf in enumerate(pdfs, 1):
        if progress: progress(number, len(pdfs), pdf.name)
        try:
            doc=fitz.open(pdf)
            for index,page in enumerate(doc):
                text=extract_text(page); score,pnr,reasons=analyze(text)
                db.upsert(dict(source_pdf=str(pdf),relative_pdf=str(pdf.relative_to(root)),page_number=index+1,page_count=len(doc),ocr_text=text,personnummer=pnr,script_personnummer=pnr,personnummer_source=("skript" if pnr else ""),score=score,score_reasons=json.dumps(reasons,ensure_ascii=False)))
            db.conn.commit(); doc.close()
        except Exception as exc: errors.append(f"{pdf}: {exc}")
    return len(pdfs), errors
