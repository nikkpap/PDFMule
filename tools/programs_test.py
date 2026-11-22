"""Release check: PDF Mule installs Tesseract (+ Greek), Ghostscript and LibreOffice by itself, then
OCRs a Greek scan, makes a PDF/A and converts a Word file. Usage: python tools/programs_test.py [programs...]"""
import importlib.util
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

root = Path(__file__).resolve().parents[1]
work = Path(tempfile.mkdtemp(prefix="pdfmule_prog_"))
shutil.copyfile(root / "PDFMule.pyw", work / "pdfmule_under_test.py")
sys.path.insert(0, str(work))
spec = importlib.util.spec_from_file_location("pdfmule_under_test", work / "pdfmule_under_test.py")
m = importlib.util.module_from_spec(spec); sys.modules[spec.name] = m; spec.loader.exec_module(m)
import pymupdf as fitz

wanted = sys.argv[1:] or ["tesseract", "ghostscript", "libreoffice"]
log = lambda v, msg=None: print(f"  [{v}] {msg}", flush=True) if msg else None
fails = []
for key in wanted:
    t = time.time()
    print(f"== {key}: before = {m.PROGRAMS[key]['find']()}", flush=True)
    try:
        path = m.ensure_program(key, log)
        print(f"== {key}: OK {path} ({time.time() - t:.0f}s)", flush=True)
    except Exception as exc:
        print(f"::error::{key} install failed: {exc}"); fails.append(key)

font = next((f for f in ["C:/Windows/Fonts/arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf",
                         "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"] if Path(f).exists()), None)
if "tesseract" in wanted and "tesseract" not in fails:
    d = fitz.open(); p = d.new_page()
    p.insert_text((60, 120), "Καλημέρα κόσμε, δοκιμή αναγνώρισης κειμένου.", fontsize=20,
                  **({"fontname": "F0", "fontfile": font} if font else {}))
    scan = fitz.open(); sp = scan.new_page(width=p.rect.width, height=p.rect.height)
    sp.insert_image(sp.rect, stream=p.get_pixmap(dpi=200).tobytes("png")); scan.save(work / "scan.pdf")
    m.eng_ocr(log, [str(work / "scan.pdf")], str(work / "ocr.pdf"), {"ocr_mode": "Smart / Auto detect", "dpi": 200})
    text = fitz.open(work / "ocr.pdf")[0].get_text()
    print("OCR text:", text.strip())
    if "Καλημέρα" not in text:
        print("::error::Greek OCR did not read the test text"); fails.append("ocr")
if "ghostscript" in wanted and "ghostscript" not in fails:
    d = fitz.open(); d.new_page().insert_text((72, 72), "PDF/A test"); d.save(work / "a.pdf")
    m.eng_pdfa(log, [str(work / "a.pdf")], str(work / "pdfa.pdf"), {"pdfa": "2"})
    print("PDF/A size:", (work / "pdfa.pdf").stat().st_size)
if "libreoffice" in wanted and "libreoffice" not in fails:
    try:
        from docx import Document
        doc = Document(); doc.add_paragraph("Word to PDF test"); doc.save(work / "t.docx")
        src = work / "t.docx"
    except Exception:
        src = work / "t.txt"; src.write_text("Word to PDF test", encoding="utf-8")
    out = m.eng_office2pdf(log, [str(src)], str(work / "office"), {})
    print("Office → PDF:", out, "pages:", len(fitz.open(out)))
sys.exit(1 if fails else 0)
