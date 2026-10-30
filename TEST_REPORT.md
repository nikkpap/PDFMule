# PDF Mule v1.5.0 test report

Environment: Linux, Python 3.13, PyMuPDF 1.28.5, PyQt6 6.11 offscreen, Tesseract 5 (eng+ell+osd), Ghostscript, LibreOffice.
Inputs: the 16 Epson scans in `pdfs-test/` plus generated text PDFs (incl. a 90°-rotated page, bookmarks, Greek text).

## Static checks
- `py_compile`: PASS
- `pyflakes`: no undefined names / errors

## Engine tests (42 cases)
| Case | Result | Detail |
|---|---|---|
| merge | PASS | 0.0s  Done |
| remove | PASS | 0.0s  Done |
| remove-all-guard | PASS | expected guard: That would remove all 1 pages. Keep at least one page. |
| extract | PASS | 0.0s  Done |
| organize | PASS | 0.0s  Done |
| compress-smart | PASS | 4.5s  Done — 7.8 MB → 5.5 MB (−30%, smart (scanned pages)) |
| compress-lossless | PASS | 0.0s  Done — 4.5 KB → 2.0 KB (−55%, lossless) |
| compress-strong | PASS | 3.2s  Done — 7.8 MB → 489.4 KB (−94%, Ghostscript /screen) |
| repair | PASS | 0.0s  Done |
| rotate | PASS | 0.0s  Done |
| numbers | PASS | 0.0s  Done |
| watermark-diag | PASS | 0.0s  Done |
| crop | PASS | 0.0s  Done |
| edit | PASS | 0.0s  Done |
| forms | PASS | 0.0s  Done |
| protect | PASS | 0.0s  Done |
| unlock | PASS | 0.0s  Done |
| protected-input-msg | PASS | expected guard: prot.pdf is password-protected. Use Unlock PDF first (or enter the correct password). |
| sign | PASS | 0.0s  Done |
| redact | PASS | 0.0s  Done — 12 area(s) permanently redacted |
| compare | PASS | 0.0s  Done — texts are 35.7% similar |
| summarize | PASS | 0.0s  Done |
| markdown | PASS | 1.1s  Done |
| extract_text | PASS | 0.0s  Done — text + 1 unique image(s) |
| copy | PASS | 0.0s  Done |
| workflow | PASS | 0.0s  Done |
| pdf2jpg | PASS | 0.0s  Page 6/6 |
| pdf2ppt | PASS | 0.6s  Done |
| pdf2xls | PASS | 1.1s  Done — 0 table(s) found |
| pdf2word | PASS | 0.1s  Done |
| pdfa | PASS | 0.1s  Done |
| ocr-smart | PASS | 5.5s  Done |
| img2pdf | PASS | 0.0s  Done — 5 page(s) |
| img2pdf-frames | PASS | [(255, 0, 0), (0, 128, 0), (0, 0, 255), (255, 255, 255), (255, 255, 0)] |
| office2pdf | PASS | 1.4s  ✓ w.docx |
| vsplit-custom-out-of-range | PASS | 0.0s  Completed 1/1 PDF(s) → 2 file(s) |
| vsplit-smart | PASS | 0.0s  Completed 1/1 PDF(s) → 3 file(s) |
| vsplit-size-batch-parallel | PASS | 1.2s  Completed 6/6 PDF(s) → 14 file(s) |
| vsplit-with-bad-file | PASS | 1.1s  Completed 2/3 PDF(s) → 2 file(s) with 1 warning(s) |
| batch-compress-parallel | PASS | 7.5s  Completed 4/5 file(s) with 1 warning(s) |
| batch-rotate-seq | PASS | 0.0s  Completed 2/2 file(s) |
| eng_split | PASS | 0.0s  Completed 2/2 PDF(s) |

Visual checks (rendered pages): page numbers, Greek diagonal watermark, Edit text, signature and crop are
upright and correctly positioned on both portrait and rotated landscape pages; redaction removes the text.
OCR of an upside-down Epson scan: page straightened (OSD), text searchable, output 452 KB vs 448 KB input.

## GUI tests (offscreen, real dialogs, real worker threads/processes)
- Main window: responsive 3-column grid at 1000 px, Greek keyword search, empty-result hint.
- Tool dialog: add file + folder (16 inputs), compress run, completion dialog.
- OCR batch cancel: honoured at the next page (in-process) / immediately with parallel workers; no orphan processes or `.part` files.
- Organize: load, rotate ×2, delete, save (7 pages, rotation 180); fast double reload → no stale thumbnails.
- Visual Merge: two queued imports while loading (16 pages, 4 sources), rotate, duplicate, save (32 pages).
- Visual Split: 8 PDFs, fixed 1-page ranges, Balanced profile → 26 files.
- Spawn workers verified when the app runs as `__main__` (the real `.pyw` launch path).

Not tested here: Windows-specific paths (console-window suppression, `C:\Program Files` discovery), WeasyPrint HTML→PDF, online translation (no Internet access in the build sandbox).

## v1.5.0 additions
- Update check against a local copy of a "remote" PDFMule.pyw (1.5.1): newer version found, CHANGELOG notes parsed; same version → no update.
- One running copy (offscreen, two real processes): a second start of 1.5.0 exits and leaves the first running; starting 1.5.1 makes 1.5.0 close and 1.5.1 keeps running.
