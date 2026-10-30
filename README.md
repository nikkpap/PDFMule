# PDF Mule v1.5.0

A modern local-first PDF toolbox written in Python / PyQt6 and inspired by the category workflow of web PDF suites. PDF Mule is an independent project and is not affiliated with iLovePDF.

## Run it

1. Install [Python 3.10+](https://www.python.org/downloads/) for Windows (tick "Add python.exe to PATH").
2. Download `PDFMule.pyw` (plus `pdf_mule_logo.png` and `pdf_mule_icon.ico` for the logo/icon) and double-click it.
   On the first start it installs everything it needs by itself, then opens.
3. For OCR, Office → PDF and strong compression also install the free programs
   [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) (tick Greek),
   [LibreOffice](https://www.libreoffice.org/) and [Ghostscript](https://ghostscript.com/releases/gsdnld.html).
   About → Dependency status shows what is found.

## Updates

PDF Mule checks this repository a few seconds after it starts. When a newer version is here it shows what
changed and, with your OK, saves it next to your file as `PDFMule_v<version>.pyw` and opens it in place of the
old window. About → **Check for updates** checks on demand; the switch next to it turns the start check off.

Each release is published by committing the new `PDFMule.pyw` (with a higher `APP_VERSION`) and a
`## v<version>` section at the top of `CHANGELOG.md` to `main`.

## v1.5 highlights

- Updates over the internet (OTA) and one running copy (a newer version takes over from an older one).
- Automatic installation of all Python packages, core and optional, without questions.

## v1.4 highlights

- **Bug-fix & hardening release** — see CHANGELOG.md for the full audit list.
- **OCR keeps the original scan quality**: an invisible text layer is added over each page instead of
  replacing it with a re-encoded image; upside-down / sideways scans are straightened automatically.
- **Compress always keeps the smallest result** and never returns a file larger than the input; scans are
  recompressed even without Ghostscript.
- **Greek (and any Unicode) text** in watermarks, page numbers and Edit/Add text.
- Correct placement on **rotated pages** for numbers, watermark, text, signature and crop.
- **Instant cancel** of parallel batches; atomic saves (no half-written PDFs, ever).
- Drag & drop into every tool, smart drop on the main window, password prompts for encrypted inputs,
  keyboard shortcuts, responsive tool grid, Greek search keywords, “Open file / Open folder” on completion.

## v1.3 highlights

- **Visual / Batch Split** with an iLovePDF-like two-pane workspace.
  - Range / Pages / Size modes.
  - Range: Custom / Fixed / Smart.
  - Smart detects level-1 bookmarks, then blank separator pages, then uses a conservative fixed fallback.
  - Visual range cards show first/last page thumbnails and output page counts.
  - Batch import has no fixed PDF-count limit imposed by the app.
  - One output folder per source PDF avoids filename collisions.
- **Smart batch processing across the app.** Tools that logically operate independently on PDFs can accept many input PDFs and process them in isolated worker processes.
- **Performance profiles:** Smart / Auto, Balanced, Maximum speed, Low memory, Single worker.
- **Smart / Auto worker sizing** reduces concurrency automatically for very large input sets.
- **Visual Merge** keeps ordered writing for correctness while thumbnail import and save run in background workers. Low-memory / Smart / Maximum-speed profiles tune thumbnail size and open-document cache.
- **Visual Organize** now generates thumbnails in a background thread, so large PDFs no longer block the UI while previews are created.
- **Streaming Images → PDF** avoids holding every decoded image in RAM at once.
- **LibreOffice batch conversion** uses an isolated temporary LibreOffice user profile per conversion so parallel conversions do not fight over one profile.
- **Compress PDF** offers Smart / Auto, Lossless, Balanced, and Strong profiles. Smart uses lossless optimization for normal PDFs and can use Ghostscript `/ebook` on larger PDFs when Ghostscript is available.
- Batch jobs create a `_PDFMule_batch_report.txt` / `_PDFMule_split_report.txt` report.
- Source PDFs are not overwritten in-place by visual editors.

## Batch-capable PDF tools

Remove pages, Extract pages, Split, Compress, Repair, OCR, Word/PPT/Excel → PDF, PDF → JPG/Word/PPT/Excel/PDF-A, Rotate, Page numbers, Watermark, Crop, Add text, Forms report, Unlock, Protect, Sign, Redact, Summarize, Translate, PDF → Markdown, Extract text/images, and Copy/Rewrite.

Multi-input tools that intentionally combine inputs rather than batch them independently remain composition workflows: Visual Merge, Scan/Images → one PDF, Image → one PDF, Compare (exactly two PDFs), and Workflow Builder (Merge → Optimize → Protect).

## Core requirements

```text
PyQt6
PyMuPDF
Pillow
```

The app can offer to install missing core Python packages on first run.

## Optional Python integrations

- `pytesseract` — OCR wrapper
- `pdf2docx` — PDF → Word
- `python-pptx` — PDF → PowerPoint
- `openpyxl` — PDF table extraction → Excel
- `weasyprint` — HTML → PDF
- `pymupdf4llm` — improved PDF → Markdown
- `deep-translator` — translation

## Optional external applications

- **Tesseract OCR** — OCR engine
- **LibreOffice** — Word / PowerPoint / Excel → PDF
- **Ghostscript** — PDF/A and balanced/strong PDF compression

## Large jobs

The app does not set an arbitrary input-file count limit. Practical limits are available memory, disk space, operating-system resources, and the tools being used. For large PDF sets, use **Smart / Auto** or **Low memory**. Independent PDF jobs use separate worker processes to avoid sharing MuPDF state across concurrent jobs.

Cancellation is cooperative: the UI remains responsive and requests cancellation immediately, while an already-running external process or PDF operation is allowed to finish its current file safely.


## v1.3.3 compatibility update

- Uses the modern PyMuPDF import: `import pymupdf as fitz`.
- Core dependency detection now checks `pymupdf` instead of deprecated `fitz`.


## v1.3.4 dependency manager

- Dependency Status opens immediately and scans in a background thread.
- Optional Python packages are detected with `importlib.util.find_spec()` instead of importing heavy modules in the GUI thread.
- The add-on installer uses Qt `QProcess` and shows live pip stdout/stderr in an in-app terminal log.
- Installation ends with explicit success/failure confirmation and an automatic dependency rescan.
- Stop and close safeguards were added for a running installer.


## v1.3.6 startup presentation

- Added a Channelizer-style fake startup splash lasting about 3 seconds.
- Shows the PDF Mule logo, staged startup status, progress bar, version/build date and ALU DEV TEAM branding.
- Main window is created hidden and shown only after the splash completes.


## v1.3.8 Smart OCR languages

- OCR language control redesigned with Smart / Auto detection and Manual multi-select.
- Full standard Tesseract language list with language names and codes.
- Searchable checkbox list and Greek + English / English-only presets.
- Smart mode evaluates each PDF independently in batch jobs.
- Dominant-script threshold defaults to 70%; Greek ≥70% selects `ell`.
- Scanned PDFs without usable embedded text use sample-page Tesseract OSD script detection.
- Manual combinations are sent directly to Tesseract, e.g. `ell+eng`.
- Selected traineddata packs are validated before OCR begins.


## v1.3.8 Smart Translate

- Smart/Auto source-language translation via deep-translator / GoogleTranslator.
- Searchable source/target language selectors with quick Greek/English targets.
- Automatic OCR fallback for scanned or textless pages before translation.
- Batch translation remains per-PDF; online translation concurrency is capped to reduce provider throttling.
- Translation chunks use paragraph/sentence boundaries and transient-error retries.


## v1.3.9 folder import

- Added **Add Folder** beside file/PDF import controls.
- Folder import scans matching files recursively through subfolders.
- Visual Merge and Visual/Batch Split also accept dropped folders.
- Duplicate files are ignored automatically.
- Generic multi-file tools use their own file filter when scanning folders (PDF, Word, PowerPoint, Excel, images, etc.).
