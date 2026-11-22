# Changelog

## v1.7.0-rc1 — release candidate: review, fixes and easier tools

Includes everything below from 1.6.1 (never released on its own), plus:

- **Fix — PDF to Word stopped with "pixmap must be grayscale or rgb to write as png"** on PDFs with CMYK/ICC, Separation or DeviceN pictures (typical for printed books). Such pictures are now converted to RGB on the fly.
- PDF to Word: progress per page, an optional page range for big books, auto-install of the converter in the .pyw.
- Fix — the Windows installers could hang forever (PowerShell `Start-Process -Wait` waited for processes the Ghostscript installer leaves running); installers now have a time limit and report their exit code.
- HTML to PDF prints with Microsoft Edge / Google Chrome (no GTK needed); a domain typed without https:// works; local files fall back to LibreOffice.
- iPhone photos (HEIC/HEIF), WebP and GIF in Images → PDF; ODT/RTF/TXT, ODP/PPS, ODS/CSV in the Office → PDF tools.
- Errors in plain words (file open in another program, disk full, damaged PDF, no Internet …) with the technical details behind "Show Details" and in ~/.pdf_mule/logs/errors.log.
- Tool windows guide you in steps (1 add files · 2 options · 3 run), show file name · size · folder instead of long paths, and the order of files can be changed by dragging.
- A background update check waits until no tool is working before it asks.
- Release candidates: versions like 1.7.0-rc1 are published as pre-releases. People on a release candidate (or who tick "Also offer test versions" in About) get the next candidate and the final version; everyone else only gets final versions.

## v1.6.1 — automatic install of Tesseract, Ghostscript and LibreOffice

- Tesseract OCR, Ghostscript and LibreOffice are installed automatically the first time a tool needs them (Windows: winget or the official signed installers with one UAC prompt; Mac: Homebrew, or Terminal installs Homebrew; LibreOffice also from the official disk image; Linux: apt/dnf/pacman with sudo/pkexec).
- OCR languages live in ~/.pdf_mule/tessdata: English, Greek and orientation (osd) are set up with Tesseract; any other language is copied from the Tesseract installation or downloaded when a document needs it. The searchable-PDF support files are provided too.
- Batches install what they need once before the parallel workers start.
- Dependency status: "Install missing programs" installs all three at once.
- Release check: a new CI job installs the programs on fresh Windows, Mac (with and without Homebrew) and Linux machines and runs a Greek OCR, a PDF/A and a Word → PDF conversion.

## v1.6.0 — Windows, Mac and Linux releases with self-update

- GitHub Actions builds and publishes each version as a release: Windows .exe (optionally code-signed), macOS app for Apple silicon and Intel (ad-hoc signed, or Developer ID + notarized when the secrets are set), Linux binary, and the .pyw. Every build is smoke-tested before release.
- OTA for every build: the app reads the latest GitHub release, picks the file for its own system, downloads it with a progress bar (cancellable), verifies SHA-256 and restarts into the new version. Windows/Linux: saved as PDFMule_v<version>-… next to the old one, which is removed at the next start. Mac: the .app bundle is replaced in place. Script: PDFMule_v<version>.pyw next to the old one.
- Release notes shown in the user's language (English/Greek) as plain text, without Markdown symbols.
- macOS/Linux: Tesseract, Ghostscript and LibreOffice found in Homebrew / Applications / usual folders.
- Packaged builds include OCR, Word, Excel, PowerPoint, Markdown and translation add-ons.
- Smoke-test mode (PDFMULE_SMOKE_TEST) used by the release checks.

## v1.5.0 — updates over the internet (OTA), automatic setup

- **Updates (OTA)**: a few seconds after start (in the background) and with About → Check for updates, PDF Mule looks for a newer `PDFMule.pyw` on GitHub, shows what is new and, with your OK, saves it next to the current file as `PDFMule_v<version>.pyw` and opens it in place of this window. Settings and recent outputs stay; the old file stays too. A switch in About turns the start check off.
- **One running copy**: starting a newer version closes the older one (unless it is in the middle of a job); starting the same version again brings the open window to the front.
- **Automatic setup**: on the first start PyQt6, PyMuPDF and Pillow are installed by themselves behind a small "please wait" window (no questions, no console), and the optional add-ons (OCR, Word, Excel, PowerPoint, Markdown, translation, HTML) are installed quietly in the background once per version. Dependency status shows the result.
- A downloaded update that is incomplete is never started (it is compiled before it is saved).

## v1.4.0 — audit, fixes, optimisation, usability

### Bugs fixed
- **Multi-page TIFF/GIF → PDF produced N copies of the last frame** (ImageSequence iterator was collected into a list). Every frame is now kept.
- **Custom split ranges beyond the page count** were swapped instead of clamped (e.g. 15–20 on a 10-page PDF became 10–15 → error). Out-of-range ranges are now clamped or skipped.
- **Split/batch worker crashed the whole batch** when a file could not be opened (`src`/`doc` referenced in `finally` before assignment).
- **`find_exe` could return the wrong program** (e.g. LibreOffice when looking for Tesseract): the Windows fallback list matched any known exe. Ghostscript is now found in any installed version folder, not only 10.05.1/10.06.0.
- **Visual Merge cancel deleted an existing file** chosen as the output target. Saves are now atomic, so nothing is deleted.
- **Organize: saving an encrypted PDF failed** (password not passed to the save worker).
- **Organize: stale thumbnails** from a previous PDF could land on the new one when reloading quickly.
- **Esc / window X / Close while a job ran destroyed a dialog with a running QThread** (Qt crash). All dialogs now ask, stop the job and close safely.
- **Cancellation was swallowed** by engines that use `except Exception` fallbacks (`JobCancelled` is now a `BaseException`).
- **Text, page numbers, signatures and crop landed in the wrong place on rotated pages.**
- **Greek text became “?”** in watermark / Edit text / page numbers (Helvetica is Latin-1 only). A Unicode TrueType font is embedded and subset.
- Watermark text larger than the page was silently not drawn; it now shrinks to fit. Image watermarks honour opacity and are embedded once.
- Removing all pages raised a cryptic MuPDF error; now a clear message.
- `page_set` crashed on `5-` / `-3`; now supports open ranges and `end`/`last`.
- Translation could crash on a `None` provider response or exceed the provider limit with very long tokens.
- Visual Split re-opened the current PDF (and re-asked its password) every time files were added or removed.
- Merge button showed “Merge_Save” (unescaped `&` mnemonic).
- Installer used `pythonw.exe -m pip` (no output) and, in a frozen EXE, would have re-launched the app instead of pip.
- Reports still said “PDFForge”.

### Performance
- OCR output keeps the original page images (invisible text layer via Tesseract `textonly_pdf`) — files stay ~the input size instead of being re-encoded.
- Compress tries the relevant strategies (lossless, PyMuPDF image recompression, Ghostscript) and keeps the smallest; scanned PDFs shrink ~30% in Smart mode without Ghostscript.
- PDF → JPG renders straight from pixmap samples (no PNG round-trip). JPEG images are embedded byte-for-byte in Images → PDF.
- Contiguous page ranges are copied with one `insert_pdf` call.
- Progress signals are throttled (no GUI flooding on 2 000-page jobs); log is capped.
- Parallel batches poll for cancel every 0.3 s and terminate worker processes immediately; temp `.part` files are cleaned up.
- Extract text/images de-duplicates repeated images; Visual Merge keeps a small LRU of open preview documents.
- External tools run without console windows on Windows.

### Usability
- Drag & drop files/folders into every tool; smart drop on the main window offers the tools that fit what was dropped (PDFs, images, Office, folders). Command-line paths (“Open with”) work too.
- Encrypted inputs: the app asks for the password before starting (tries the last password first).
- Completion dialog with elapsed time, **Open file** and **Open folder**; batch log shows `[n/N]` and remaining time.
- OCR: skip pages that already have text, auto-straighten, default 300 dpi.
- New options: rotate selected pages only; page-number formats incl. “Σελίδα 1 από N” and skip cover; diagonal/horizontal watermark; signature on first/last/all pages + position; protect with confirm password and copy permission; multiple redaction terms; Images → PDF page size (A4/Letter/A3) and margins.
- Keyboard shortcuts: Ctrl+F search (Enter opens first match), Del, Ctrl+A, Ctrl+D, Ctrl+←/→, Ctrl+S.
- Responsive tool grid; Greek/synonym search keywords; “no results” hint.
- Tool dialogs scroll on small screens; output can’t overwrite its input; last output folder remembered.
- Visible spin-box arrows and checkbox ticks in both themes; label background artefacts removed; light-theme hover/disabled states.
- Recent outputs: date, Open file/folder, double-click, clear, hides missing files.
- Splash can be skipped with a click.

## v1.3.9

- Added **Add Folder** to Visual Merge, Visual / Batch Split, and all compatible multi-file tool dialogs.
- Folder import recursively scans subfolders and de-duplicates matching files.
- Visual Merge and Split now support drag-and-drop of folders as well as PDFs.
- Generic folder import respects the active tool file type filter.

## v1.3.8

- Upgraded Translate PDF with Smart/Auto source detection.
- Added searchable source/target language selectors and Greek/English quick targets.
- Added automatic OCR fallback for scanned/textless pages.
- Added smarter translation chunking and retry handling.
- Bounded batch translation concurrency to reduce online-provider rate limits.

## v1.3.7

- Added Smart OCR language detection with a configurable 70% dominant-script threshold.
- Greek-dominant documents automatically use Tesseract `ell`.
- Added scan fallback using Tesseract OSD on sample pages when the PDF has no embedded text.
- Added searchable full Tesseract language list and checkbox multi-select.
- Added Greek + English and English-only presets.
- Batch OCR detects language independently for every PDF.
- Added validation for missing selected Tesseract traineddata packs.

## v1.3.6

- Added a frameless centered **PDF Mule startup splash** inspired by The_Channelizer.
- Fake startup sequence lasts approximately 3 seconds with staged messages and a smooth progress bar.
- Splash includes the approved PDF Mule logo, version, build date and ALU DEV TEAM branding.
- Main window stays hidden until the startup animation completes.

## v1.3.5

- Corrected the LinkedIn profile URL in About to `https://www.linkedin.com/in/nickparidis/`.

## v1.3.4

- Rebuilt Dependency Status as a non-blocking dependency dashboard.
- Removed heavy optional-library imports from the GUI thread; module detection now uses `find_spec`.
- Added live in-app terminal output for pip installs/upgrades using `QProcess`.
- Added indeterminate install progress, Stop control, final success/failure confirmation, and automatic post-install rescan.
- Prevents the apparent Windows “Not Responding” state caused by first-time imports of heavy optional libraries.

## v1.3.3

- Replaced deprecated `import fitz` with `import pymupdf as fitz`.
- Updated bootstrap dependency detection from `fitz` to `pymupdf`.
- Existing `fitz.*` engine calls remain unchanged through the alias.
- Removes the PyMuPDF deprecation warning and improves forward compatibility.


## v1.3.2

- Rebranded the app to **PDF Mule**.
- Integrated the first approved **PDF Mule** logo into the **About** window.
- Added application icon assets (`pdf_mule_logo.png` and `pdf_mule_icon.ico`) for the app and Windows EXE builds.
- Upgraded the **About** dialog to match the ALU DEV TEAM style used in the other NikkPap desktop apps.
- Updated PyInstaller build instructions to bundle the logo/icon resources.

## 1.3.1

### About / identity
- Replaced the basic About message box with a dedicated modern About dialog.
- Added ALU DEV TEAM @ 2026 identity.
- Added Nikolaos K. Paridis シ (nikkpap), email, GitHub and LinkedIn links.
- Added the recurring community motto used across the developer's other apps.
- Kept the independent-project / no-affiliation notice.

## 1.3.0

### New
- Visual / Batch Split workspace with Range / Pages / Size tabs.
- Custom ranges with multiple range rows and visual first/last-page preview cards.
- Fixed-size page chunks.
- Smart splitting from level-1 PDF bookmarks, blank separator pages, then fallback chunks.
- Approximate target-file-size splitting.
- Batch Split of many PDFs with isolated worker processes.
- Generic Smart Batch dispatcher for nearly every independent one-PDF tool.
- Performance profiles: Smart / Auto, Balanced, Maximum speed, Low memory, Single worker.
- Smart worker reduction for unusually large input sets.
- Batch result reports.
- Improved thumbnail-card styling in light and dark themes.
- Visual Merge performance profile.
- Background thumbnail generation in Organize PDF.

### Performance / reliability
- Images → PDF changed to streaming one image/frame at a time instead of retaining all decoded images.
- Parallel LibreOffice conversions receive isolated temporary user profiles.
- Batch output naming is collision-resistant through numbered source prefixes / source subfolders.
- Input controls are locked while jobs run to avoid changing a live job definition.
- Multi-file drag/drop now asks whether the PDFs should be visually merged or batch split.
- Visual Split caches preview thumbnails with a bounded cache.
- Encrypted PDFs in Visual Split prompt for passwords before background child processes start.
- Organize rotation reuses already-rendered thumbnails instead of re-rendering every selected page.
- Organize prevents closing while a save operation is still writing.

### Compression
- Smart / Auto, Lossless, Balanced and Strong compression profiles.
- Balanced / Strong use Ghostscript when available; otherwise PDF Mule safely falls back to lossless PyMuPDF optimization.

### Code audit fixes
- Simplified page-edit progress calculation to avoid repeated list conversion inside loops.
- Added output existence verification for LibreOffice conversions.
- Added explicit per-input error continuation/reporting for batch operations.
- Added worker/process bounds and low-memory behavior.

## 1.2.0
- Unlimited-count Batch Split input selection.
- Isolated process workers for parallel split.
- Visual Merge with cross-document drag/drop.
- Background thumbnail loading and cancellable merge save.

## 1.1.0
- Visual Organize PDF with page thumbnails, drag/drop reorder, rotate/delete, and preview.

## 1.0.0
- Initial category-based PDF toolbox.
