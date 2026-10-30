# PDFMule.pyw
# Modern local PDF toolbox inspired by the category-based workflow of iLovePDF.
# Created by Nikolaos K. Paridis シ (nikkpap) — ALU DEV TEAM @ 2026.
# Single-file desktop app for Windows/Linux/macOS.
# Core: PyQt6 + PyMuPDF + Pillow. Optional adapters add OCR/Office/Word/PPT/AI-like helpers.

from __future__ import annotations

import os
import sys
import re
import json
import math
import shutil
import subprocess
import tempfile
import time
import traceback
import difflib
import multiprocessing
import importlib.util
import functools
import glob
from concurrent.futures import ProcessPoolExecutor, as_completed, wait, FIRST_COMPLETED
from collections import OrderedDict
from pathlib import Path
from dataclasses import dataclass
from typing import Callable, Optional, Any

APP_NAME = "PDF Mule"
APP_VERSION = "1.5.0"
BUILD_DATE = "2026-10-30"
APP_HOME = Path.home() / ".pdf_mule"
SETTINGS_FILE = APP_HOME / "settings.json"
try:
    APP_HOME.mkdir(parents=True, exist_ok=True)
except OSError:
    pass  # settings simply won't persist; the app still works


def app_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


def asset_path(*names: str) -> Path:
    base = app_base_dir()
    for name in names:
        p = base / name
        if p.exists():
            return p
    return base / names[0]


def app_icon_path() -> Path:
    return asset_path("pdf_mule_icon.ico", "pdf_mule_icon.png", "pdf_mule_logo.png")


def app_logo_path() -> Path:
    return asset_path("pdf_mule_logo.png", "pdf_mule_reliable_document_packhorse.png")


def python_console_executable() -> str:
    """Return python.exe instead of pythonw.exe so pip output can be captured."""
    exe = Path(sys.executable)
    if exe.name.lower() == "pythonw.exe":
        candidate = exe.with_name("python.exe")
        if candidate.exists():
            return str(candidate)
    return str(exe)


def _no_window_kwargs() -> dict:
    """Hide console windows for child processes on Windows (pythonw / frozen EXE)."""
    if os.name == "nt":
        return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    return {}


def run_quiet(cmd, timeout=600):
    """subprocess.run with captured text output, no console flash, and friendly timeout errors."""
    try:
        return subprocess.run(cmd, capture_output=True, text=True, errors="replace",
                              timeout=timeout, **_no_window_kwargs())
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"{Path(cmd[0]).name} did not finish within {timeout} s and was stopped.")


# ----------------------------- Bootstrap core deps -----------------------------

def _pip_install(packages, quiet_window=True) -> tuple[bool, str]:
    """pip install; retried with --user when the Python folder is not writable."""
    base = [python_console_executable(), "-m", "pip", "install", "--upgrade", "--disable-pip-version-check"]
    out = ""
    for extra in ([], ["--user"]):
        try:
            cp = subprocess.run(base + extra + list(packages), capture_output=True, text=True, errors="replace",
                                timeout=1800, **(_no_window_kwargs() if quiet_window else {}))
        except Exception as exc:
            return False, str(exc)
        out = (cp.stdout or "") + (cp.stderr or "")
        if cp.returncode == 0:
            return True, out
    return False, out


def _bootstrap_core():
    """Installs PyQt6 / PyMuPDF / Pillow by itself on the first start, behind a small
    'please wait' window, then restarts the app. No questions, no console."""
    missing = []
    for import_name, pip_name in [("PyQt6", "PyQt6"), ("pymupdf", "PyMuPDF"), ("PIL", "Pillow")]:
        if importlib.util.find_spec(import_name) is None:
            missing.append(pip_name)
    if not missing:
        return
    if getattr(sys, "frozen", False):
        raise SystemExit("Missing core dependencies in this build: " + ", ".join(missing))
    result = {}
    try:
        import threading
        import tkinter as tk
        root = tk.Tk(); root.title(f"{APP_NAME} – first start"); root.geometry("460x150"); root.resizable(False, False)
        tk.Label(root, text=f"{APP_NAME} is installing what it needs…", font=("Segoe UI", 12, "bold")).pack(pady=(22, 6))
        tk.Label(root, text=", ".join(missing) + "\nThis happens only once and can take a few minutes.",
                 font=("Segoe UI", 9)).pack()
        dots = tk.Label(root, text="", font=("Segoe UI", 14)); dots.pack(pady=6)

        def work():
            result["ok"], result["log"] = _pip_install(missing)

        th = threading.Thread(target=work, daemon=True); th.start()

        def poll(n=0):
            if th.is_alive():
                dots.config(text="●" * (n % 6 + 1)); root.after(300, poll, n + 1)
            else:
                root.destroy()
        poll(); root.mainloop()
    except Exception:
        result["ok"], result["log"] = _pip_install(missing)
    if result.get("ok"):
        os.execv(sys.executable, [sys.executable, *sys.argv])
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk(); r.withdraw()
        messagebox.showerror(f"{APP_NAME} – setup failed",
                             "Could not install: " + ", ".join(missing) + "\n\nCheck the Internet connection and try again.\n\n"
                             + (result.get("log") or "")[-700:])
        r.destroy()
    except Exception:
        pass
    raise SystemExit("Missing core dependencies: " + ", ".join(missing))

_bootstrap_core()

import pymupdf as fitz  # Modern PyMuPDF API; keep fitz alias for existing code
from PIL import Image
from PyQt6.QtCore import Qt, QSize, QThread, pyqtSignal, QUrl, QProcess, QTimer
from PyQt6.QtGui import QColor, QFont, QIcon, QDesktopServices, QDragEnterEvent, QDropEvent, QPixmap, QImage, QTransform, QShortcut, QKeySequence


def load_app_icon() -> QIcon:
    p = app_icon_path()
    try:
        if p.exists():
            return QIcon(str(p))
    except Exception:
        pass
    return QIcon()

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QScrollArea, QFrame, QFileDialog, QMessageBox,
    QDialog, QListWidget, QListWidgetItem, QAbstractItemView, QComboBox, QSpinBox,
    QDoubleSpinBox, QCheckBox, QProgressBar, QPlainTextEdit, QStackedWidget,
    QFormLayout, QDialogButtonBox, QInputDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QGroupBox, QRadioButton, QButtonGroup, QTextEdit, QListView, QSplitter,
    QStackedLayout, QSizePolicy
)

# ----------------------------- Theme -----------------------------
DARK_QSS = r"""
QWidget { background:#0f1115; color:#e9edf3; font-family:'Segoe UI'; font-size:10pt; }
QLabel, QCheckBox, QRadioButton { background:transparent; }
QMainWindow { background:#0f1115; }
QFrame#sidebar { background:#151922; border-right:1px solid #252b37; }
QLabel#brand { font-size:20pt; font-weight:800; color:#ffffff; }
QLabel#brandAccent { color:#ff4d5f; }
QLabel#heroTitle { font-size:22pt; font-weight:800; }
QLabel#heroSub { color:#9aa5b5; font-size:10.5pt; }
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QTextEdit, QPlainTextEdit, QListWidget, QTableWidget {
  background:#161b24; border:1px solid #2a3240; border-radius:10px; padding:8px; selection-background-color:#ff4d5f;
}
QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus { border:1px solid #ff4d5f; }
QPushButton { background:#202632; border:1px solid #2b3443; border-radius:10px; padding:9px 12px; font-weight:600; }
QPushButton:hover { background:#2a3240; border-color:#46536a; }
QPushButton#primary { background:#ff4d5f; color:white; border:none; }
QPushButton#primary:hover { background:#ff6170; }
QPushButton#nav { text-align:left; padding:10px 12px; border:none; background:transparent; color:#b9c2cf; }
QPushButton#nav:hover { background:#202632; color:white; }
QPushButton#nav[active="true"] { background:#2a2024; color:#ff7180; border-left:3px solid #ff4d5f; }
QFrame#toolCard { background:#171c25; border:1px solid #262e3a; border-radius:14px; }
QFrame#toolCard:hover { border:1px solid #ff4d5f; background:#1b212c; }
QLabel#toolIcon { font-size:20pt; background:#242b36; border-radius:12px; padding:8px; }
QLabel#toolTitle { font-weight:800; font-size:11pt; }
QLabel#toolDesc { color:#99a4b4; }
QFrame#topbar { background:#11151c; border-bottom:1px solid #252b37; }
QFrame#previewPanel { background:#121720; border:1px solid #2a3240; border-radius:16px; }
QFrame#splitRangeCard { background:#151a22; border:1px dashed #4b5568; border-radius:14px; }
QFrame#splitRangeCard[selected="true"] { border:2px solid #ff4d5f; background:#1d171c; }
QPushButton#modeTab { background:#161b24; color:#9aa5b5; border:1px solid #2a3240; border-radius:10px; padding:12px; }
QPushButton#modeTab[active="true"] { background:#251a1f; color:#ff7180; border:1px solid #ff4d5f; }
QLabel#pagePaper { background:#ffffff; border:1px solid #d9dde5; border-radius:6px; color:#20242c; }
QLabel#badge { background:#2a3240; color:#dce3ee; border-radius:9px; padding:3px 7px; font-size:8.5pt; }
QListWidget#thumbView { background:#10141b; border:1px solid #252d39; border-radius:14px; padding:8px; }
QListWidget#thumbView::item { background:#171c25; border:1px solid #2a3240; border-radius:12px; padding:8px; margin:4px; }
QListWidget#thumbView::item:selected { background:#281b20; border:2px solid #ff4d5f; color:white; }


QProgressBar { border:1px solid #2b3443; border-radius:7px; text-align:center; background:#161b24; }
QProgressBar::chunk { background:#ff4d5f; border-radius:6px; }
QGroupBox { border:1px solid #2a3240; border-radius:10px; margin-top:12px; padding-top:12px; font-weight:700; }
QGroupBox::title { subcontrol-origin:margin; left:12px; padding:0 5px; }
QScrollBar:vertical { background:#11151c; width:10px; margin:0; }
QScrollBar::handle:vertical { background:#394355; border-radius:5px; min-height:24px; }
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical { height:0; }
QPushButton:disabled { color:#5d6676; }
QPushButton#primary:disabled { background:#5a2a31; color:#b9a0a5; }
"""

LIGHT_QSS = r"""
QWidget { background:#f6f7f9; color:#20242c; font-family:'Segoe UI'; font-size:10pt; }
QLabel, QCheckBox, QRadioButton { background:transparent; }
QMainWindow { background:#f6f7f9; }
QFrame#sidebar { background:#ffffff; border-right:1px solid #e3e7ed; }
QLabel#brand { font-size:20pt; font-weight:800; color:#20242c; }
QLabel#heroTitle { font-size:22pt; font-weight:800; }
QLabel#heroSub { color:#697586; font-size:10.5pt; }
QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QTextEdit, QPlainTextEdit, QListWidget, QTableWidget {
  background:#ffffff; border:1px solid #d8dee8; border-radius:10px; padding:8px; selection-background-color:#ff4d5f;
}
QPushButton { background:#ffffff; border:1px solid #d8dee8; border-radius:10px; padding:9px 12px; font-weight:600; }
QPushButton:hover { background:#f0f3f7; }
QPushButton#primary { background:#ff4d5f; color:white; border:none; }
QPushButton#nav { text-align:left; padding:10px 12px; border:none; background:transparent; color:#596579; }
QPushButton#nav:hover { background:#f3f5f8; color:#20242c; }
QPushButton#nav[active="true"] { background:#fff0f2; color:#e8374a; border-left:3px solid #ff4d5f; }
QFrame#toolCard { background:#ffffff; border:1px solid #e3e7ed; border-radius:14px; }
QLabel#toolIcon { font-size:20pt; background:#f2f4f7; border-radius:12px; padding:8px; }
QLabel#toolTitle { font-weight:800; font-size:11pt; }
QLabel#toolDesc { color:#697586; }
QFrame#topbar { background:#ffffff; border-bottom:1px solid #e3e7ed; }
QFrame#previewPanel { background:#ffffff; border:1px solid #e0e5ec; border-radius:16px; }
QFrame#splitRangeCard { background:#ffffff; border:1px dashed #aab2bf; border-radius:14px; }
QFrame#splitRangeCard[selected="true"] { border:2px solid #ef3340; background:#fff7f8; }
QPushButton#modeTab { background:#f4f6f9; color:#707b8d; border:1px solid #e0e5ec; border-radius:10px; padding:12px; }
QPushButton#modeTab[active="true"] { background:#fff5f6; color:#ef3340; border:1px solid #ef3340; }
QLabel#pagePaper { background:#ffffff; border:1px solid #d7dce4; border-radius:6px; color:#20242c; }
QLabel#badge { background:#eef1f5; color:#566173; border-radius:9px; padding:3px 7px; font-size:8.5pt; }
QListWidget#thumbView { background:#f2f4f7; border:1px solid #e0e5ec; border-radius:14px; padding:8px; }
QListWidget#thumbView::item { background:#ffffff; border:1px solid #dfe4eb; border-radius:12px; padding:8px; margin:4px; }
QListWidget#thumbView::item:selected { background:#fff4f5; border:2px solid #ef3340; color:#20242c; }


QFrame#toolCard:hover { border:1px solid #ff4d5f; background:#fffafa; }
QLineEdit:focus, QComboBox:focus, QTextEdit:focus, QPlainTextEdit:focus { border:1px solid #ff4d5f; }
QPushButton#primary:hover { background:#ff6170; }
QPushButton:disabled { color:#a6aebb; }
QPushButton#primary:disabled { background:#f3b3ba; color:#ffffff; }
QProgressBar { border:1px solid #d8dee8; border-radius:7px; text-align:center; background:#ffffff; }
QProgressBar::chunk { background:#ff4d5f; border-radius:6px; }
QGroupBox { border:1px solid #e0e5ec; border-radius:10px; margin-top:12px; padding-top:12px; font-weight:700; }
QGroupBox::title { subcontrol-origin:margin; left:12px; padding:0 5px; }
QScrollBar:vertical { background:#f2f4f7; width:10px; margin:0; }
QScrollBar::handle:vertical { background:#c5ccd7; border-radius:5px; min-height:24px; }
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical { height:0; }
"""

def _ui_asset_dir() -> Path:
    d = APP_HOME / "ui"
    try:
        d.mkdir(parents=True, exist_ok=True)
    except OSError:
        d = Path(tempfile.gettempdir()) / "pdfmule_ui"; d.mkdir(parents=True, exist_ok=True)
    return d


@functools.lru_cache(maxsize=2)
def _ui_assets(theme: str) -> dict:
    """Draw tiny arrow / check-mark PNGs once. Styling QSpinBox buttons or QCheckBox
    indicators via QSS removes the native glyphs, so we supply our own."""
    from PyQt6.QtGui import QPainter, QPolygonF, QPen
    from PyQt6.QtCore import QPointF
    col = QColor("#c9d1dc" if theme == "dark" else "#4a5568")
    out = {}
    d = _ui_asset_dir()
    for name, pts in (("up", [(2, 9), (6, 4), (10, 9)]), ("down", [(2, 4), (6, 9), (10, 4)])):
        img = QImage(12, 12, QImage.Format.Format_ARGB32); img.fill(Qt.GlobalColor.transparent)
        p = QPainter(img); p.setRenderHint(QPainter.RenderHint.Antialiasing)
        p.setPen(Qt.PenStyle.NoPen); p.setBrush(col)
        p.drawPolygon(QPolygonF([QPointF(x, y) for x, y in pts])); p.end()
        f = d / f"{name}_{theme}.png"; img.save(str(f)); out[name] = f.as_posix()
    img = QImage(14, 14, QImage.Format.Format_ARGB32); img.fill(Qt.GlobalColor.transparent)
    p = QPainter(img); p.setRenderHint(QPainter.RenderHint.Antialiasing)
    pen = QPen(QColor("#ffffff")); pen.setWidthF(2.2); p.setPen(pen)
    p.drawPolyline(QPolygonF([QPointF(3, 7.5), QPointF(6, 10.5), QPointF(11, 4)])); p.end()
    f = d / "check.png"; img.save(str(f)); out["check"] = f.as_posix()
    return out


def themed_qss(theme: str) -> str:
    base = DARK_QSS if theme == "dark" else LIGHT_QSS
    try:
        a = _ui_assets(theme)
    except Exception:
        return base
    border = "#3a4456" if theme == "dark" else "#b9c2cf"
    btn = "#232a36" if theme == "dark" else "#eef1f5"
    box_bg = "#161b24" if theme == "dark" else "#ffffff"
    return base + f"""
QAbstractSpinBox {{ padding-right:24px; }}
QAbstractSpinBox::up-button, QAbstractSpinBox::down-button {{ subcontrol-origin:border; width:20px; border:none; background:{btn}; }}
QAbstractSpinBox::up-button {{ subcontrol-position:top right; border-top-right-radius:9px; }}
QAbstractSpinBox::down-button {{ subcontrol-position:bottom right; border-bottom-right-radius:9px; }}
QAbstractSpinBox::up-arrow {{ image:url("{a['up']}"); width:10px; height:10px; }}
QAbstractSpinBox::down-arrow {{ image:url("{a['down']}"); width:10px; height:10px; }}
QAbstractSpinBox::up-button:hover, QAbstractSpinBox::down-button:hover {{ background:#ff4d5f; }}
QCheckBox {{ spacing:8px; }}
QCheckBox::indicator, QListWidget::indicator {{ width:16px; height:16px; border:1px solid {border}; border-radius:4px; background:{box_bg}; }}
QCheckBox::indicator:hover, QListWidget::indicator:hover {{ border-color:#ff4d5f; }}
QCheckBox::indicator:checked, QListWidget::indicator:checked {{ background:#ff4d5f; border-color:#ff4d5f; image:url("{a['check']}"); }}
QCheckBox::indicator:disabled {{ background:{btn}; }}
QComboBox::down-arrow {{ image:url("{a['down']}"); width:10px; height:10px; }}
QComboBox::drop-down {{ border:none; width:24px; }}
QToolTip {{ background:{box_bg}; color:{'#e9edf3' if theme == 'dark' else '#20242c'}; border:1px solid {border}; padding:5px; }}
"""


# ----------------------------- Model -----------------------------
@dataclass(frozen=True)
class ToolDef:
    key: str
    title: str
    category: str
    icon: str
    desc: str
    multi: bool = False
    extensions: str = "PDF files (*.pdf)"
    batchable: bool = False

TOOLS = [
    ToolDef("merge", "Visual Merge PDF", "Organize PDF", "🧩", "Merge unlimited PDFs visually: thumbnails, drag/drop pages, rotate, duplicate and delete.", True),
    ToolDef("split", "Visual / Batch Split PDF", "Organize PDF", "✂", "Custom, fixed, smart, selected-page and target-size splitting for one or many PDFs.", True, batchable=True),
    ToolDef("remove", "Remove pages", "Organize PDF", "🗑", "Delete the same selected pages from one or many PDFs.", True, batchable=True),
    ToolDef("extract", "Extract pages", "Organize PDF", "📤", "Extract the same page selection from one or many PDFs.", True, batchable=True),
    ToolDef("organize", "Organize PDF", "Organize PDF", "↕", "Visual thumbnails: reorder, rotate, delete and preview pages."),
    ToolDef("scan", "Scan / Images to PDF", "Organize PDF", "📷", "Build one PDF from scans or image files.", True, "Images (*.jpg *.jpeg *.png *.bmp *.tif *.tiff)"),

    ToolDef("compress", "Compress PDF", "Optimize PDF", "🗜", "Batch optimize and recompress PDF objects with smart background processing.", True, batchable=True),
    ToolDef("repair", "Repair PDF", "Optimize PDF", "🛠", "Batch rewrite PDFs into clean files and rebuild object tables.", True, batchable=True),
    ToolDef("ocr", "OCR PDF", "Optimize PDF", "🔎", "Batch OCR scanned PDFs with Tesseract while the UI remains responsive.", True, batchable=True),

    ToolDef("img2pdf", "JPG / Image to PDF", "Convert to PDF", "🖼", "Convert one or more images into one PDF.", True, "Images (*.jpg *.jpeg *.png *.bmp *.tif *.tiff)"),
    ToolDef("word2pdf", "Word to PDF", "Convert to PDF", "W", "Batch convert DOC/DOCX using LibreOffice.", True, "Word files (*.doc *.docx)", True),
    ToolDef("ppt2pdf", "PowerPoint to PDF", "Convert to PDF", "P", "Batch convert PPT/PPTX using LibreOffice.", True, "PowerPoint (*.ppt *.pptx)", True),
    ToolDef("xls2pdf", "Excel to PDF", "Convert to PDF", "X", "Batch convert XLS/XLSX using LibreOffice.", True, "Excel files (*.xls *.xlsx)", True),
    ToolDef("html2pdf", "HTML to PDF", "Convert to PDF", "🌐", "Convert a local HTML page or URL using WeasyPrint."),

    ToolDef("pdf2jpg", "PDF to JPG", "Convert from PDF", "🖼", "Batch render every PDF page to JPG images.", True, batchable=True),
    ToolDef("pdf2word", "PDF to Word", "Convert from PDF", "W", "Batch convert PDFs to DOCX with pdf2docx.", True, batchable=True),
    ToolDef("pdf2ppt", "PDF to PowerPoint", "Convert from PDF", "P", "Batch create PowerPoint decks from PDF pages.", True, batchable=True),
    ToolDef("pdf2xls", "PDF to Excel", "Convert from PDF", "X", "Batch extract detected tables to XLSX workbooks.", True, batchable=True),
    ToolDef("pdfa", "PDF to PDF/A", "Convert from PDF", "A", "Batch archival conversion through Ghostscript.", True, batchable=True),

    ToolDef("rotate", "Rotate PDF", "Edit PDF", "↻", "Batch rotate pages by 90°, 180° or 270°.", True, batchable=True),
    ToolDef("numbers", "Page numbers", "Edit PDF", "#", "Batch add configurable page numbers.", True, batchable=True),
    ToolDef("watermark", "Watermark", "Edit PDF", "💧", "Batch stamp text or an image on PDF pages.", True, batchable=True),
    ToolDef("crop", "Crop PDF", "Edit PDF", "▣", "Batch crop page margins by millimetres.", True, batchable=True),
    ToolDef("edit", "Edit / Add text", "Edit PDF", "✎", "Batch overlay text on one page or all pages.", True, batchable=True),
    ToolDef("forms", "PDF Forms", "Edit PDF", "☑", "Batch inspect interactive form fields and export reports.", True, batchable=True),

    ToolDef("unlock", "Unlock PDF", "PDF Security", "🔓", "Batch remove open passwords when you know the password.", True, batchable=True),
    ToolDef("protect", "Protect PDF", "PDF Security", "🔒", "Batch encrypt PDFs with AES-256 password protection.", True, batchable=True),
    ToolDef("sign", "Sign PDF", "PDF Security", "✍", "Batch place a signature image on documents.", True, batchable=True),
    ToolDef("redact", "Redact PDF", "PDF Security", "⬛", "Batch permanently redact matching text.", True, batchable=True),
    ToolDef("compare", "Compare PDF", "PDF Security", "⇄", "Create a side-by-side text-diff HTML report from exactly two PDFs.", True),

    ToolDef("summarize", "AI-style Summarizer", "PDF Intelligence", "✨", "Batch create local extractive summaries from PDF text.", True, batchable=True),
    ToolDef("translate", "Translate PDF", "PDF Intelligence", "🌍", "Smart batch translation with auto language detection and OCR fallback for scanned pages.", True, batchable=True),
    ToolDef("markdown", "PDF to Markdown", "PDF Intelligence", "M↓", "Batch convert PDF content to Markdown.", True, batchable=True),
    ToolDef("extract_text", "PDF Extract", "PDF Intelligence", "TXT", "Batch extract plain text and embedded images.", True, batchable=True),

    ToolDef("copy", "Copy / Rewrite PDF", "Workflows", "⧉", "Batch create clean duplicates with rewritten objects.", True, batchable=True),
    ToolDef("workflow", "Workflow Builder", "Workflows", "⚙", "Merge → optimize → optionally protect in one run.", True),
]

# Extra search words (synonyms + Greek) so the search box finds tools by everyday terms.
TOOL_KEYWORDS = {
    "merge": "combine join ενωση συνένωση συγχώνευση", "split": "separate cut divide διαχωρισμός διαίρεση",
    "remove": "delete pages διαγραφή σελίδων", "extract": "pages pick εξαγωγή σελίδων",
    "organize": "reorder sort arrange ταξινόμηση οργάνωση", "scan": "scanner photos σάρωση",
    "compress": "reduce shrink size smaller συμπίεση μέγεθος", "repair": "fix broken corrupt επισκευή",
    "ocr": "text recognition searchable scan αναγνώριση κειμένου", "img2pdf": "jpg png photo picture εικόνα",
    "word2pdf": "docx doc office", "ppt2pdf": "pptx slides office", "xls2pdf": "xlsx sheet office",
    "pdf2jpg": "image png picture export εικόνα", "pdf2word": "docx doc edit", "pdf2ppt": "pptx slides",
    "pdf2xls": "xlsx table sheet πίνακας", "pdfa": "archive archival αρχειοθέτηση",
    "rotate": "turn orientation περιστροφή", "numbers": "pagination footer αρίθμηση σελίδων",
    "watermark": "stamp draft confidential υδατογράφημα σφραγίδα", "crop": "trim margins περικοπή",
    "edit": "annotate write text κείμενο", "forms": "fields form φόρμα", "unlock": "decrypt remove password ξεκλείδωμα",
    "protect": "encrypt password secure κωδικός προστασία", "sign": "signature υπογραφή",
    "redact": "censor hide black μαύρισμα απόκρυψη", "compare": "diff difference σύγκριση",
    "summarize": "summary abstract περίληψη", "translate": "language greek english μετάφραση",
    "markdown": "md text", "extract_text": "txt images εξαγωγή κειμένου", "copy": "duplicate rewrite αντίγραφο",
    "workflow": "pipeline automation ροή",
}

CATEGORIES = [
    "All tools", "Organize PDF", "Optimize PDF", "Convert to PDF", "Convert from PDF",
    "Edit PDF", "PDF Security", "PDF Intelligence", "Workflows"
]

# ----------------------------- Utilities -----------------------------
def load_settings():
    defaults = {"theme": "dark", "last_dir": str(Path.home()), "recent": []}
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        defaults.update(data)
    except Exception:
        pass
    return defaults


def save_settings(data):
    """Atomic write so a crash mid-save can never corrupt settings.json."""
    try:
        tmp = SETTINGS_FILE.with_suffix(".json.tmp")
        tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, SETTINGS_FILE)
    except Exception:
        pass


def add_recent(settings: dict, tool: str, path: str):
    rec = [r for r in settings.get("recent", []) if r.get("path") != path]
    rec.insert(0, {"tool": tool, "path": path, "when": time.strftime("%Y-%m-%d %H:%M")})
    settings["recent"] = rec[:15]
    save_settings(settings)


def fmt_size(n: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if abs(n) < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} GB"


def fmt_elapsed(seconds: float) -> str:
    seconds = max(0, int(round(seconds)))
    m, s = divmod(seconds, 60)
    h, m = divmod(m, 60)
    return f"{h}h {m:02d}m {s:02d}s" if h else (f"{m}m {s:02d}s" if m else f"{s}s")


def filter_suffixes(qt_filter: str) -> set[str]:
    """Return lower-case suffixes from a Qt file filter, e.g. '*.pdf *.docx'."""
    return {"." + ext.lower() for ext in re.findall(r"\*\.([A-Za-z0-9]+)", qt_filter or "")}


def collect_folder_files(folder: str, qt_filter: str, recursive: bool = True) -> list[str]:
    """Collect matching files from a folder, recursively by default."""
    root = Path(folder)
    if not root.is_dir():
        return []
    suffixes = filter_suffixes(qt_filter)
    iterator = root.rglob("*") if recursive else root.glob("*")
    found = []
    for p in iterator:
        try:
            if not p.is_file():
                continue
            if suffixes and p.suffix.lower() not in suffixes:
                continue
            found.append(str(p))
        except OSError:
            continue
    return sorted(dict.fromkeys(found), key=lambda s: s.casefold())


def page_set(spec: str, page_count: int) -> list[int]:
    """Parse a 1-based page spec into zero-based unique pages, preserving order.

    Accepts: 1-3,5,8-6 (reverse ranges), 5- (to the end), -3 (from the start),
    'end' / 'last' as the last page, and ';' as an alternative separator.
    """
    spec = (spec or "").replace(" ", "").replace(";", ",").lower()
    if not spec:
        return list(range(page_count))

    def num(token: str, default: int) -> int:
        if token == "":
            return default
        if token in ("end", "last", "n"):
            return page_count
        if not token.isdigit():
            raise ValueError(f"'{token}' is not a page number. Use e.g. 1-3,5,8 or 10-end.")
        return int(token)

    out: list[int] = []
    seen: set[int] = set()
    for chunk in spec.split(","):
        if not chunk:
            continue
        if "-" in chunk:
            a_txt, b_txt = chunk.split("-", 1)
            a, b = num(a_txt, 1), num(b_txt, page_count)
            step = 1 if b >= a else -1
            vals = range(a, b + step, step)
        else:
            vals = [num(chunk, 1)]
        for p in vals:
            if p < 1 or p > page_count:
                raise ValueError(f"Page {p} is outside this document (1–{page_count}).")
            z = p - 1
            if z not in seen:
                seen.add(z)
                out.append(z)
    return out


def mm_to_pt(mm: float) -> float:
    return float(mm) * 72.0 / 25.4


def safe_output(input_path: str, suffix: str, ext: str = ".pdf") -> str:
    p = Path(input_path)
    return str(p.with_name(p.stem + suffix + ext))


def open_folder(path: str):
    p = Path(path)
    target = p if p.is_dir() else p.parent
    QDesktopServices.openUrl(QUrl.fromLocalFile(str(target.resolve())))


def _windows_install_candidates() -> list[str]:
    """Typical install locations for the optional external programs on Windows."""
    if os.name != "nt":
        return []
    roots = [os.environ.get("ProgramFiles", r"C:\Program Files"),
             os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"),
             os.environ.get("LOCALAPPDATA", "")]
    out = []
    for r in filter(None, roots):
        out.append(os.path.join(r, "LibreOffice", "program", "soffice.exe"))
        out.append(os.path.join(r, "Tesseract-OCR", "tesseract.exe"))
        out.append(os.path.join(r, "Programs", "Tesseract-OCR", "tesseract.exe"))
        # Any installed Ghostscript version, newest first.
        out.extend(sorted(glob.glob(os.path.join(r, "gs", "gs*", "bin", "gswin*c.exe")), reverse=True))
    return out


@functools.lru_cache(maxsize=32)
def _find_exe_cached(names: tuple) -> Optional[str]:
    for name in names:
        x = shutil.which(name)
        if x:
            return x
    # Only return a known install location whose file name is one of the requested names.
    wanted = {Path(n).name.lower() for n in names} | {Path(n).name.lower() + ".exe" for n in names}
    for c in _windows_install_candidates():
        if Path(c).name.lower() in wanted and Path(c).exists():
            return c
    return None


def find_exe(names) -> Optional[str]:
    return _find_exe_cached(tuple(names))


def find_ghostscript() -> Optional[str]:
    return find_exe(["gswin64c.exe", "gswin32c.exe", "gswin64c", "gswin32c", "gs"])


def find_tesseract() -> Optional[str]:
    return find_exe(["tesseract", "tesseract.exe"])


def find_soffice() -> Optional[str]:
    return find_exe(["soffice", "soffice.exe", "libreoffice"])


def open_pdf(path: str, opts: Optional[dict] = None):
    """Open a PDF and authenticate it when a password is supplied.

    Gives a clear, user-facing message instead of a MuPDF stack trace for
    damaged, non-PDF, or password-protected inputs.
    """
    try:
        doc = fitz.open(path)
    except Exception as exc:
        raise ValueError(f"{Path(path).name} could not be opened as a PDF ({exc}).")
    if doc.needs_pass:
        opts = opts or {}
        pw = (opts.get("passwords") or {}).get(path) or opts.get("input_password") or ""
        if not pw or not doc.authenticate(pw):
            doc.close()
            raise ValueError(f"{Path(path).name} is password-protected. Use Unlock PDF first "
                             "(or enter the correct password).")
    if not doc.is_pdf:
        # Images/XPS/EPUB opened by MuPDF are converted so PDF-only calls work.
        pdf_bytes = doc.convert_to_pdf(); doc.close()
        doc = fitz.open("pdf", pdf_bytes)
    return doc


def _same_file(a: str, b: str) -> bool:
    try:
        return os.path.normcase(str(Path(a).resolve())) == os.path.normcase(str(Path(b).resolve()))
    except Exception:
        return False


def save_pdf(doc, out: str, **kwargs):
    """Save atomically: write <out>.part first, then replace. A failed or cancelled
    job never leaves a truncated PDF behind, and an existing file is only replaced
    after a complete write."""
    if getattr(doc, "name", "") and _same_file(doc.name, out):
        raise ValueError("The output file must be different from the input file.")
    kwargs.setdefault("garbage", 3)
    kwargs.setdefault("deflate", True)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    tmp = str(out) + ".part"
    try:
        doc.save(tmp, **kwargs)
        os.replace(tmp, out)
    except BaseException:
        try:
            Path(tmp).unlink(missing_ok=True)
        except Exception:
            pass
        raise
    return out


def text_from_pdf(path: str) -> str:
    doc = open_pdf(path)
    try:
        return "\n\n".join(page.get_text("text") for page in doc)
    finally:
        doc.close()


# ---- text placement helpers (rotation-aware, Unicode/Greek-capable) ----
_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\segoeui.ttf", r"C:\Windows\Fonts\calibri.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/TTF/DejaVuSans.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/Library/Fonts/Arial Unicode.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf",
]


@functools.lru_cache(maxsize=1)
def _unicode_font_file() -> Optional[str]:
    windir = os.environ.get("WINDIR")
    cands = ([os.path.join(windir, "Fonts", n) for n in ("arial.ttf", "segoeui.ttf")] if windir else []) + _FONT_CANDIDATES
    for c in cands:
        if Path(c).exists():
            return c
    return None


def font_kwargs(text: str) -> dict:
    """Base-14 Helvetica covers Latin-1 only. For Greek/Cyrillic/etc. embed a system
    TrueType font (subset at save time) so the text does not turn into '?' marks."""
    if all(ord(ch) < 256 for ch in text or ""):
        return {"fontname": "helv"}
    ff = _unicode_font_file()
    if ff:
        return {"fontname": "PMUni", "fontfile": ff}
    return {"fontname": "helv", "encoding": fitz.TEXT_ENCODING_GREEK}


def text_width(text: str, fontsize: float, fk: dict) -> float:
    try:
        if "fontfile" in fk:
            return _cached_font(fk["fontfile"]).text_length(text, fontsize=fontsize)
        return fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
    except Exception:
        return len(text) * fontsize * 0.55


@functools.lru_cache(maxsize=4)
def _cached_font(fontfile: str):
    return fitz.Font(fontfile=fontfile)


def finish_fonts(doc, fk_used: bool):
    """Subset embedded TrueType fonts so a Greek watermark adds kB, not MB."""
    if fk_used:
        try:
            doc.subset_fonts()
        except Exception:
            pass


def to_page_space(page, visual_rect):
    """Convert a rectangle in what-the-user-sees coordinates (page.rect, rotation
    applied) into the unrotated coordinates PyMuPDF insert_* methods expect."""
    return fitz.Rect(visual_rect) * page.derotation_matrix


def to_page_point(page, visual_point):
    return fitz.Point(visual_point) * page.derotation_matrix


# ----------------------------- Dependency status -----------------------------
PYTHON_OPTIONAL_DEPS = [
    ("OCR", "pytesseract", "pytesseract"),
    ("PDF → Word", "pdf2docx", "pdf2docx"),
    ("PowerPoint output", "pptx", "python-pptx"),
    ("Excel output", "openpyxl", "openpyxl"),
    ("HTML → PDF", "weasyprint", "weasyprint"),
    ("PDF → Markdown", "pymupdf4llm", "pymupdf4llm"),
    ("Translate", "deep_translator", "deep-translator"),
]

EXTERNAL_OPTIONAL_DEPS = [
    ("LibreOffice", ["soffice", "soffice.exe", "libreoffice"]),
    ("Tesseract OCR", ["tesseract", "tesseract.exe"]),
    ("Ghostscript", ["gswin64c.exe", "gswin32c.exe", "gs"]),
]


def scan_optional_dependencies(progress):
    """Fast, non-importing dependency scan safe for a background QThread."""
    _find_exe_cached.cache_clear()
    importlib.invalidate_caches()
    result = {"python": [], "external": []}
    total = len(PYTHON_OPTIONAL_DEPS) + len(EXTERNAL_OPTIONAL_DEPS)
    done = 0

    for label, module_name, package_name in PYTHON_OPTIONAL_DEPS:
        try:
            found = importlib.util.find_spec(module_name) is not None
            detail = "module found" if found else "not installed"
        except Exception as exc:
            found = False
            detail = f"scan error: {exc}"
        result["python"].append({
            "label": label, "module": module_name, "package": package_name,
            "found": found, "detail": detail,
        })
        done += 1
        progress(int(done * 100 / total), f"Checking {label}…")

    for label, names in EXTERNAL_OPTIONAL_DEPS:
        try:
            path = find_exe(names)
        except Exception as exc:
            path = None
            detail = f"scan error: {exc}"
        else:
            detail = path or "not found"
        result["external"].append({"label": label, "found": bool(path), "detail": detail})
        done += 1
        progress(int(done * 100 / total), f"Checking {label}…")

    return result


# ----------------------------- Worker -----------------------------
class JobCancelled(BaseException):
    """BaseException on purpose: engine code that does `except Exception` (fallback
    paths, per-file error capture) must not swallow a user cancellation."""
    pass


class Worker(QThread):
    progress = pyqtSignal(int, str)
    finished_ok = pyqtSignal(object)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    # Only routine per-page/per-chunk counters are throttled; every other message
    # (warnings, decisions, results) always reaches the log.
    _ROUTINE = re.compile(r"^(Page|Slide|OCR page|Thumbnail|Writing page|Placed page|Added|Removed page|"
                          r"Translating chunk|Rendering|Rotating|Stamping|Numbering|Cropping|Redacting|Reading)\b")

    def __init__(self, fn: Callable, *args, **kwargs):
        super().__init__()
        self.fn = fn; self.args = args; self.kwargs = kwargs
        self._last_emit = 0.0
        self._last_value = -1

    def _progress(self, value, message=None):
        """Engine progress callback. Also the cancellation checkpoint.

        progress(None, None) is a pure cancellation check (no UI update).
        Emissions are throttled (~25/s) so 2 000-page jobs do not flood the GUI
        thread with signals and log lines.
        """
        if self.isInterruptionRequested():
            raise JobCancelled()
        if value is None and message is None:
            return
        value = self._last_value if value is None else int(value)
        message = "" if message is None else str(message)
        now = time.monotonic()
        important = value >= 100 or not self._ROUTINE.match(message)
        if important or now - self._last_emit >= 0.04 or value - self._last_value >= 5:
            self._last_emit = now; self._last_value = value
            self.progress.emit(value, message)

    def run(self):
        try:
            result = self.fn(self._progress, *self.args, **self.kwargs)
            if self.isInterruptionRequested():
                raise JobCancelled()
            self.finished_ok.emit(result)
        except JobCancelled:
            self.cancelled.emit()
        except Exception:
            self.failed.emit(traceback.format_exc())


def friendly_error(tb: str) -> str:
    """Last meaningful line of a traceback, without the exception class noise."""
    lines = [l for l in (tb or "").strip().splitlines() if l.strip()]
    if not lines:
        return "Unknown error"
    last = lines[-1]
    m = re.match(r"^[\w.]+(Error|Exception|Warning)?:\s*(.*)$", last)
    return (m.group(2) if m and m.group(2) else last).strip()


# ----------------------------- Engines -----------------------------
def eng_merge(progress, files, out, opts):
    if not files:
        raise ValueError("No PDFs to merge.")
    dst = fitz.open()
    try:
        for i, f in enumerate(files):
            src = open_pdf(f, opts)
            try:
                dst.insert_pdf(src)
            finally:
                src.close()
            progress(int((i + 1) * 90 / len(files)), f"Added {Path(f).name}")
        save_pdf(dst, out, garbage=4)
    finally:
        dst.close()
    progress(100, "Done")
    return out


def _split_one_pdf_job(args):
    """Process-safe split helper. Each worker opens its own PDF instance."""
    source, output_root, spec, each_page, batch_mode, batch_index = args
    src = None
    try:
        src = open_pdf(source)
        n = len(src)
        pages = page_set(spec, n) if spec else list(range(n))
        stem = _sanitize_stem(Path(source).stem)
        base_root = Path(output_root)
        out_dir = base_root / f"{batch_index:03d}_{stem}" if batch_mode else base_root
        out_dir.mkdir(parents=True, exist_ok=True)
        created = []
        if each_page:
            for pno in pages:
                d = fitz.open()
                try:
                    d.insert_pdf(src, from_page=pno, to_page=pno)
                    target = out_dir / f"{stem}_page_{pno+1:04d}.pdf"
                    save_pdf(d, str(target))
                finally:
                    d.close()
                created.append(str(target))
        else:
            ranges = [x for x in (spec or '').replace(' ', '').replace(';', ',').split(',') if x] or [f"1-{n}"]
            for i, chunk in enumerate(ranges):
                group = page_set(chunk, n)
                d = fitz.open()
                try:
                    for pno in group:
                        d.insert_pdf(src, from_page=pno, to_page=pno)
                    target = out_dir / f"{stem}_part_{i+1:03d}.pdf"
                    save_pdf(d, str(target))
                finally:
                    d.close()
                created.append(str(target))
        return source, str(out_dir), len(created), None
    except Exception as ex:
        return source, '', 0, f"{type(ex).__name__}: {ex}"
    finally:
        if src is not None:
            src.close()


# ----------------------------- Parallel job runner -----------------------------
def effective_workers(opts: dict, files: list, cap: int = 16) -> int:
    """Translate the performance profile + file sizes into a worker count."""
    cpu = os.cpu_count() or 2
    try:
        workers = int(opts.get("workers", min(4, cpu)))
    except (TypeError, ValueError):
        workers = min(4, cpu)
    workers = max(1, min(cap, workers))
    profile = str(opts.get("performance_profile", "Smart / Auto"))
    if profile.startswith("Smart"):
        try:
            sizes = [Path(f).stat().st_size for f in files]
            largest, total = max(sizes, default=0), sum(sizes)
            if largest >= 500 * 1024**2 or total >= 1500 * 1024**2:
                workers = 1
            elif largest >= 150 * 1024**2 or total >= 700 * 1024**2:
                workers = min(workers, 2)
            else:
                workers = min(workers, 4, max(1, cpu - 1))
        except Exception:
            workers = min(workers, 4)
    elif profile in ("Low memory", "Single worker"):
        workers = 1
    elif profile == "Balanced":
        workers = min(workers, 2)
    elif profile == "Maximum speed":
        workers = min(cap, max(workers, cpu))
    return max(1, min(workers, len(files) or 1))


_PART_RE = re.compile(r"\.(pdf|txt|md)\.part$|\.part\.(docx|pptx|xlsx)$", re.I)


def _cleanup_partials(output_root, since: float):
    """Remove temp files of worker processes that were terminated mid-save.

    Only our own temp names (*.pdf.part, *.part.docx …) inside the job's output
    folder that were created during this job are touched."""
    if not output_root:
        return
    time.sleep(0.2)  # let terminated processes release file handles (Windows)
    try:
        for f in Path(output_root).rglob("*part*"):
            try:
                if _PART_RE.search(f.name) and f.is_file() and f.stat().st_mtime >= since - 1:
                    f.unlink()
            except Exception:
                pass
    except Exception:
        pass


def run_parallel(progress, fn, jobs: list, max_workers: int, output_root: Optional[str] = None):
    """Yield (job, result, exception) as jobs finish in isolated spawn processes.

    The loop wakes every 0.3 s so a Cancel click is honoured within a fraction
    of a second: running worker processes are terminated instead of being waited
    for (atomic saves mean no half-written outputs are left behind).
    """
    started = time.time()
    ctx = multiprocessing.get_context("spawn")
    pool = ProcessPoolExecutor(max_workers=max_workers, mp_context=ctx)
    pending = {pool.submit(fn, j): j for j in jobs}
    finished = False
    try:
        while pending:
            done, _ = wait(list(pending), timeout=0.3, return_when=FIRST_COMPLETED)
            progress(None, None)  # cancellation checkpoint
            for fut in done:
                job = pending.pop(fut)
                try:
                    yield job, fut.result(), None
                except Exception as exc:  # e.g. BrokenProcessPool after a hard crash
                    yield job, None, exc
        finished = True
    except JobCancelled:
        _cleanup_partials(output_root, started)
        raise
    finally:
        if finished:
            pool.shutdown(wait=True)
        else:
            for p in list(getattr(pool, "_processes", {}).values()):
                try:
                    p.terminate()
                except Exception:
                    pass
            pool.shutdown(wait=False, cancel_futures=True)


def eng_split(progress, files, out, opts):
    """Batch splitter with no app-level input count limit."""
    out_dir = Path(out); out_dir.mkdir(parents=True, exist_ok=True)
    spec = opts.get("pages", "")
    each_page = bool(opts.get("each_page", True))
    batch_mode = len(files) > 1
    jobs = [(f, str(out_dir), spec, each_page, batch_mode, i+1) for i, f in enumerate(files)]
    errors, successes = [], 0
    workers = effective_workers(opts, files)

    def record(source, count, err, done):
        nonlocal successes
        if err:
            errors.append(f"{Path(source).name}: {err}")
            msg = f"⚠ {Path(source).name} failed — {err}"
        else:
            successes += 1
            msg = f"✓ {Path(source).name}: {count} output PDF(s)"
        progress(int(done * 100 / len(jobs)), msg)

    if len(jobs) == 1 or workers == 1:
        for i, job in enumerate(jobs):
            progress(int(i * 100 / len(jobs)), f"Splitting {Path(job[0]).name}…")
            source, _, count, err = _split_one_pdf_job(job)
            record(source, count, err, i + 1)
    else:
        progress(1, f"Batch split: {len(jobs)} PDFs using up to {workers} worker processes")
        done = 0
        for job, res, exc in run_parallel(progress, _split_one_pdf_job, jobs, workers, str(out_dir)):
            done += 1
            if exc is not None:
                record(job[0], 0, f"{type(exc).__name__}: {exc}", done)
            else:
                record(res[0], res[2], res[3], done)

    report = out_dir / "_PDFMule_split_report.txt"
    report.write_text(
        f"{APP_NAME} {APP_VERSION} batch split report\n\n"
        f"Inputs: {len(files)}\nSuccessful: {successes}\nFailed: {len(errors)}\n\n" +
        ("\n".join(errors) if errors else "No errors."), encoding="utf-8")
    if successes == 0:
        raise RuntimeError("All PDFs failed. See _PDFMule_split_report.txt in the output folder.\n" + "\n".join(errors[:8]))
    progress(100, f"Completed {successes}/{len(files)} PDF(s)" + (f" with {len(errors)} warning(s)" if errors else ""))
    return str(out_dir)


def eng_remove(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        spec = (opts.get("pages") or "").strip()
        if not spec:
            raise ValueError("Enter the pages to remove, e.g. 1,3-5 or 10-end.")
        remove = page_set(spec, len(doc))
        if len(remove) >= len(doc):
            raise ValueError(f"That would remove all {len(doc)} pages. Keep at least one page.")
        progress(30, f"Removing {len(remove)} page(s)")
        doc.delete_pages(sorted(remove))
        progress(80, f"Saving {len(doc)} remaining page(s)")
        save_pdf(doc, out, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def _copy_pages(progress, files, out, opts, verb):
    src = open_pdf(files[0], opts)
    dst = fitz.open()
    try:
        wanted = page_set(opts.get("pages", ""), len(src))
        if not wanted:
            raise ValueError("No pages selected.")
        if wanted == list(range(wanted[0], wanted[-1] + 1)):
            # Contiguous run: one insert call is much faster than page-by-page.
            dst.insert_pdf(src, from_page=wanted[0], to_page=wanted[-1])
        else:
            for i, p in enumerate(wanted):
                dst.insert_pdf(src, from_page=p, to_page=p)
                progress(int((i+1)*90/len(wanted)), f"{verb} page {p+1}")
        save_pdf(dst, out, garbage=4)
    finally:
        dst.close(); src.close()
    progress(100, "Done"); return out


def eng_extract(progress, files, out, opts):
    if not (opts.get("pages") or "").strip():
        raise ValueError("Enter the pages to extract, e.g. 1-3,5 or 10-end.")
    return _copy_pages(progress, files, out, opts, "Page")


def eng_organize(progress, files, out, opts):
    return _copy_pages(progress, files, out, opts, "Placed")


_PAPER_SIZES = {"A4": (595.28, 841.89), "Letter": (612.0, 792.0), "A3": (841.89, 1190.55)}


def _images_to_pdf(progress, files, out, opts):
    """Build one PDF from images, streaming one frame at a time.

    - every frame of multi-page TIFF/GIF is kept (the old code collected the
      ImageSequence iterator into a list, which yields N copies of the LAST frame);
    - camera EXIF orientation is applied, so phone photos are upright;
    - transparent PNGs are flattened on white instead of black;
    - untouched JPEGs are embedded byte-for-byte (fast, no generation loss).
    """
    import io
    from PIL import ImageSequence, ImageOps
    if not files:
        raise ValueError("No images selected.")
    page_mode = str(opts.get("page_size", "Image size (DPI)"))
    margin = mm_to_pt(float(opts.get("margin_mm", 0) or 0))
    doc = fitz.open(); total = len(files); pages_added = 0; skipped = []
    try:
        for i, f in enumerate(files):
            try:
                with Image.open(f) as im:
                    n_frames = getattr(im, "n_frames", 1)
                    try:
                        orientation = im.getexif().get(0x0112, 1)
                    except Exception:
                        orientation = 1
                    passthrough = (n_frames == 1 and im.format == "JPEG" and orientation in (1, None)
                                   and im.mode in ("RGB", "L", "CMYK"))
                    for frame in ImageSequence.Iterator(im):
                        progress(None, None)
                        dpi = frame.info.get("dpi", im.info.get("dpi", (150, 150)))
                        try:
                            dx = float(dpi[0] or 150); dy = float(dpi[1] or dx)
                        except Exception:
                            dx = dy = 150.0
                        if dx < 30 or dy < 30:
                            dx = dy = 150.0  # bogus DPI metadata would create giant pages
                        if passthrough:
                            data = Path(f).read_bytes(); w, h = frame.size
                        else:
                            fr = ImageOps.exif_transpose(frame) if n_frames == 1 else frame.copy()
                            if fr.mode in ("RGBA", "LA", "P", "PA"):
                                fr = fr.convert("RGBA")
                                bg = Image.new("RGB", fr.size, (255, 255, 255))
                                bg.paste(fr, mask=fr.split()[-1]); fr = bg
                            elif fr.mode not in ("RGB", "L", "1"):
                                fr = fr.convert("RGB")
                            bio = io.BytesIO()
                            if im.format in ("JPEG", "WEBP", "MPO") and fr.mode != "1":
                                fr.save(bio, "JPEG", quality=92)
                            else:
                                fr.save(bio, "PNG", compress_level=6)
                            data = bio.getvalue(); w, h = fr.size
                            if fr is not frame:
                                fr.close()
                        img_w, img_h = w * 72.0 / dx, h * 72.0 / dy
                        if page_mode.startswith(("A4", "Letter", "A3")):
                            pw, ph = _PAPER_SIZES[page_mode.split()[0]]
                            if img_w > img_h:
                                pw, ph = ph, pw  # landscape paper for landscape images
                        else:
                            pw, ph = img_w + 2 * margin, img_h + 2 * margin
                        page = doc.new_page(width=max(1.0, pw), height=max(1.0, ph))
                        box = fitz.Rect(margin, margin, pw - margin, ph - margin)
                        page.insert_image(box, stream=data, keep_proportion=True)
                        pages_added += 1
            except JobCancelled:
                raise
            except Exception as exc:
                skipped.append(f"{Path(f).name}: {exc}")
                progress(int((i+1)*90/total), f"⚠ Skipped {Path(f).name} — {exc}")
                continue
            progress(int((i+1)*90/total), f"Added {Path(f).name}")
        if pages_added == 0:
            raise ValueError("No image pages could be read.\n" + "\n".join(skipped[:5]))
        save_pdf(doc, out, garbage=3)
    finally:
        doc.close()
    progress(100, f"Done — {pages_added} page(s)" + (f", {len(skipped)} file(s) skipped" if skipped else ""))
    return out


def _pdf_image_ratio(doc, sample: int = 6) -> float:
    """Rough share of pages whose content is mainly raster images (scans/photos)."""
    n = len(doc)
    if not n:
        return 0.0
    idx = sorted({int(i * (n - 1) / max(1, sample - 1)) for i in range(min(n, sample))})
    hits = 0
    for i in idx:
        try:
            p = doc[i]
            if p.get_images(full=False) and len((p.get_text("text") or "").strip()) < 200:
                hits += 1
        except Exception:
            pass
    return hits / max(1, len(idx))


def eng_compress(progress, files, out, opts):
    """Try the appropriate strategies, then keep the SMALLEST result.

    Never returns a file larger than the original: if nothing helps, the output
    is a byte-identical copy and the log says the PDF is already optimised.
    """
    source = files[0]
    if _same_file(source, out):
        raise ValueError("The output file must be different from the input file.")
    profile = str(opts.get("compression_profile", "Smart / Auto"))
    orig_size = Path(source).stat().st_size
    gs = find_ghostscript()
    has_rewrite = hasattr(fitz.Document, "rewrite_images")
    candidates = []  # (size, path, label)

    with tempfile.TemporaryDirectory(prefix="pdfmule_cmp_") as td:
        def lossless():
            progress(20, "Compression: lossless object/stream optimisation")
            p = os.path.join(td, "lossless.pdf")
            d = open_pdf(source, opts)
            try:
                d.save(p, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True,
                       clean=True, use_objstms=1)
            finally:
                d.close()
            candidates.append((Path(p).stat().st_size, p, "lossless"))

        def pymupdf_images(dpi_threshold, dpi_target, quality, label):
            if not has_rewrite:
                return
            progress(45, f"Compression: recompressing images ({label}: >{dpi_threshold} dpi → {dpi_target} dpi, JPEG q{quality})")
            p = os.path.join(td, f"img_{dpi_target}_{quality}.pdf")
            d = open_pdf(source, opts)
            try:
                d.rewrite_images(dpi_threshold=dpi_threshold, dpi_target=dpi_target, quality=quality)
                d.save(p, garbage=4, deflate=True, clean=True, use_objstms=1)
            finally:
                d.close()
            candidates.append((Path(p).stat().st_size, p, label))

        def ghostscript(preset):
            if not gs:
                return
            progress(60, f"Compression: Ghostscript {preset}")
            p = os.path.join(td, f"gs_{preset.strip('/')}.pdf")
            cp = run_quiet([gs, "-dBATCH", "-dNOPAUSE", "-dQUIET", "-dSAFER", "-sDEVICE=pdfwrite",
                            f"-dPDFSETTINGS={preset}", "-dCompatibilityLevel=1.6",
                            f"-sOutputFile={p}", source], timeout=900)
            if cp.returncode == 0 and Path(p).exists() and Path(p).stat().st_size > 0:
                candidates.append((Path(p).stat().st_size, p, f"Ghostscript {preset}"))
            else:
                progress(62, "⚠ Ghostscript failed; using PyMuPDF results")

        lossless()
        progress(None, None)
        if profile.startswith("Balanced"):
            pymupdf_images(200, 150, 75, "balanced")
            if not has_rewrite:
                ghostscript("/ebook")
        elif profile.startswith("Strong"):
            pymupdf_images(110, 96, 55, "strong")
            ghostscript("/screen")
        elif profile.startswith("Smart"):
            doc = open_pdf(source, opts)
            try:
                img_ratio = _pdf_image_ratio(doc)
            finally:
                doc.close()
            best_lossless = candidates[0][0]
            # Scans/photos barely shrink losslessly; recompress mildly when it is worth it.
            if img_ratio >= 0.5 and best_lossless > 0.9 * orig_size and orig_size > 1024**2:
                pymupdf_images(220, 200, 80, "smart (scanned pages)")
            elif orig_size >= 25 * 1024**2 and gs:
                ghostscript("/ebook")
        # "Lossless" profile: only the lossless candidate.

        progress(90, "Choosing the smallest result")
        candidates.sort()
        if not candidates or candidates[0][0] >= orig_size:
            shutil.copy2(source, out + ".part"); os.replace(out + ".part", out)
            progress(100, f"Done — the PDF is already optimised ({fmt_size(orig_size)}); kept original content.")
            return out
        size, path, label = candidates[0]
        shutil.copyfile(path, out + ".part"); os.replace(out + ".part", out)
    saved = 100.0 * (orig_size - size) / max(1, orig_size)
    progress(100, f"Done — {fmt_size(orig_size)} → {fmt_size(size)} (−{saved:.0f}%, {label})")
    return out


def eng_repair(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        progress(35, "Rebuilding xref and cleaning content streams")
        save_pdf(doc, out, garbage=4, clean=True)
    finally:
        doc.close()
    progress(100, "Done"); return out


# Standard Tesseract language packs. The selector exposes the full standard set;
# the OCR engine validates that the selected traineddata packs are installed.
TESSERACT_LANGUAGE_NAMES = {
    "afr":"Afrikaans", "amh":"Amharic", "ara":"Arabic", "asm":"Assamese", "aze":"Azerbaijani",
    "aze_cyrl":"Azerbaijani (Cyrillic)", "bel":"Belarusian", "ben":"Bengali", "bod":"Tibetan", "bos":"Bosnian",
    "bre":"Breton", "bul":"Bulgarian", "cat":"Catalan / Valencian", "ceb":"Cebuano", "ces":"Czech",
    "chi_sim":"Chinese - Simplified", "chi_sim_vert":"Chinese - Simplified vertical", "chi_tra":"Chinese - Traditional",
    "chi_tra_vert":"Chinese - Traditional vertical", "chr":"Cherokee", "cos":"Corsican", "cym":"Welsh", "dan":"Danish",
    "deu":"German", "div":"Divehi", "dzo":"Dzongkha", "ell":"Greek", "eng":"English", "enm":"English (Middle)",
    "epo":"Esperanto", "est":"Estonian", "eus":"Basque", "fao":"Faroese", "fas":"Persian", "fil":"Filipino",
    "fin":"Finnish", "fra":"French", "frk":"German (Fraktur)", "frm":"French (Middle)", "fry":"Frisian",
    "gla":"Scottish Gaelic", "gle":"Irish", "glg":"Galician", "grc":"Greek (Ancient)", "guj":"Gujarati",
    "hat":"Haitian / Haitian Creole", "heb":"Hebrew", "hin":"Hindi", "hrv":"Croatian", "hun":"Hungarian",
    "hye":"Armenian", "iku":"Inuktitut", "ind":"Indonesian", "isl":"Icelandic", "ita":"Italian",
    "ita_old":"Italian (Old)", "jav":"Javanese", "jpn":"Japanese", "jpn_vert":"Japanese vertical", "kan":"Kannada",
    "kat":"Georgian", "kat_old":"Georgian (Old)", "kaz":"Kazakh", "khm":"Khmer", "kir":"Kyrgyz", "kmr":"Kurdish (Kurmanji)",
    "kor":"Korean", "kor_vert":"Korean vertical", "lao":"Lao", "lat":"Latin", "lav":"Latvian", "lit":"Lithuanian",
    "ltz":"Luxembourgish", "mal":"Malayalam", "mar":"Marathi", "mkd":"Macedonian", "mlt":"Maltese", "mon":"Mongolian",
    "mri":"Maori", "msa":"Malay", "mya":"Burmese", "nep":"Nepali", "nld":"Dutch / Flemish", "nor":"Norwegian",
    "oci":"Occitan", "ori":"Odia / Oriya", "pan":"Punjabi", "pol":"Polish", "por":"Portuguese", "pus":"Pashto",
    "que":"Quechua", "ron":"Romanian / Moldavian", "rus":"Russian", "san":"Sanskrit", "sin":"Sinhala", "slk":"Slovak",
    "slv":"Slovenian", "snd":"Sindhi", "spa":"Spanish / Castilian", "spa_old":"Spanish (Old)", "sqi":"Albanian",
    "srp":"Serbian (Cyrillic)", "srp_latn":"Serbian (Latin)", "sun":"Sundanese", "swa":"Swahili", "swe":"Swedish",
    "syr":"Syriac", "tam":"Tamil", "tat":"Tatar", "tel":"Telugu", "tgk":"Tajik", "tha":"Thai", "tir":"Tigrinya",
    "ton":"Tonga", "tur":"Turkish", "uig":"Uyghur", "ukr":"Ukrainian", "urd":"Urdu", "uzb":"Uzbek",
    "uzb_cyrl":"Uzbek (Cyrillic)", "vie":"Vietnamese", "yid":"Yiddish", "yor":"Yoruba"
}


def _tesseract_installed_languages(tess: str) -> set[str]:
    """Return installed Tesseract traineddata language codes without importing heavy modules."""
    try:
        kwargs = {"capture_output": True, "text": True, "timeout": 8}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        cp = subprocess.run([tess, "--list-langs"], **kwargs)
        raw = (cp.stdout or "") + "\n" + (cp.stderr or "")
        langs = set()
        for line in raw.splitlines():
            s = line.strip()
            if not s or s.lower().startswith("list of available languages"):
                continue
            # Tesseract may print absolute tessdata paths in diagnostics; only keep language-like ids.
            if re.fullmatch(r"[A-Za-z0-9_]+", s):
                langs.add(s)
        return langs
    except Exception:
        return set()


def _unicode_script_counts(text: str) -> dict[str, int]:
    counts = {k: 0 for k in ("Greek", "Latin", "Cyrillic", "Arabic", "Hebrew", "Devanagari", "Thai", "Hangul", "Japanese", "Han", "Other")}
    for ch in text:
        if not ch.isalpha():
            continue
        cp = ord(ch)
        if 0x0370 <= cp <= 0x03FF or 0x1F00 <= cp <= 0x1FFF:
            counts["Greek"] += 1
        elif (0x0041 <= cp <= 0x007A) or (0x00C0 <= cp <= 0x024F) or (0x1E00 <= cp <= 0x1EFF):
            counts["Latin"] += 1
        elif 0x0400 <= cp <= 0x052F:
            counts["Cyrillic"] += 1
        elif (0x0600 <= cp <= 0x06FF) or (0x0750 <= cp <= 0x077F) or (0x08A0 <= cp <= 0x08FF):
            counts["Arabic"] += 1
        elif 0x0590 <= cp <= 0x05FF:
            counts["Hebrew"] += 1
        elif 0x0900 <= cp <= 0x097F:
            counts["Devanagari"] += 1
        elif 0x0E00 <= cp <= 0x0E7F:
            counts["Thai"] += 1
        elif 0xAC00 <= cp <= 0xD7AF:
            counts["Hangul"] += 1
        elif (0x3040 <= cp <= 0x30FF):
            counts["Japanese"] += 1
        elif 0x4E00 <= cp <= 0x9FFF:
            counts["Han"] += 1
        else:
            counts["Other"] += 1
    return counts


_LATIN_LANGUAGE_HINTS = {
    "eng": (" the ", " and ", " of ", " to ", " in ", " for ", " with ", " is ", " are "),
    "deu": (" der ", " die ", " das ", " und ", " ist ", " mit ", " für ", " von ", " nicht "),
    "fra": (" le ", " la ", " les ", " de ", " des ", " et ", " est ", " pour ", " avec "),
    "spa": (" el ", " la ", " los ", " las ", " de ", " y ", " para ", " con ", " que "),
    "ita": (" il ", " lo ", " la ", " gli ", " di ", " e ", " per ", " con ", " che "),
    "por": (" o ", " a ", " os ", " as ", " de ", " e ", " para ", " com ", " que "),
    "nld": (" de ", " het ", " een ", " en ", " van ", " voor ", " met ", " niet "),
}


def _guess_latin_language(text: str, installed: set[str]) -> str | None:
    hay = " " + re.sub(r"\s+", " ", text.lower()) + " "
    best_code, best_score = None, 0
    for code, words in _LATIN_LANGUAGE_HINTS.items():
        if installed and code not in installed:
            continue
        score = sum(hay.count(w) for w in words)
        if score > best_score:
            best_code, best_score = code, score
    return best_code if best_score >= 3 else None


def _script_to_tess_lang(script: str, installed: set[str]) -> str | None:
    # Greek is intentionally unambiguous: this is the main Smart OCR use case.
    candidates = {
        "Greek": ["ell"],
        "Arabic": ["ara"],
        "Hebrew": ["heb"],
        "Devanagari": ["hin"],
        "Thai": ["tha"],
        "Hangul": ["kor"],
        "Japanese": ["jpn"],
        "Han": ["chi_sim", "chi_tra"],
        "Cyrillic": ["rus", "ukr", "bul", "srp", "mkd", "bel"],
        "Latin": ["eng"],
    }.get(script, [])
    if installed:
        for c in candidates:
            if c in installed:
                return c
        return None
    return candidates[0] if candidates else None


def _smart_ocr_language(src, pytesseract, installed: set[str], threshold: float, progress) -> tuple[str, str]:
    """Detect the document's dominant script/language. Threshold is a ratio, e.g. .70."""
    progress(2, "Smart OCR: analyzing document language…")
    chunks = []
    for i in range(min(len(src), 8)):
        try:
            t = src[i].get_text("text") or ""
            if t.strip():
                chunks.append(t[:12000])
        except Exception:
            pass
    text = "\n".join(chunks)
    counts = _unicode_script_counts(text)
    total = sum(counts.values())
    if total >= 40:
        dominant = max((k for k in counts if k != "Other"), key=lambda k: counts[k], default="Other")
        ratio = counts.get(dominant, 0) / max(1, total)
        # Explicit 70%-style dominant-script rule requested by the app design.
        if ratio >= threshold:
            if dominant == "Latin":
                code = _guess_latin_language(text, installed) or _script_to_tess_lang("Latin", installed)
            else:
                code = _script_to_tess_lang(dominant, installed)
            if code:
                return code, f"embedded text: {dominant} {ratio*100:.0f}% → {TESSERACT_LANGUAGE_NAMES.get(code, code)} ({code})"

    # Scanned/no-text fallback: Tesseract OSD can identify the script from rendered sample pages.
    progress(4, "Smart OCR: no decisive embedded text; sampling scanned pages…")
    osd_scripts = []
    for i in range(min(len(src), 3)):
        try:
            pix = src[i].get_pixmap(dpi=110, alpha=False)
            mode = "RGB" if pix.n >= 3 else "L"
            img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
            data = pytesseract.image_to_osd(img, output_type=pytesseract.Output.DICT)
            script = str(data.get("script", "")).strip()
            conf = float(data.get("script_conf", 0) or 0)
            if script and conf >= 1.0:
                osd_scripts.append((script, conf))
        except Exception:
            continue
    if osd_scripts:
        scores = {}
        for script, conf in osd_scripts:
            scores[script] = scores.get(script, 0.0) + conf
        script = max(scores, key=scores.get)
        aliases = {
            "Greek":"Greek", "Latin":"Latin", "Cyrillic":"Cyrillic", "Arabic":"Arabic", "Hebrew":"Hebrew",
            "Devanagari":"Devanagari", "Thai":"Thai", "Hangul":"Hangul", "Korean":"Hangul",
            "Japanese":"Japanese", "Han":"Han", "HanS":"Han", "HanT":"Han"
        }
        normalized = aliases.get(script, script)
        code = _script_to_tess_lang(normalized, installed)
        if code:
            return code, f"scan sample / OSD: {script} → {TESSERACT_LANGUAGE_NAMES.get(code, code)} ({code})"

    # Safe fallback. Prefer English if present, otherwise any installed OCR language except OSD.
    if not installed or "eng" in installed:
        return "eng", "fallback → English (eng)"
    usable = sorted(x for x in installed if x != "osd")
    if usable:
        code = usable[0]
        return code, f"fallback → {TESSERACT_LANGUAGE_NAMES.get(code, code)} ({code})"
    return "eng", "fallback → English (eng)"

def _pix_to_pil(pix):
    mode = "RGB" if pix.n >= 3 else "L"
    if pix.alpha:
        return Image.frombytes("RGBA" if pix.n == 4 else "LA", (pix.width, pix.height), pix.samples).convert(mode)
    return Image.frombytes(mode, (pix.width, pix.height), pix.samples)


def _setup_tesseract():
    try:
        import pytesseract
    except Exception:
        raise RuntimeError("OCR needs the pytesseract package. Install it from Dependency status.")
    tess = find_tesseract()
    if not tess:
        raise RuntimeError("Tesseract OCR was not found. Install Tesseract (Windows: UB-Mannheim build) "
                           "and reopen PDF Mule.")
    pytesseract.pytesseract.tesseract_cmd = tess
    return pytesseract, tess


def _page_has_text(page, min_alpha: int = 30) -> bool:
    try:
        txt = page.get_text("text") or ""
    except Exception:
        return False
    return sum(1 for ch in txt if ch.isalpha()) >= min_alpha


def _auto_upright(page, pytesseract) -> int:
    """Use Tesseract OSD to fix upside-down / sideways scans. Only acts when OSD is
    confident; returns the rotation that was added (0 if none)."""
    try:
        pix = page.get_pixmap(dpi=150, alpha=False)
        osd = pytesseract.image_to_osd(_pix_to_pil(pix), output_type=pytesseract.Output.DICT)
        add = int(osd.get("rotate", 0)) % 360
        conf = float(osd.get("orientation_conf", 0) or 0)
        if add and conf >= 1.0:
            page.set_rotation((page.rotation + add) % 360)
            return add
    except Exception:
        pass
    return 0


def eng_ocr(progress, files, out, opts):
    """Searchable-PDF OCR that keeps the ORIGINAL page images.

    Tesseract produces an invisible text-only layer which is overlaid on each
    page, so the output keeps the scan quality and roughly the input size (the
    old approach replaced every page by a re-encoded Tesseract image).
    """
    pytesseract, tess = _setup_tesseract()
    installed = _tesseract_installed_languages(tess)
    mode = str(opts.get("ocr_mode", "Smart / Auto detect"))
    dpi = int(opts.get("dpi", 300))
    threshold = max(0.50, min(0.95, float(opts.get("smart_threshold", 70)) / 100.0))
    skip_text = bool(opts.get("skip_text_pages", True))
    auto_rotate = bool(opts.get("auto_rotate", True))
    src = open_pdf(files[0], opts)
    try:
        if mode.startswith("Smart"):
            lang, why = _smart_ocr_language(src, pytesseract, installed, threshold, progress)
            if installed and lang not in installed:
                raise RuntimeError(
                    f"Smart OCR detected {TESSERACT_LANGUAGE_NAMES.get(lang, lang)} ({lang}), "
                    f"but {lang}.traineddata is not installed in Tesseract.")
            progress(6, f"Smart OCR language: {why}")
        else:
            requested = [x.strip() for x in str(opts.get("lang", "eng")).split("+") if x.strip()]
            if not requested:
                raise ValueError("Select at least one OCR language in Manual multi-select mode.")
            missing = [x for x in requested if installed and x not in installed]
            if missing:
                readable = ", ".join(f"{TESSERACT_LANGUAGE_NAMES.get(x,x)} ({x})" for x in missing)
                raise RuntimeError("These selected Tesseract language packs are not installed: " + readable +
                                   ". Install the corresponding .traineddata files and try again.")
            lang = "+".join(requested)
            progress(6, "Manual OCR languages: " + ", ".join(TESSERACT_LANGUAGE_NAMES.get(x, x) for x in requested))
        if auto_rotate and installed and "osd" not in installed:
            auto_rotate = False
            progress(7, "⚠ osd.traineddata not installed — automatic page straightening disabled")

        n = len(src); done_ocr = skipped = rotated = 0
        for i in range(n):
            page = src[i]
            if skip_text and _page_has_text(page):
                skipped += 1
                progress(6 + int((i+1)*88/max(1, n)), f"Page {i+1}/{n} already has text — kept as is")
                continue
            if auto_rotate and _auto_upright(page, pytesseract):
                rotated += 1
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            pdf_bytes = pytesseract.image_to_pdf_or_hocr(_pix_to_pil(pix), extension="pdf", lang=lang,
                                                         config="-c textonly_pdf=1")
            del pix
            layer = fitz.open("pdf", pdf_bytes)
            try:
                # Text layer is in visual orientation; rotate it back into page space.
                page.show_pdf_page(page.rect * page.derotation_matrix, layer, 0,
                                   overlay=True, rotate=page.rotation)
            finally:
                layer.close()
            done_ocr += 1
            progress(6 + int((i+1)*88/max(1, n)), f"OCR page {i+1}/{n} • {lang}")
        progress(95, f"OCR'd {done_ocr} page(s)" + (f", kept {skipped} text page(s)" if skipped else "")
                 + (f", straightened {rotated} page(s)" if rotated else ""))
        save_pdf(src, out, garbage=4)
    finally:
        src.close()
    progress(100, "Done"); return out


def eng_office2pdf(progress, files, out, opts):
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice was not found. Install LibreOffice and try again.")
    out_dir = Path(out); out_dir.mkdir(parents=True, exist_ok=True)
    made = []
    for i, f in enumerate(files):
        progress(int(i*100/max(1, len(files))), f"Converting {Path(f).name}…")
        # A private LibreOffice profile makes concurrent batch conversions safe and
        # avoids clashing with a LibreOffice window the user already has open.
        with tempfile.TemporaryDirectory(prefix="pdfmule_lo_") as profile_dir:
            profile_uri = Path(profile_dir).resolve().as_uri()
            cmd = [soffice, f"-env:UserInstallation={profile_uri}", "--headless", "--norestore",
                   "--convert-to", "pdf", "--outdir", str(out_dir), f]
            cp = run_quiet(cmd, timeout=300)
        target = out_dir / (Path(f).stem + ".pdf")
        if cp.returncode != 0 or not target.exists():
            msg = (cp.stderr or cp.stdout or "").strip()
            raise RuntimeError(f"LibreOffice could not convert {Path(f).name}." + (f"\n{msg[-400:]}" if msg else ""))
        made.append(str(target))
        progress(int((i+1)*100/max(1, len(files))), f"✓ {Path(f).name}")
    return str(out_dir) if len(made) != 1 else made[0]


def eng_html2pdf(progress, files, out, opts):
    try:
        from weasyprint import HTML
    except Exception:
        raise RuntimeError("HTML to PDF needs WeasyPrint. Install it from Dependency status "
                           "(on Windows WeasyPrint also needs the GTK/Pango runtime).")
    source = (opts.get("source") or "").strip().strip('"')
    if not source:
        raise ValueError("Enter a URL or local HTML file path.")
    progress(20, "Loading HTML")
    if re.match(r"^[a-z]+://", source, re.I):
        html = HTML(url=source)
    else:
        if not Path(source).exists():
            raise ValueError(f"File not found: {source}")
        html = HTML(filename=source)
    tmp = out + ".part"
    try:
        html.write_pdf(tmp); os.replace(tmp, out)
    finally:
        Path(tmp).unlink(missing_ok=True)
    progress(100, "Done"); return out


def eng_pdf2jpg(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    out_dir = Path(out); out_dir.mkdir(parents=True, exist_ok=True)
    dpi = int(opts.get("dpi", 180)); quality = int(opts.get("quality", 90))
    stem = _sanitize_stem(Path(files[0]).stem)
    try:
        n = len(doc); digits = max(3, len(str(n)))
        for i, p in enumerate(doc):
            pix = p.get_pixmap(dpi=dpi, alpha=False)
            im = _pix_to_pil(pix); del pix   # direct samples → PIL, no PNG round-trip
            im.save(out_dir / f"{stem}_{i+1:0{digits}d}.jpg", quality=quality, optimize=True,
                    dpi=(dpi, dpi))
            im.close(); progress(int((i+1)*100/max(1, n)), f"Page {i+1}/{n}")
    finally:
        doc.close()
    return str(out_dir)


def eng_pdf2word(progress, files, out, opts):
    try:
        from pdf2docx import Converter
    except Exception:
        raise RuntimeError("PDF to Word needs pdf2docx. Install optional packages from Dependency status.")
    progress(10, "Reading PDF (layout analysis can take a while for long documents)")
    tmp = out + ".part.docx"
    cv = Converter(files[0], password=opts.get("input_password") or None)
    try:
        cv.convert(tmp)
    finally:
        cv.close()
    os.replace(tmp, out)
    progress(100, "Done"); return out


def eng_pdf2ppt(progress, files, out, opts):
    try:
        from pptx import Presentation
        from pptx.util import Emu
    except Exception:
        raise RuntimeError("PDF to PowerPoint needs python-pptx. Install optional packages from Dependency status.")
    import io
    doc = open_pdf(files[0], opts)
    try:
        if not len(doc):
            raise ValueError("The PDF has no pages.")
        prs = Presentation()
        # Slide size follows the first page's aspect ratio (A4, Letter, 16:9 …), max 13.333 in wide.
        r0 = doc[0].rect
        width_in = 13.333 if r0.width >= r0.height else 7.5
        prs.slide_width = Emu(int(width_in * 914400))
        prs.slide_height = Emu(int(width_in * 914400 * r0.height / max(1.0, r0.width)))
        blank = prs.slide_layouts[6]
        sw, sh = prs.slide_width, prs.slide_height
        for i, page in enumerate(doc):
            slide = prs.slides.add_slide(blank)
            pix = page.get_pixmap(dpi=150, alpha=False)
            stream = io.BytesIO(pix.tobytes("jpeg", jpg_quality=88) if hasattr(pix, "tobytes") else pix.tobytes("png"))
            # Fit the page image inside the slide, centred, preserving aspect ratio.
            scale = min(sw / pix.width, sh / pix.height)
            w, h = int(pix.width * scale), int(pix.height * scale)
            slide.shapes.add_picture(stream, int((sw - w) / 2), int((sh - h) / 2), width=w, height=h)
            progress(int((i+1)*95/len(doc)), f"Slide {i+1}/{len(doc)}")
        tmp = out + ".part.pptx"; prs.save(tmp); os.replace(tmp, out)
    finally:
        doc.close()
    progress(100, "Done"); return out


_XLSX_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _xlsx_cell(val):
    if val is None:
        return None
    s = _XLSX_ILLEGAL.sub("", str(val))
    if re.fullmatch(r"-?\d{1,15}", s.strip()):
        try:
            return int(s.strip())
        except ValueError:
            pass
    if re.fullmatch(r"-?\d{1,15}[.,]\d+", s.strip()):
        try:
            return float(s.strip().replace(",", "."))
        except ValueError:
            pass
    return s[:32000]


def eng_pdf2xls(progress, files, out, opts):
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font
    except Exception:
        raise RuntimeError("PDF to Excel needs openpyxl. Install optional packages from Dependency status.")
    doc = open_pdf(files[0], opts)
    wb = Workbook(); wb.remove(wb.active)
    tables_found = 0
    try:
        for i, page in enumerate(doc):
            ws = wb.create_sheet(f"Page {i+1}")
            row = 0
            try:
                tabs = page.find_tables().tables
            except Exception:
                tabs = []
            for t_idx, table in enumerate(tabs):
                try:
                    data = table.extract()
                except Exception:
                    continue
                tables_found += 1
                if t_idx:
                    row += 2
                for r_i, r in enumerate(data):
                    row += 1
                    for c, val in enumerate(r, 1):
                        cell = ws.cell(row, c, _xlsx_cell(val))
                        if r_i == 0:
                            cell.font = Font(bold=True)
            if row == 0:
                # No table detected: keep the text so the sheet is never silently empty.
                ws.cell(1, 1, "No table detected on this page — plain text follows:").font = Font(italic=True)
                for line in (page.get_text("text") or "").splitlines():
                    ws.cell(ws.max_row + 1, 1, _xlsx_cell(line))
            progress(int((i+1)*95/max(1, len(doc))), f"Page {i+1}/{len(doc)}")
        tmp = out + ".part.xlsx"; wb.save(tmp); os.replace(tmp, out)
    finally:
        doc.close()
    progress(100, f"Done — {tables_found} table(s) found"); return out


def eng_pdfa(progress, files, out, opts):
    gs = find_ghostscript()
    if not gs:
        raise RuntimeError("PDF/A conversion needs Ghostscript. Install Ghostscript and ensure gs/gswin64c is on PATH.")
    level = str(opts.get("pdfa", "2")).strip()[:1] or "2"
    tmp = out + ".part"
    cmd = [gs, f"-dPDFA={level}", "-dBATCH", "-dNOPAUSE", "-dQUIET", "-dNOOUTERSAVE", "-sDEVICE=pdfwrite",
           "-dPDFACompatibilityPolicy=1", "-sColorConversionStrategy=RGB",
           f"-sOutputFile={tmp}", files[0]]
    progress(20, f"Ghostscript PDF/A-{level} conversion")
    cp = run_quiet(cmd, timeout=900)
    if cp.returncode != 0 or not Path(tmp).exists():
        Path(tmp).unlink(missing_ok=True)
        raise RuntimeError((cp.stderr or cp.stdout or "Ghostscript failed.").strip()[-600:])
    os.replace(tmp, out)
    progress(100, "Done"); return out


def eng_rotate(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        angle = int(str(opts.get("angle", 90)).split()[0].rstrip("°"))
        spec = (opts.get("pages") or "").strip()
        targets = page_set(spec, len(doc)) if spec else range(len(doc))
        for i, pno in enumerate(targets):
            p = doc[pno]; p.set_rotation((p.rotation + angle) % 360)
            progress(int((i+1)*90/max(1, len(doc))), f"Rotating page {pno+1}")
        save_pdf(doc, out, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def _pos_rect(page_rect, w, h, position, margin=24):
    """Rectangle of size w×h inside the VISUAL page rect at a named position."""
    pr = page_rect
    pos = position.lower()
    x = pr.x0 + margin if "left" in pos else pr.x1 - w - margin if "right" in pos else pr.x0 + (pr.width - w) / 2
    y = pr.y0 + margin if "top" in pos else pr.y1 - h - margin if "bottom" in pos else pr.y0 + (pr.height - h) / 2
    return fitz.Rect(x, y, x + w, y + h)


def eng_numbers(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        start = int(opts.get("start", 1)); size = float(opts.get("font_size", 11))
        position = str(opts.get("position", "Bottom center"))
        fmt = str(opts.get("number_format", "1"))
        skip_first = bool(opts.get("skip_first", False))
        n = len(doc); used_uni = False
        for i, p in enumerate(doc):
            if skip_first and i == 0:
                continue
            num = start + i - (1 if skip_first else 0)
            total = n - (1 if skip_first else 0)
            txt = {"1": str(num), "Page 1": f"Page {num}", "1 / N": f"{num} / {total + start - 1}",
                   "Page 1 of N": f"Page {num} of {total + start - 1}", "Σελίδα 1 από N": f"Σελίδα {num} από {total + start - 1}"
                   }.get(fmt, str(num))
            fk = font_kwargs(txt); used_uni |= "fontfile" in fk
            w = text_width(txt, size, fk) + 4; h = size * 1.6
            vis = _pos_rect(p.rect, w, h, position, margin=max(18, size * 1.5))
            p.insert_textbox(to_page_space(p, vis), txt, fontsize=size, align=1,
                             color=(0.15, 0.15, 0.15), rotate=p.rotation, **fk)
            progress(int((i+1)*90/max(1, n)), f"Numbering page {i+1}")
        finish_fonts(doc, used_uni)
        save_pdf(doc, out, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def _faded_image_bytes(path: str, opacity: float) -> bytes:
    """PNG of the watermark image with its alpha scaled by `opacity`."""
    import io
    from PIL import ImageOps
    with Image.open(path) as im:
        im = ImageOps.exif_transpose(im).convert("RGBA")
        if opacity < 0.999:
            a = im.split()[-1].point(lambda v: int(v * opacity))
            im.putalpha(a)
        bio = io.BytesIO(); im.save(bio, "PNG"); return bio.getvalue()


def _watermark_template(w, h, text, fs, fk, opacity, angle):
    """One-page PDF (visual orientation) holding the text watermark."""
    t = fitz.open(); pg = t.new_page(width=w, height=h)
    tw = text_width(text, fs, fk)
    centre = fitz.Point(w / 2, h / 2)
    start = fitz.Point(centre.x - tw / 2, centre.y + fs * 0.35)
    kw = dict(fontsize=fs, color=(0.55, 0.55, 0.55), fill_opacity=opacity, **fk)
    if angle:
        kw["morph"] = (centre, fitz.Matrix(angle))
    pg.insert_text(start, text, **kw)
    return t


def eng_watermark(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    templates = {}
    try:
        text = str(opts.get("text") or "").strip()
        opacity = max(0.02, min(1.0, float(opts.get("opacity", 0.18))))
        size = float(opts.get("font_size", 42))
        img = str(opts.get("image") or "").strip().strip('"')
        layout = str(opts.get("layout", "Diagonal"))
        if img and not Path(img).exists():
            raise ValueError(f"Watermark image not found: {img}")
        if not img and not text:
            raise ValueError("Enter watermark text or choose an image.")
        img_stream = _faded_image_bytes(img, opacity) if img else None
        img_xref = 0; fk = font_kwargs(text); n = len(doc)
        angle = 45 if layout.startswith("Diag") else 0
        for i, p in enumerate(doc):
            vr = p.rect
            if img_stream:
                box = fitz.Rect(vr.width*0.2, vr.height*0.3, vr.width*0.8, vr.height*0.7)
                # Embed the picture once and reference it from every page (xref reuse).
                img_xref = p.insert_image(to_page_space(p, box), stream=None if img_xref else img_stream,
                                          xref=img_xref, overlay=True, keep_proportion=True,
                                          rotate=p.rotation)
            else:
                key = (round(vr.width), round(vr.height))
                if key not in templates:
                    fs = size
                    max_w = (min(vr.width, vr.height) * 1.3) if angle else vr.width - 40
                    while fs > 6 and text_width(text, fs, fk) > max_w:
                        fs -= 1  # shrink to fit instead of silently drawing nothing
                    templates[key] = _watermark_template(vr.width, vr.height, text, fs, fk, opacity, angle)
                p.show_pdf_page(vr * p.derotation_matrix, templates[key], 0, overlay=True, rotate=p.rotation)
            progress(int((i+1)*90/max(1, n)), f"Stamping page {i+1}")
        finish_fonts(doc, "fontfile" in fk and not img_stream)
        save_pdf(doc, out, garbage=4)
    finally:
        for t in templates.values():
            t.close()
        doc.close()
    progress(100, "Done"); return out


def eng_crop(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        l = mm_to_pt(opts.get("left", 0)); r = mm_to_pt(opts.get("right", 0))
        t = mm_to_pt(opts.get("top", 0)); b = mm_to_pt(opts.get("bottom", 0))
        for i, p in enumerate(doc):
            vr = p.rect  # what the user sees (rotation applied)
            vis = fitz.Rect(vr.x0 + l, vr.y0 + t, vr.x1 - r, vr.y1 - b)
            if vis.width <= 10 or vis.height <= 10:
                raise ValueError(f"Crop margins are too large for page {i+1}.")
            # set_cropbox wants unrotated coordinates relative to the current CropBox origin.
            unrot = to_page_space(p, vis)
            cb = p.cropbox
            p.set_cropbox(fitz.Rect(unrot.x0 + cb.x0, unrot.y0 + cb.y0, unrot.x1 + cb.x0, unrot.y1 + cb.y0) & p.mediabox)
            progress(int((i+1)*90/max(1, len(doc))), f"Cropping page {i+1}")
        save_pdf(doc, out, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def eng_edit(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        text = str(opts.get("text") or "")
        if not text.strip():
            raise ValueError("Enter the text to place on the page.")
        size = float(opts.get("font_size", 18))
        pages = list(range(len(doc))) if opts.get("all_pages") else [max(0, min(len(doc)-1, int(opts.get("page", 1))-1))]
        fk = font_kwargs(text)
        x = mm_to_pt(opts.get("x", 20)); y = mm_to_pt(opts.get("y", 20))
        for idx, pno in enumerate(pages):
            p = doc[pno]
            # y is the top of the text, measured on the page as the user sees it.
            pt = to_page_point(p, fitz.Point(x, y + size))
            p.insert_text(pt, text, fontsize=size, color=(0, 0, 0), rotate=p.rotation, **fk)
            progress(int((idx+1)*90/max(1, len(pages))), f"Page {pno+1}")
        finish_fonts(doc, "fontfile" in fk)
        save_pdf(doc, out, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def eng_forms(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        lines = [f"{APP_NAME} form report — {Path(files[0]).name}", ""]
        total = 0
        for i, p in enumerate(doc):
            widgets = list(p.widgets() or [])
            total += len(widgets)
            if widgets:
                lines.append(f"Page {i+1}: {len(widgets)} field(s)")
            for w in widgets:
                lines.append(f"  - {w.field_name or '(unnamed)'} | type={w.field_type_string} | "
                             f"value={w.field_value!r} | rect={tuple(round(v,1) for v in w.rect)}")
            progress(int((i+1)*90/max(1, len(doc))), f"Page {i+1}")
        lines.insert(1, f"Total fields: {total}" + ("" if total else "  (this PDF has no interactive form fields)"))
        Path(out).write_text("\n".join(lines), encoding="utf-8")
    finally:
        doc.close()
    progress(100, "Done"); return out


def eng_unlock(progress, files, out, opts):
    try:
        doc = fitz.open(files[0])
    except Exception as exc:
        raise ValueError(f"{Path(files[0]).name} could not be opened ({exc}).")
    try:
        pw = opts.get("password", "")
        if not doc.is_encrypted:
            progress(30, "This PDF is not encrypted — writing a clean copy")
        if doc.needs_pass and not doc.authenticate(pw):
            raise ValueError("Wrong PDF password.")
        progress(45, "Decrypting")
        save_pdf(doc, out, encryption=fitz.PDF_ENCRYPT_NONE, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def eng_protect(progress, files, out, opts):
    user = opts.get("password", ""); owner = opts.get("owner_password", "") or user
    if not user:
        raise ValueError("Enter a password.")
    confirm = opts.get("password_confirm")
    if confirm is not None and confirm != user:
        raise ValueError("The two passwords do not match.")
    doc = open_pdf(files[0], opts)
    try:
        perm = fitz.PDF_PERM_ACCESSIBILITY | fitz.PDF_PERM_PRINT | fitz.PDF_PERM_PRINT_HQ
        if opts.get("allow_copy", True):
            perm |= fitz.PDF_PERM_COPY
        progress(40, "Encrypting AES-256")
        save_pdf(doc, out, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=owner, user_pw=user,
                 permissions=perm, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def eng_sign(progress, files, out, opts):
    sig = str(opts.get("image") or "").strip().strip('"')
    if not sig or not Path(sig).exists():
        raise ValueError("Choose a signature image.")
    doc = open_pdf(files[0], opts)
    try:
        where = str(opts.get("sign_pages", "Last page"))
        position = str(opts.get("position", "Bottom right"))
        targets = {"Last page": [len(doc)-1], "First page": [0], "All pages": list(range(len(doc)))}.get(where, [len(doc)-1])
        width = mm_to_pt(opts.get("width", 45)); margin = mm_to_pt(15)
        with Image.open(sig) as im:
            ratio = im.height / max(1, im.width)
        height = width * ratio
        xref = 0
        for pno in targets:
            p = doc[pno]
            vis = _pos_rect(p.rect, width, height, position, margin=margin)
            xref = p.insert_image(to_page_space(p, vis), filename=None if xref else sig, xref=xref,
                                  overlay=True, keep_proportion=True, rotate=p.rotation)
        progress(70, f"Signature placed on {len(targets)} page(s)")
        save_pdf(doc, out, garbage=4)
    finally:
        doc.close()
    progress(100, "Done"); return out


def eng_redact(progress, files, out, opts):
    raw = str(opts.get("text") or "")
    terms = [t.strip() for t in re.split(r"[\n;|]", raw) if t.strip()]
    if not terms:
        raise ValueError("Enter text to redact (separate several terms with ; ).")
    doc = open_pdf(files[0], opts)
    try:
        total = 0; per_term = {t: 0 for t in terms}
        for i, p in enumerate(doc):
            page_hits = 0
            for term in terms:
                hits = p.search_for(term)
                for r in hits:
                    p.add_redact_annot(r, fill=(0, 0, 0))
                per_term[term] += len(hits); page_hits += len(hits)
            if page_hits:
                p.apply_redactions()
                total += page_hits
            progress(int((i+1)*90/max(1, len(doc))), f"Redacting page {i+1}: {page_hits} match(es)")
        if total == 0:
            raise ValueError("No matching text found. Scanned documents need OCR first.")
        progress(92, "Matches: " + ", ".join(f"'{t}' ×{c}" for t, c in per_term.items()))
        # Remove metadata too: redacted words often also appear in the title/subject.
        try:
            doc.set_metadata({}); doc.del_xml_metadata()
        except Exception:
            pass
        save_pdf(doc, out, garbage=4, clean=True)
    finally:
        doc.close()
    progress(100, f"Done — {total} area(s) permanently redacted"); return out


def eng_compare(progress, files, out, opts):
    if len(files) != 2:
        raise ValueError("Compare needs exactly two PDFs.")
    a = text_from_pdf(files[0]).splitlines(); progress(35, "Read first PDF")
    b = text_from_pdf(files[1]).splitlines(); progress(65, "Read second PDF")
    if not any(x.strip() for x in a + b):
        raise ValueError("Neither PDF contains text. Run OCR on scanned PDFs first.")
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    similarity = sm.ratio() * 100
    html = difflib.HtmlDiff(wrapcolumn=100).make_file(a, b, fromdesc=Path(files[0]).name, todesc=Path(files[1]).name,
                                                       context=True, numlines=3)
    html = html.replace("<body>", f"<body><h2 style='font-family:sans-serif'>{APP_NAME} comparison — "
                                  f"{similarity:.1f}% similar</h2>", 1)
    Path(out).write_text(html, encoding="utf-8"); progress(100, f"Done — texts are {similarity:.1f}% similar"); return out


def eng_summarize(progress, files, out, opts):
    text = text_from_pdf(files[0]); progress(35, "Extracted text")
    if not text.strip():
        raise ValueError("This PDF has no text layer. Run OCR first.")
    # Local extractive heuristic: sentences scored by average word frequency.
    word_re = r"[A-Za-zΑ-Ωα-ωΆ-ώϊϋΐΰ]{4,}"
    sentences = re.split(r"(?<=[.!?;·])\s+", re.sub(r"\s+", " ", text))
    words = re.findall(word_re, text.lower())
    stop = {"this","that","with","from","have","were","will","your","their","there","which","been","they","what",
            "about","would","these","other","into","more","some","than","them","only","also","such",
            "είναι","αυτό","αυτή","αυτά","όπως","οποία","οποίο","οποίος","μέσω","μετά","πριν","έχει","έχουν",
            "στην","στον","στις","στους","στο","στη","από","για","και","των","την","που","ένα","μια","προς","κατά"}
    freq = {}
    for w in words:
        if w not in stop:
            freq[w] = freq.get(w, 0) + 1
    scored = []
    for idx, s in enumerate(sentences):
        sw = re.findall(word_re, s.lower())
        score = sum(freq.get(w, 0) for w in sw) / max(1, len(sw))
        if 30 <= len(s) <= 500:
            scored.append((score, -idx, idx, s.strip()))
    n = int(opts.get("sentences", 10))
    chosen = sorted(sorted(scored, reverse=True)[:n], key=lambda x: x[2])
    summary = "\n\n".join(x[3] for x in chosen) or text[:4000]
    header = f"{APP_NAME} summary — {Path(files[0]).name}\n{'-'*60}\n\n"
    Path(out).write_text(header + summary, encoding="utf-8"); progress(100, "Done"); return out


# Google Translate / deep-translator language list used by the Translate PDF UI.
# Codes are accepted by GoogleTranslator in deep-translator.
TRANSLATE_LANGUAGE_NAMES = {
    "af":"Afrikaans", "sq":"Albanian", "am":"Amharic", "ar":"Arabic", "hy":"Armenian", "as":"Assamese",
    "ay":"Aymara", "az":"Azerbaijani", "bm":"Bambara", "eu":"Basque", "be":"Belarusian", "bn":"Bengali",
    "bho":"Bhojpuri", "bs":"Bosnian", "bg":"Bulgarian", "ca":"Catalan", "ceb":"Cebuano", "ny":"Chichewa",
    "zh-CN":"Chinese (Simplified)", "zh-TW":"Chinese (Traditional)", "co":"Corsican", "hr":"Croatian",
    "cs":"Czech", "da":"Danish", "dv":"Dhivehi", "doi":"Dogri", "nl":"Dutch", "en":"English",
    "eo":"Esperanto", "et":"Estonian", "ee":"Ewe", "tl":"Filipino / Tagalog", "fi":"Finnish", "fr":"French",
    "fy":"Frisian", "gl":"Galician", "ka":"Georgian", "de":"German", "el":"Greek", "gn":"Guarani",
    "gu":"Gujarati", "ht":"Haitian Creole", "ha":"Hausa", "haw":"Hawaiian", "he":"Hebrew", "hi":"Hindi",
    "hmn":"Hmong", "hu":"Hungarian", "is":"Icelandic", "ig":"Igbo", "ilo":"Ilocano", "id":"Indonesian",
    "ga":"Irish", "it":"Italian", "ja":"Japanese", "jw":"Javanese", "kn":"Kannada", "kk":"Kazakh",
    "km":"Khmer", "rw":"Kinyarwanda", "gom":"Konkani", "ko":"Korean", "kri":"Krio", "ku":"Kurdish (Kurmanji)",
    "ckb":"Kurdish (Sorani)", "ky":"Kyrgyz", "lo":"Lao", "la":"Latin", "lv":"Latvian", "ln":"Lingala",
    "lt":"Lithuanian", "lg":"Luganda", "lb":"Luxembourgish", "mk":"Macedonian", "mai":"Maithili",
    "mg":"Malagasy", "ms":"Malay", "ml":"Malayalam", "mt":"Maltese", "mi":"Maori", "mr":"Marathi",
    "mni-Mtei":"Meiteilon (Manipuri)", "lus":"Mizo", "mn":"Mongolian", "my":"Myanmar (Burmese)",
    "ne":"Nepali", "no":"Norwegian", "or":"Odia (Oriya)", "om":"Oromo", "ps":"Pashto", "fa":"Persian",
    "pl":"Polish", "pt":"Portuguese", "pa":"Punjabi", "qu":"Quechua", "ro":"Romanian", "ru":"Russian",
    "sm":"Samoan", "sa":"Sanskrit", "gd":"Scots Gaelic", "nso":"Sepedi", "sr":"Serbian", "st":"Sesotho",
    "sn":"Shona", "sd":"Sindhi", "si":"Sinhala", "sk":"Slovak", "sl":"Slovenian", "so":"Somali",
    "es":"Spanish", "su":"Sundanese", "sw":"Swahili", "sv":"Swedish", "tg":"Tajik", "ta":"Tamil",
    "tt":"Tatar", "te":"Telugu", "th":"Thai", "ti":"Tigrinya", "ts":"Tsonga", "tr":"Turkish",
    "tk":"Turkmen", "ak":"Twi", "uk":"Ukrainian", "ur":"Urdu", "ug":"Uyghur", "uz":"Uzbek",
    "vi":"Vietnamese", "cy":"Welsh", "xh":"Xhosa", "yi":"Yiddish", "yo":"Yoruba", "zu":"Zulu"
}

TRANSLATE_TO_TESSERACT = {
    "af":"afr", "am":"amh", "ar":"ara", "az":"aze", "be":"bel", "bn":"ben", "bs":"bos", "bg":"bul",
    "ca":"cat", "cs":"ces", "cy":"cym", "da":"dan", "de":"deu", "el":"ell", "en":"eng", "eo":"epo",
    "et":"est", "eu":"eus", "fa":"fas", "fi":"fin", "fr":"fra", "ga":"gle", "gl":"glg", "gu":"guj",
    "he":"heb", "hi":"hin", "hr":"hrv", "hu":"hun", "hy":"hye", "id":"ind", "is":"isl", "it":"ita",
    "ja":"jpn", "ka":"kat", "kk":"kaz", "km":"khm", "ko":"kor", "ky":"kir", "lo":"lao", "lt":"lit",
    "lv":"lav", "mk":"mkd", "ml":"mal", "mr":"mar", "ms":"msa", "mt":"mlt", "my":"mya", "ne":"nep",
    "nl":"nld", "no":"nor", "pa":"pan", "pl":"pol", "pt":"por", "ps":"pus", "ro":"ron", "ru":"rus",
    "sa":"san", "si":"sin", "sk":"slk", "sl":"slv", "es":"spa", "sq":"sqi", "sr":"srp", "sv":"swe",
    "sw":"swa", "ta":"tam", "te":"tel", "th":"tha", "tr":"tur", "uk":"ukr", "ur":"urd", "uz":"uzb",
    "vi":"vie", "yi":"yid", "zh-CN":"chi_sim", "zh-TW":"chi_tra"
}


def _translation_chunks(text: str, max_chars: int = 3500) -> list[str]:
    """Split text on paragraph/sentence boundaries where practical, avoiding mid-word cuts."""
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    chunks, current = [], ""
    for para in paragraphs:
        parts = [para]
        if len(para) > max_chars:
            parts = re.split(r"(?<=[.!?;:])\s+", para)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if len(part) > max_chars:
                # Hard-slice pathological "words" (URLs, base64 …) so no chunk exceeds the provider limit.
                words = [w[k:k + max_chars] for w in part.split() for k in range(0, len(w), max_chars)]
                buf = ""
                for word in words:
                    candidate = (buf + " " + word).strip()
                    if len(candidate) > max_chars and buf:
                        if current:
                            chunks.append(current)
                            current = ""
                        chunks.append(buf)
                        buf = word
                    else:
                        buf = candidate
                if buf:
                    if current and len(current) + len(buf) + 2 <= max_chars:
                        current += "\n\n" + buf
                    else:
                        if current:
                            chunks.append(current)
                        current = buf
                continue
            candidate = part if not current else current + "\n\n" + part
            if len(candidate) <= max_chars:
                current = candidate
            else:
                if current:
                    chunks.append(current)
                current = part
    if current:
        chunks.append(current)
    return chunks or ([text.strip()] if text.strip() else [])


def _translation_source_text(progress, pdf_path: str, opts: dict) -> tuple[str, int, str]:
    """Extract page text and optionally OCR pages that appear scanned/empty."""
    doc = open_pdf(pdf_path, opts)
    try:
        page_texts = []
        ocr_indexes = []
        auto_ocr = bool(opts.get("translate_ocr_scans", True))
        for i, page in enumerate(doc):
            txt = page.get_text("text") or ""
            page_texts.append(txt)
            if auto_ocr and sum(1 for ch in txt if ch.isalpha()) < 30:
                ocr_indexes.append(i)

        ocr_lang = ""
        if ocr_indexes:
            try:
                pytesseract, tess = _setup_tesseract()
            except RuntimeError as exc:
                raise RuntimeError(f"{len(ocr_indexes)} page(s) appear scanned. {exc} "
                                   "Or untick automatic OCR for scanned pages.")
            installed = _tesseract_installed_languages(tess)
            source_mode = str(opts.get("translate_source_mode", "Smart / Auto detect"))
            source_code = str(opts.get("translate_source", "en"))
            candidate = TRANSLATE_TO_TESSERACT.get(source_code) if source_mode.startswith("Manual") else None
            if candidate and (not installed or candidate in installed):
                ocr_lang = candidate
                why = f"manual source {TRANSLATE_LANGUAGE_NAMES.get(source_code, source_code)} → OCR {candidate}"
            else:
                ocr_lang, why = _smart_ocr_language(doc, pytesseract, installed, 0.70, progress)
                if source_mode.startswith("Manual"):
                    why = "manual source has no matching installed OCR pack; " + why
            can_rotate = not installed or "osd" in installed

            progress(12, f"Scan fallback: {len(ocr_indexes)} page(s) • {why}")
            for n, page_index in enumerate(ocr_indexes, 1):
                page = doc[page_index]
                if can_rotate:
                    _auto_upright(page, pytesseract)
                pix = page.get_pixmap(dpi=200, alpha=False)
                page_texts[page_index] = pytesseract.image_to_string(_pix_to_pil(pix), lang=ocr_lang)
                progress(12 + int(n * 18 / max(1, len(ocr_indexes))), f"OCR page {page_index+1}/{len(doc)} • {ocr_lang}")
    finally:
        doc.close()
    assembled = [f"=== Page {i} ===\n\n{txt.strip()}" for i, txt in enumerate(page_texts, 1) if txt.strip()]
    return "\n\n".join(assembled), len(ocr_indexes), ocr_lang


def _translate_with_retry(translator, chunk: str, attempts: int = 4) -> str:
    last = None
    for attempt in range(attempts):
        try:
            res = translator.translate(chunk)
            return res if isinstance(res, str) else (chunk if res is None else str(res))
        except Exception as exc:
            last = exc
            if attempt + 1 < attempts:
                time.sleep(1.0 + attempt * 2.0)
    raise RuntimeError(f"Translation provider failed after {attempts} attempts "
                       f"(check the Internet connection): {last}")


def eng_translate(progress, files, out, opts):
    try:
        from deep_translator import GoogleTranslator
    except Exception:
        raise RuntimeError("Translate PDF needs deep-translator. Install optional packages from Dependency status.")

    source_mode = str(opts.get("translate_source_mode", "Smart / Auto detect"))
    source = "auto" if source_mode.startswith("Smart") else str(opts.get("translate_source", "en"))
    target = str(opts.get("translate_target", "el"))
    if source != "auto" and source == target:
        raise ValueError("Source and target language are the same. Choose a different target language.")

    progress(2, "Preparing text for translation…")
    text, ocr_pages, ocr_lang = _translation_source_text(progress, files[0], opts)
    if not text.strip():
        raise ValueError("No readable text was found in this PDF.")

    src_label = "Auto detect" if source == "auto" else f"{TRANSLATE_LANGUAGE_NAMES.get(source, source)} ({source})"
    tgt_label = f"{TRANSLATE_LANGUAGE_NAMES.get(target, target)} ({target})"
    detail = f"Source: {src_label} → Target: {tgt_label}"
    if ocr_pages:
        detail += f" • OCR fallback: {ocr_pages} page(s) [{ocr_lang}]"
    progress(32, detail)

    chunks = _translation_chunks(text, 3400)
    if not chunks:
        raise ValueError("No translatable text was found in this PDF.")
    translator = GoogleTranslator(source=source, target=target)
    translated = []
    for i, chunk in enumerate(chunks):
        translated.append(_translate_with_retry(translator, chunk))
        progress(32 + int((i + 1) * 64 / len(chunks)), f"Translating chunk {i+1}/{len(chunks)}")

    header = (f"{APP_NAME} translated document\n"
              f"Source file: {Path(files[0]).name}\n"
              f"{detail}\n"
              f"{'-'*72}\n\n")
    Path(out).write_text(header + "\n\n".join(translated), encoding="utf-8")
    progress(100, "Translation completed")
    return out


def eng_markdown(progress, files, out, opts):
    md = None
    if importlib.util.find_spec("pymupdf4llm") is not None:
        try:
            import pymupdf4llm
            progress(10, "Converting with pymupdf4llm (layout-aware)")
            doc = open_pdf(files[0], opts)
            try:
                md = pymupdf4llm.to_markdown(doc)
            finally:
                doc.close()
            progress(90, "Converted with pymupdf4llm")
        except Exception as exc:
            progress(15, f"⚠ pymupdf4llm failed ({exc}); using the built-in converter")
            md = None
    if md is None:
        doc = open_pdf(files[0], opts); parts = []
        try:
            for i, p in enumerate(doc):
                parts.append(f"## Page {i+1}\n\n" + (p.get_text("text") or "").strip())
                progress(int((i+1)*90/max(1, len(doc))), f"Page {i+1}")
        finally:
            doc.close()
        md = f"# {Path(files[0]).stem}\n\n" + "\n\n".join(parts)
    Path(out).write_text(md, encoding="utf-8"); progress(100, "Done"); return out


def eng_extract_text(progress, files, out, opts):
    out_dir = Path(out); out_dir.mkdir(parents=True, exist_ok=True)
    doc = open_pdf(files[0], opts); alltxt = []; img_count = 0; seen = set()
    try:
        for i, p in enumerate(doc):
            alltxt.append(f"\n--- Page {i+1} ---\n" + (p.get_text("text") or ""))
            for img in p.get_images(full=True):
                xref = img[0]
                if xref in seen:
                    continue  # logos/backgrounds repeated on every page are saved once
                seen.add(xref)
                try:
                    info = doc.extract_image(xref)
                    if not info or not info.get("image"):
                        continue
                    img_count += 1
                    (out_dir / f"page{i+1:03d}_image{img_count:03d}.{info['ext']}").write_bytes(info["image"])
                except Exception:
                    pass
            progress(int((i+1)*90/max(1, len(doc))), f"Page {i+1}")
        (out_dir / "extracted_text.txt").write_text("".join(alltxt), encoding="utf-8")
    finally:
        doc.close()
    progress(100, f"Done — text + {img_count} unique image(s)"); return str(out_dir)


def eng_copy(progress, files, out, opts):
    doc = open_pdf(files[0], opts)
    try:
        progress(40, "Rewriting PDF")
        save_pdf(doc, out, garbage=4, clean=True)
    finally:
        doc.close()
    progress(100, "Done"); return out


def eng_workflow(progress, files, out, opts):
    """Merge → optimise → optionally protect, entirely in memory (no temp file)."""
    if not files:
        raise ValueError("Add at least one PDF.")
    password = opts.get("password", "")
    dst = fitz.open()
    try:
        for i, f in enumerate(files):
            src = open_pdf(f, opts)
            try:
                dst.insert_pdf(src)
            finally:
                src.close()
            progress(int((i+1)*50/len(files)), f"Merge: added {Path(f).name}")
        progress(60, "Optimizing")
        kw = dict(garbage=4, clean=True, deflate_images=True, deflate_fonts=True)
        if password:
            kw.update(encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=password, owner_pw=password,
                      permissions=fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ACCESSIBILITY)
            progress(75, "Encrypting AES-256")
        save_pdf(dst, out, **kw)
    finally:
        dst.close()
    progress(100, "Done"); return out

ENGINE = {
    "merge":eng_merge,"split":eng_split,"remove":eng_remove,"extract":eng_extract,"organize":eng_organize,
    "scan":_images_to_pdf,"img2pdf":_images_to_pdf,"compress":eng_compress,"repair":eng_repair,"ocr":eng_ocr,
    "word2pdf":eng_office2pdf,"ppt2pdf":eng_office2pdf,"xls2pdf":eng_office2pdf,"html2pdf":eng_html2pdf,
    "pdf2jpg":eng_pdf2jpg,"pdf2word":eng_pdf2word,"pdf2ppt":eng_pdf2ppt,"pdf2xls":eng_pdf2xls,"pdfa":eng_pdfa,
    "rotate":eng_rotate,"numbers":eng_numbers,"watermark":eng_watermark,"crop":eng_crop,"edit":eng_edit,"forms":eng_forms,
    "unlock":eng_unlock,"protect":eng_protect,"sign":eng_sign,"redact":eng_redact,"compare":eng_compare,
    "summarize":eng_summarize,"translate":eng_translate,"markdown":eng_markdown,"extract_text":eng_extract_text,
    "copy":eng_copy,"workflow":eng_workflow,
}


# ----------------------------- Smart batch dispatch -----------------------------
BATCH_DIR_ENGINES = {"pdf2jpg", "extract_text", "word2pdf", "ppt2pdf", "xls2pdf"}
BATCH_EXTENSIONS = {
    "remove": ("_removed", ".pdf"), "extract": ("_extracted", ".pdf"),
    "compress": ("_compressed", ".pdf"), "repair": ("_repaired", ".pdf"), "ocr": ("_ocr", ".pdf"),
    "pdf2word": ("", ".docx"), "pdf2ppt": ("", ".pptx"), "pdf2xls": ("", ".xlsx"), "pdfa": ("_pdfa", ".pdf"),
    "rotate": ("_rotated", ".pdf"), "numbers": ("_numbered", ".pdf"), "watermark": ("_watermarked", ".pdf"),
    "crop": ("_cropped", ".pdf"), "edit": ("_edited", ".pdf"), "forms": ("_forms", ".txt"),
    "unlock": ("_unlocked", ".pdf"), "protect": ("_protected", ".pdf"), "sign": ("_signed", ".pdf"),
    "redact": ("_redacted", ".pdf"), "summarize": ("_summary", ".txt"), "translate": ("_translated", ".txt"),
    "markdown": ("", ".md"), "copy": ("_copy", ".pdf"),
}


def _sanitize_stem(name: str) -> str:
    name = re.sub(r"[^\w. -]+", "_", name, flags=re.UNICODE).strip(" ._")
    return name[:120] or "document"


def _batch_target(engine_key: str, source: str, output_root: str, index: int) -> str:
    root = Path(output_root); root.mkdir(parents=True, exist_ok=True)
    stem = _sanitize_stem(Path(source).stem)
    # Directory-producing engines get an isolated per-input folder to avoid collisions.
    if engine_key in BATCH_DIR_ENGINES:
        return str(root / f"{index:03d}_{stem}")
    suffix, ext = BATCH_EXTENSIONS.get(engine_key, (f"_{engine_key}", ".pdf"))
    return str(root / f"{index:03d}_{stem}{suffix}{ext}")


def _batch_engine_job(args):
    """Runs in a child process: one input file, isolated MuPDF state."""
    engine_key, source, target, opts = args
    fn = ENGINE[engine_key]
    try:
        result = fn(lambda *_: None, [source], target, opts)
        return source, str(result), None
    except Exception as ex:
        return source, "", f"{type(ex).__name__}: {ex}"


def eng_batch(progress, files, out, opts):
    """Run a one-input engine for many files with process isolation and bounded parallelism."""
    engine_key = str(opts.get("_engine_key", ""))
    if engine_key not in ENGINE:
        raise ValueError(f"Unknown batch engine: {engine_key}")
    if not files:
        raise ValueError("No input files.")
    root = Path(out); root.mkdir(parents=True, exist_ok=True)
    continue_on_error = bool(opts.get("continue_on_error", True))
    workers = effective_workers(opts, files)
    # Online translation providers throttle bursts: keep batch translation bounded.
    if engine_key == "translate":
        workers = min(workers, 4 if opts.get("performance_profile") == "Maximum speed" else 2)
    # Tesseract is itself multi-threaded; two OCR processes are usually the sweet spot.
    if engine_key == "ocr":
        workers = min(workers, 2 if opts.get("performance_profile") != "Maximum speed" else workers)

    clean_opts = {k: v for k, v in opts.items() if not k.startswith("_")}
    jobs = [(engine_key, f, _batch_target(engine_key, f, str(root), i), clean_opts) for i, f in enumerate(files, 1)]
    successes, errors = [], []
    t0 = time.monotonic()

    def record(src, result, err, done):
        if err:
            errors.append(f"{Path(src).name}: {err}")
            if not continue_on_error:
                raise RuntimeError(errors[-1])
            msg = f"⚠ {Path(src).name} — {err}"
        else:
            successes.append(result); msg = f"✓ {Path(src).name}"
        eta = ""
        if done < len(jobs):
            remaining = (time.monotonic() - t0) / done * (len(jobs) - done)
            eta = f"  ·  ~{fmt_elapsed(remaining)} left"
        progress(int(done * 100 / len(jobs)), f"{msg}  [{done}/{len(jobs)}]{eta}")

    if workers == 1 or len(jobs) == 1:
        # In-process: show per-page progress inside each file and allow mid-file cancel.
        for i, (key, src, target, o) in enumerate(jobs):
            base, span = i * 100 / len(jobs), 100 / len(jobs)
            progress(int(base), f"Batch {i+1}/{len(jobs)} · {Path(src).name}")
            sub = lambda v, m=None, b=base, s=span: progress(None if v is None else int(b + s * v / 100), m)
            try:
                result, err = str(ENGINE[key](sub, [src], target, o)), None
            except JobCancelled:
                raise
            except Exception as ex:
                result, err = "", f"{type(ex).__name__}: {ex}"
            record(src, result, err, i + 1)
    else:
        progress(1, f"Smart batch: {len(jobs)} files · {workers} isolated worker processes")
        done = 0
        for job, res, exc in run_parallel(progress, _batch_engine_job, jobs, workers, str(root)):
            done += 1
            if exc is not None:
                record(job[1], "", f"{type(exc).__name__}: {exc}", done)
            else:
                record(res[0], res[1], res[2], done)

    report = root / "_PDFMule_batch_report.txt"
    report.write_text(
        f"{APP_NAME} {APP_VERSION} batch report\nEngine: {engine_key}\nInputs: {len(files)}\n"
        f"Successful: {len(successes)}\nFailed: {len(errors)}\nDuration: {fmt_elapsed(time.monotonic() - t0)}\n\n" +
        ("\n".join(errors) if errors else "No errors."), encoding="utf-8")
    if not successes:
        raise RuntimeError("All batch items failed. See _PDFMule_batch_report.txt\n" + "\n".join(errors[:5]))
    progress(100, f"Completed {len(successes)}/{len(files)} file(s)" + (f" with {len(errors)} warning(s)" if errors else ""))
    return str(root)


# ----------------------------- Visual / Smart Split engine -----------------------------
def _smart_ranges_for_doc(doc) -> tuple[list[tuple[int,int]], str]:
    """Return 1-based inclusive ranges and a short explanation."""
    n=len(doc)
    if n <= 1: return ([(1,n)] if n else []), "single-page document"
    try:
        toc=doc.get_toc(simple=True) or []
        starts=sorted({int(row[2]) for row in toc if len(row)>=3 and int(row[0])==1 and 1 <= int(row[2]) <= n})
        if len(starts) >= 2:
            ranges=[]
            if starts[0] > 1: starts=[1]+starts
            for i,a in enumerate(starts):
                b=(starts[i+1]-1) if i+1<len(starts) else n
                if a<=b: ranges.append((a,b))
            return ranges, f"{len(ranges)} sections detected from level-1 bookmarks"
    except Exception:
        pass
    # Detect nearly blank separator pages. This is intentionally conservative.
    blanks=[]
    for i,p in enumerate(doc):
        try:
            txt=re.sub(r"\s+","",p.get_text("text") or "")
            imgs=p.get_images(full=True)
            if len(txt) <= 2 and not imgs: blanks.append(i+1)
        except Exception:
            pass
    if blanks:
        ranges=[]; start=1
        for blank in blanks:
            if blank>start: ranges.append((start,blank-1))
            start=blank+1
        if start<=n: ranges.append((start,n))
        if ranges: return ranges, f"{len(ranges)} sections detected around blank separator pages"
    # Stable fallback: about 8 chunks, capped to 10 pages per chunk.
    chunk=max(1,min(10, math.ceil(n/8)))
    ranges=[(a,min(n,a+chunk-1)) for a in range(1,n+1,chunk)]
    return ranges, f"no bookmarks/separators found; fallback to {chunk}-page chunks"


def _visual_split_one_job(args):
    source, output_root, opts, batch_index = args
    doc=None
    try:
        doc=open_pdf(source, opts)
        n=len(doc); stem=_sanitize_stem(Path(source).stem)
        batch_mode=bool(opts.get("batch_mode",False))
        out_dir=Path(output_root)/(f"{batch_index:03d}_{stem}" if batch_mode else stem+"_split")
        out_dir.mkdir(parents=True,exist_ok=True)
        mode=opts.get("split_mode","range")
        created=[]

        def save_group(pages, label):
            if not pages: return
            d=fitz.open()
            try:
                if pages==list(range(pages[0],pages[-1]+1)):
                    d.insert_pdf(doc,from_page=pages[0],to_page=pages[-1])  # contiguous: one fast call
                else:
                    for pno in pages: d.insert_pdf(doc,from_page=pno,to_page=pno)
                target=out_dir/f"{stem}_{label}.pdf"
                save_pdf(d,str(target),garbage=3,clean=True)
            finally:
                d.close()
            created.append(str(target))

        if mode == "range":
            range_mode=opts.get("range_mode","custom")
            if range_mode == "smart":
                ranges,_=_smart_ranges_for_doc(doc)
            elif range_mode == "fixed":
                chunk=max(1,int(opts.get("fixed_pages",10)))
                ranges=[(a,min(n,a+chunk-1)) for a in range(1,n+1,chunk)]
            else:
                ranges=[]
                for a,b in opts.get("ranges",[]):
                    a,b=int(a),int(b)
                    if a>b: a,b=b,a
                    a,b=max(1,a),min(n,b)          # clamp AFTER ordering
                    if a<=b: ranges.append((a,b))   # range entirely outside this PDF → skipped
                if not ranges: ranges=[(1,n)]
            if opts.get("merge_ranges"):
                pages=[]
                for a,b in ranges: pages.extend(range(a-1,b))
                save_group(pages,"selected_ranges")
            else:
                for i,(a,b) in enumerate(ranges,1): save_group(list(range(a-1,b)),f"range_{i:03d}_{a}-{b}")
        elif mode == "pages":
            pages=page_set(opts.get("pages", ""),n)
            if opts.get("each_page",True):
                for pno in pages: save_group([pno],f"page_{pno+1:04d}")
            else: save_group(pages,"selected_pages")
        elif mode == "size":
            target=max(0.1,float(opts.get("target_mb",5.0)))*1024*1024
            groups=[]; current=[]; approx=0
            # Estimate each page by serializing it alone. More work, but keeps the target mode meaningful.
            for pno in range(n):
                one=fitz.open(); one.insert_pdf(doc,from_page=pno,to_page=pno)
                weight=len(one.tobytes(garbage=3,deflate=True)); one.close()
                if current and approx+weight>target:
                    groups.append(current); current=[]; approx=0
                current.append(pno); approx+=weight
            if current: groups.append(current)
            for i,g in enumerate(groups,1): save_group(g,f"sizepart_{i:03d}_p{g[0]+1}-{g[-1]+1}")
        else:
            raise ValueError(f"Unknown split mode: {mode}")
        return source,str(out_dir),len(created),None
    except Exception as ex:
        return source,"",0,f"{type(ex).__name__}: {ex}"
    finally:
        if doc is not None: doc.close()


def eng_visual_split(progress, files, out, opts):
    root=Path(out); root.mkdir(parents=True,exist_ok=True)
    workers=effective_workers(opts, files)
    opts=dict(opts); opts["batch_mode"]=len(files)>1
    jobs=[(f,str(root),opts,i) for i,f in enumerate(files,1)]
    errors=[]; success=0; total_parts=0

    def record(src,count,err,done):
        nonlocal success,total_parts
        if err: errors.append(f"{Path(src).name}: {err}")
        else: success+=1; total_parts+=count
        progress(int(done*100/len(jobs)), (f"⚠ {Path(src).name}: {err}" if err else f"✓ {Path(src).name}: {count} output PDF(s)"))

    if len(jobs)==1 or workers==1:
        for i,j in enumerate(jobs,1):
            progress(int((i-1)*100/len(jobs)),f"Splitting {Path(j[0]).name}…")
            src,od,count,err=_visual_split_one_job(j)
            record(src,count,err,i)
    else:
        progress(1,f"Batch split · {len(jobs)} PDFs · {workers} isolated workers")
        done=0
        for job,res,exc in run_parallel(progress,_visual_split_one_job,jobs,workers,str(root)):
            done+=1
            if exc is not None: record(job[0],0,f"{type(exc).__name__}: {exc}",done)
            else: record(res[0],res[2],res[3],done)
    report=root/"_PDFMule_split_report.txt"
    report.write_text(f"{APP_NAME} {APP_VERSION} visual/batch split report\nInputs: {len(files)}\nSuccessful: {success}\nFailed: {len(errors)}\nOutput PDFs: {total_parts}\n\n"+("\n".join(errors) if errors else "No errors."),encoding="utf-8")
    if not success: raise RuntimeError("All PDFs failed. See _PDFMule_split_report.txt\n"+"\n".join(errors[:5]))
    progress(100,f"Completed {success}/{len(files)} PDF(s) → {total_parts} file(s)"+(f" with {len(errors)} warning(s)" if errors else ""))
    return str(root)


def stop_threads_for_close(parent, threads, what="A job") -> bool:
    """Called from dialog reject(): Esc, the window X and Close all end up here.

    Destroying a dialog while its QThread still runs crashes Qt, so running
    threads are asked to stop and given time to finish before the dialog closes.
    """
    running = [t for t in threads if t is not None and t.isRunning()]
    if not running:
        return True
    ans = QMessageBox.question(parent, "Job running", f"{what} is still running.\n\nStop it and close this window?")
    if ans != QMessageBox.StandardButton.Yes:
        return False
    for t in running:
        t.requestInterruption()
    QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
    try:
        deadline = time.monotonic() + 15
        while any(t.isRunning() for t in running) and time.monotonic() < deadline:
            QApplication.processEvents(); time.sleep(0.05)
    finally:
        QApplication.restoreOverrideCursor()
    if any(t.isRunning() for t in running):
        QMessageBox.information(parent, "Still stopping",
                                "The current file is still being processed. The window will be closable as soon as it finishes.")
        return False
    return True


# ----------------------------- Visual Organize PDF -----------------------------
class OrganizeThumbnailWorker(QThread):
    page_ready = pyqtSignal(int, bytes)
    progress = pyqtSignal(int, str)
    failed = pyqtSignal(str)
    done = pyqtSignal()

    def __init__(self, path: str, password: str = "", thumb_w: int = 140, thumb_h: int = 190):
        super().__init__(); self.path=path; self.password=password; self.thumb_w=thumb_w; self.thumb_h=thumb_h

    def run(self):
        doc=None
        try:
            doc=fitz.open(self.path)
            if doc.needs_pass and not doc.authenticate(self.password):
                raise ValueError("Could not authenticate the PDF for thumbnail rendering.")
            total=max(1,len(doc))
            for pno in range(len(doc)):
                if self.isInterruptionRequested(): break
                page=doc[pno]; rect=page.rect
                scale=min(self.thumb_w/max(1.0,rect.width),self.thumb_h/max(1.0,rect.height)); scale=max(0.12,min(0.42,scale))
                pix=page.get_pixmap(matrix=fitz.Matrix(scale,scale),alpha=False)
                self.page_ready.emit(pno,pix.tobytes("png"))
                if pno % 3 == 0 or pno+1==len(doc): self.progress.emit(int((pno+1)*100/total),f"Thumbnail {pno+1}/{len(doc)}")
        except Exception:
            self.failed.emit(traceback.format_exc())
        finally:
            if doc is not None:
                try: doc.close()
                except Exception: pass
            self.done.emit()


class OrganizeSaveWorker(QThread):
    progress = pyqtSignal(int, str)
    finished_ok = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, input_pdf: str, output_pdf: str, page_plan: list[dict], password: str = ""):
        super().__init__()
        self.input_pdf = input_pdf
        self.output_pdf = output_pdf
        self.page_plan = page_plan
        self.password = password

    def run(self):
        src = dst = None
        try:
            if not self.page_plan:
                raise ValueError("The output cannot be empty. Keep at least one page.")
            # Encrypted sources need the password that was used to open them in the dialog.
            src = open_pdf(self.input_pdf, {"input_password": self.password})
            dst = fitz.open()
            total = len(self.page_plan)
            for i, entry in enumerate(self.page_plan):
                pno = int(entry["orig"])
                delta = int(entry.get("rotation", 0)) % 360
                dst.insert_pdf(src, from_page=pno, to_page=pno)
                if delta:
                    page = dst[-1]
                    page.set_rotation((int(page.rotation) + delta) % 360)
                if i % 5 == 0 or i == total - 1:
                    self.progress.emit(int((i + 1) * 90 / total), f"Writing page {i+1}/{total}")
            self.progress.emit(95, "Optimizing…")
            save_pdf(dst, self.output_pdf, garbage=4, clean=True)
            self.progress.emit(100, "Done")
            self.finished_ok.emit(self.output_pdf)
        except Exception:
            self.failed.emit(traceback.format_exc())
        finally:
            for d in (dst, src):
                if d is not None:
                    try: d.close()
                    except Exception: pass


class OrganizeDialog(QDialog):
    """Visual PDF page organizer with thumbnail drag/drop and non-destructive edits."""
    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.pdf_path = ""
        self.doc = None
        self.worker = None
        self.loader = None
        self.thumb_items = {}
        self.thumb_base = {}        # orig page -> unrotated thumbnail (rotate from this, never re-rotate a rotated icon)
        self._retired_loaders = []  # keep interrupted threads alive until they finish
        self.password = ""
        self.result_path = None
        self.setWindowTitle(f"Organize PDF — {APP_NAME}")
        self.resize(1180, 780)
        self.setMinimumSize(980, 660)
        self.setAcceptDrops(True)

        root = QVBoxLayout(self)
        root.setContentsMargins(18, 16, 18, 16)
        root.setSpacing(10)

        top = QHBoxLayout()
        titlebox = QVBoxLayout()
        title = QLabel("↕  Visual Organize PDF")
        title.setObjectName("heroTitle")
        sub = QLabel("Drag thumbnails to reorder. Multi-select pages to rotate or delete, with live preview on the right.")
        sub.setObjectName("heroSub")
        titlebox.addWidget(title); titlebox.addWidget(sub)
        top.addLayout(titlebox, 1)
        self.open_pdf_btn = QPushButton("＋ Open PDF")
        self.open_pdf_btn.clicked.connect(self.choose_pdf)
        top.addWidget(self.open_pdf_btn)
        root.addLayout(top)

        toolbar = QHBoxLayout()
        self.info = QLabel("No PDF loaded")
        self.info.setStyleSheet("font-weight:700;")
        toolbar.addWidget(self.info)
        toolbar.addStretch()
        self.select_all_btn = QPushButton("Select all")
        self.left_btn = QPushButton("↶ Rotate left")
        self.right_btn = QPushButton("↷ Rotate right")
        self.delete_btn = QPushButton("🗑 Delete selected")
        self.reset_btn = QPushButton("Reset")
        for b in (self.select_all_btn, self.left_btn, self.right_btn, self.delete_btn, self.reset_btn):
            b.setEnabled(False); toolbar.addWidget(b)
        self.select_all_btn.clicked.connect(self.select_all)
        self.left_btn.clicked.connect(lambda: self.rotate_selected(-90))
        self.right_btn.clicked.connect(lambda: self.rotate_selected(90))
        self.delete_btn.clicked.connect(self.delete_selected)
        self.reset_btn.clicked.connect(self.reset_pages)
        root.addLayout(toolbar)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        left_frame = QFrame()
        left_layout = QVBoxLayout(left_frame)
        left_layout.setContentsMargins(0,0,0,0)
        self.pages = QListWidget()
        self.pages.setObjectName("thumbView")
        self.pages.setViewMode(QListView.ViewMode.IconMode)
        self.pages.setResizeMode(QListView.ResizeMode.Adjust)
        self.pages.setMovement(QListView.Movement.Snap)
        self.pages.setWrapping(True)
        self.pages.setSpacing(10)
        self.pages.setIconSize(QSize(140, 190))
        self.pages.setGridSize(QSize(170, 230))
        self.pages.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.pages.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.pages.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.pages.setDropIndicatorShown(True)
        self.pages.setMinimumWidth(600)
        self.pages.itemSelectionChanged.connect(self.update_preview)
        self.pages.model().rowsMoved.connect(lambda *_: self.refresh_numbers())
        left_layout.addWidget(self.pages)
        splitter.addWidget(left_frame)

        preview_frame = QFrame()
        preview_frame.setObjectName("previewPanel")
        preview_layout = QVBoxLayout(preview_frame)
        preview_layout.setContentsMargins(14,14,14,14)
        ph = QLabel("Live preview")
        ph.setObjectName("toolTitle")
        preview_layout.addWidget(ph)
        self.preview = QLabel("Select a page")
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setMinimumSize(300, 430)
        self.preview.setStyleSheet("border:1px solid #2a3240; border-radius:10px; padding:8px;")
        preview_layout.addWidget(self.preview, 1)
        self.preview_meta = QLabel("")
        self.preview_meta.setObjectName("heroSub")
        self.preview_meta.setAlignment(Qt.AlignmentFlag.AlignCenter)
        preview_layout.addWidget(self.preview_meta)
        splitter.addWidget(preview_frame)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 2)
        root.addWidget(splitter, 1)

        self.progress = QProgressBar(); self.progress.setRange(0,100); self.progress.setValue(0)
        root.addWidget(self.progress)
        self.status = QLabel("Tip: Ctrl+click for multi-select · Del deletes · Ctrl+A selects all · Ctrl+←/→ rotates · drag to reorder.")
        self.status.setObjectName("heroSub")
        root.addWidget(self.status)
        for keys, fn in (("Delete", self.delete_selected), ("Ctrl+A", self.select_all),
                         ("Ctrl+Left", lambda: self.rotate_selected(-90)), ("Ctrl+Right", lambda: self.rotate_selected(90)),
                         ("Ctrl+S", self.save_pdf), ("Ctrl+O", self.choose_pdf)):
            QShortcut(QKeySequence(keys), self, activated=fn)

        bottom = QHBoxLayout()
        self.open_result_btn = QPushButton("Open output folder")
        self.open_result_btn.setEnabled(False)
        self.open_result_btn.clicked.connect(lambda: open_folder(self.result_path) if self.result_path else None)
        bottom.addWidget(self.open_result_btn)
        bottom.addStretch()
        close = QPushButton("Close"); close.clicked.connect(self.reject)
        self.save_btn = QPushButton("Save organized PDF"); self.save_btn.setObjectName("primary")
        self.save_btn.setEnabled(False); self.save_btn.clicked.connect(self.save_pdf)
        bottom.addWidget(close); bottom.addWidget(self.save_btn)
        root.addLayout(bottom)

    def reject(self):
        if self.worker is not None and self.worker.isRunning():
            QMessageBox.information(self,"Save in progress","The organized PDF is still being written. Wait for the save to finish before closing this window.")
            return
        self._retire_loader()
        for l in self._retired_loaders: l.wait(5000)
        if self.doc is not None:
            try: self.doc.close()
            except Exception: pass
            self.doc = None
        super().reject()


    def dragEnterEvent(self, e: QDragEnterEvent):
        urls = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        if any(Path(x).suffix.lower() == ".pdf" for x in urls):
            e.acceptProposedAction()

    def dropEvent(self, e: QDropEvent):
        for u in e.mimeData().urls():
            f = u.toLocalFile()
            if f and Path(f).suffix.lower() == ".pdf":
                self.load_pdf(f); break
        e.acceptProposedAction()

    def choose_pdf(self):
        start = self.settings.get("last_dir", str(Path.home()))
        f, _ = QFileDialog.getOpenFileName(self, "Open PDF", start, "PDF files (*.pdf)")
        if f:
            self.settings["last_dir"] = str(Path(f).parent); save_settings(self.settings)
            self.load_pdf(f)

    def load_pdf(self, path: str):
        try:
            self._retire_loader()
            if self.doc is not None:
                self.doc.close(); self.doc = None
            prev_pw = self.password if path == self.pdf_path else ""
            self.doc = fitz.open(path); self.password=""
            if self.doc.needs_pass and prev_pw and self.doc.authenticate(prev_pw):
                self.password = prev_pw  # Reset re-opens the same file: don't ask again
            elif self.doc.needs_pass:
                pw, ok = QInputDialog.getText(self, "Encrypted PDF", "Password:", QLineEdit.EchoMode.Password)
                if not ok or not self.doc.authenticate(pw):
                    self.doc.close(); self.doc = None
                    raise ValueError("The PDF could not be opened with that password.")
                self.password=pw
            self.pdf_path = path
            self.pages.clear(); self.thumb_items={}; self.thumb_base={}
            self.preview.clear(); self.preview.setText("Preparing page workspace…")
            count = len(self.doc)
            # Create lightweight placeholders immediately; actual rendering runs in a background QThread.
            for i in range(count):
                entry={"orig":i,"rotation":0}; item=QListWidgetItem(); item.setData(Qt.ItemDataRole.UserRole,entry); item.setToolTip(f"Original page {i+1}"); item.setText(f"Page {i+1}"); item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter); self.pages.addItem(item); self.thumb_items[i]=item
                if i and i % 300 == 0: QApplication.processEvents()
            self.info.setText(f"{Path(path).name}  •  {count} pages")
            self.status.setText("Pages are ready. Thumbnails are rendering in the background…")
            for b in (self.select_all_btn,self.left_btn,self.right_btn,self.delete_btn,self.reset_btn,self.save_btn): b.setEnabled(count>0)
            self.progress.setValue(0)
            if count:self.pages.setCurrentRow(0)
            self.refresh_numbers()
            loader=OrganizeThumbnailWorker(path,self.password); self.loader=loader
            # Ignore late signals from a previous (interrupted) loader: they belong to another PDF.
            loader.page_ready.connect(lambda pno,data,l=loader: self.on_thumb_ready(pno,data) if l is self.loader else None)
            loader.progress.connect(lambda p,m,l=loader: self.on_thumb_progress(p,m) if l is self.loader else None)
            loader.failed.connect(lambda e,l=loader: self.on_thumb_fail(e) if l is self.loader else None)
            loader.done.connect(lambda l=loader: self.on_thumb_done() if l is self.loader else None)
            loader.start()
        except Exception as ex:
            QMessageBox.critical(self,"Open PDF",str(ex)); self.info.setText("No PDF loaded")

    def _retire_loader(self):
        if self.loader is not None and self.loader.isRunning():
            self.loader.requestInterruption()
            if not self.loader.wait(1500):
                self._retired_loaders.append(self.loader)  # finishes on its own; signals are ignored
        self.loader = None
        self._retired_loaders = [l for l in self._retired_loaders if l.isRunning()]

    def _apply_icon(self, item):
        d = item.data(Qt.ItemDataRole.UserRole) or {}
        base = self.thumb_base.get(int(d.get("orig", -1)))
        if base is None or base.isNull():
            return
        rot = int(d.get("rotation", 0)) % 360
        pm = base.transformed(QTransform().rotate(rot), Qt.TransformationMode.SmoothTransformation) if rot else base
        item.setIcon(QIcon(pm.scaled(self.pages.iconSize(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)))

    def on_thumb_ready(self,pno,data):
        pm=QPixmap(); pm.loadFromData(data)
        self.thumb_base[int(pno)]=pm
        item=self.thumb_items.get(int(pno))
        if item is not None: self._apply_icon(item)

    def on_thumb_progress(self,p,msg): self.progress.setValue(p); self.status.setText(msg)
    def on_thumb_fail(self,err): self.status.setText("Thumbnail rendering failed (" + friendly_error(err) + "); page editing is still available.")
    def on_thumb_done(self):
        self.loader=None; self.progress.setValue(0); self.status.setText("Loaded. Drag thumbnails to reorder; Ctrl+click for multiple pages.")

    def render_page(self, pno: int, extra_rotation: int, target: QSize) -> QPixmap:
        if self.doc is None:
            return QPixmap()
        page = self.doc[pno]
        # Modest render scale keeps large documents responsive.
        zoom = 1.1 if max(target.width(), target.height()) < 300 else 1.6
        matrix = fitz.Matrix(zoom, zoom).prerotate(int(extra_rotation) % 360)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        fmt = QImage.Format.Format_RGB888 if pix.n == 3 else QImage.Format.Format_RGBA8888
        img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()
        pm = QPixmap.fromImage(img)
        return pm.scaled(target, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)

    def refresh_numbers(self):
        for i in range(self.pages.count()):
            item = self.pages.item(i); d = item.data(Qt.ItemDataRole.UserRole) or {}
            rot = int(d.get("rotation", 0)) % 360
            suffix = f"  ↻{rot}°" if rot else ""
            item.setText(f"Page {i+1}{suffix}")
        self.update_count()

    def update_count(self):
        if not self.pdf_path: return
        n = self.pages.count(); sel = len(self.pages.selectedItems())
        self.info.setText(f"{Path(self.pdf_path).name}  •  {n} pages  •  {sel} selected")
        self.save_btn.setEnabled(n > 0 and self.worker is None)

    def select_all(self):
        self.pages.selectAll(); self.update_count()

    def selected_items(self):
        return self.pages.selectedItems()

    def rotate_selected(self, delta: int):
        items = self.selected_items()
        if not items: return
        for item in items:
            d = dict(item.data(Qt.ItemDataRole.UserRole) or {})
            d["rotation"] = (int(d.get("rotation", 0)) + delta) % 360
            item.setData(Qt.ItemDataRole.UserRole, d)
            self._apply_icon(item)
        self.refresh_numbers(); self.update_preview()
        self.status.setText(f"Rotated {len(items)} selected page(s).")

    def delete_selected(self):
        rows = sorted([self.pages.row(i) for i in self.selected_items()], reverse=True)
        if not rows: return
        if len(rows) == self.pages.count():
            QMessageBox.warning(self, "Delete pages", "Keep at least one page in the output."); return
        for r in rows: self.pages.takeItem(r)
        self.refresh_numbers()
        if self.pages.count(): self.pages.setCurrentRow(min(rows[-1], self.pages.count()-1))
        self.status.setText(f"Removed {len(rows)} page(s) from the output plan. Original PDF is unchanged.")

    def reset_pages(self):
        if not self.pdf_path: return
        current = self.pdf_path
        self.load_pdf(current)
        self.status.setText("Page order and rotations reset to the original PDF.")

    def update_preview(self):
        self.update_count()
        item = self.pages.currentItem()
        if item is None and self.pages.selectedItems(): item = self.pages.selectedItems()[0]
        if item is None or self.doc is None:
            self.preview.setText("Select a page"); self.preview_meta.setText(""); return
        d = item.data(Qt.ItemDataRole.UserRole) or {}
        pno = int(d.get("orig", 0)); rot = int(d.get("rotation", 0)) % 360
        target = QSize(max(260, self.preview.width()-24), max(360, self.preview.height()-24))
        self.preview.setPixmap(self.render_page(pno, rot, target))
        self.preview_meta.setText(f"Output page {self.pages.row(item)+1}  •  Original page {pno+1}  •  Extra rotation {rot}°")

    def page_plan(self):
        return [dict(self.pages.item(i).data(Qt.ItemDataRole.UserRole) or {}) for i in range(self.pages.count())]

    def save_pdf(self):
        if not self.pdf_path or self.pages.count() == 0: return
        start = self.settings.get("last_dir", str(Path(self.pdf_path).parent))
        default = str(Path(start) / f"{Path(self.pdf_path).stem}_organized.pdf")
        out, _ = QFileDialog.getSaveFileName(self, "Save organized PDF", default, "PDF files (*.pdf)")
        if not out: return
        if Path(out).resolve() == Path(self.pdf_path).resolve():
            QMessageBox.warning(self, "Save organized PDF", "Choose a new output file. The original PDF will not be overwritten in-place."); return
        self.settings["last_dir"] = str(Path(out).parent); save_settings(self.settings)
        self.save_btn.setEnabled(False); self.open_pdf_btn.setEnabled(False); self.pages.setEnabled(False)
        for b in (self.select_all_btn, self.left_btn, self.right_btn, self.delete_btn, self.reset_btn): b.setEnabled(False)
        self.progress.setValue(0); self.status.setText("Saving organized PDF…")
        self.worker = OrganizeSaveWorker(self.pdf_path, out, self.page_plan(), self.password)
        self.worker.progress.connect(self.on_save_progress)
        self.worker.finished_ok.connect(self.on_save_ok)
        self.worker.failed.connect(self.on_save_fail)
        self.worker.start()

    def on_save_progress(self, p, msg):
        self.progress.setValue(p); self.status.setText(msg)

    def on_save_ok(self, out):
        self.result_path = out; self.worker = None; self.progress.setValue(100)
        self.status.setText(f"✓ Saved: {out}")
        self.open_pdf_btn.setEnabled(True); self.save_btn.setEnabled(True); self.pages.setEnabled(True); self.open_result_btn.setEnabled(True)
        for b in (self.select_all_btn, self.left_btn, self.right_btn, self.delete_btn, self.reset_btn): b.setEnabled(True)
        add_recent(self.settings, "Organize PDF", out)
        QMessageBox.information(self, "Organize PDF", f"Organized PDF saved successfully.\n\n{out}")

    def on_save_fail(self, err):
        self.worker = None; self.open_pdf_btn.setEnabled(True); self.save_btn.setEnabled(True); self.pages.setEnabled(True)
        for b in (self.select_all_btn, self.left_btn, self.right_btn, self.delete_btn, self.reset_btn): b.setEnabled(True)
        self.status.setText("✗ Save failed")
        QMessageBox.critical(self, "Organize PDF", friendly_error(err))


# ----------------------------- Visual Merge PDF -----------------------------
class MergeThumbnailWorker(QThread):
    source_info = pyqtSignal(str, int)
    page_ready = pyqtSignal(str, int, bytes)
    source_failed = pyqtSignal(str, str)
    status = pyqtSignal(str)
    all_done = pyqtSignal()

    def __init__(self, files: list[str], thumb_w: int = 130, thumb_h: int = 175):
        super().__init__()
        self.files = list(files)
        self.thumb_w = thumb_w
        self.thumb_h = thumb_h

    def run(self):
        try:
            for source in self.files:
                if self.isInterruptionRequested():
                    break
                doc = None
                try:
                    doc = fitz.open(source)
                    if doc.needs_pass:
                        self.source_failed.emit(source, "Password-protected PDF. Unlock it first, then import it here.")
                        continue
                    self.source_info.emit(source, len(doc))
                    for pno in range(len(doc)):
                        if self.isInterruptionRequested():
                            break
                        page = doc[pno]
                        rect = page.rect
                        scale = min(self.thumb_w / max(1.0, rect.width), self.thumb_h / max(1.0, rect.height))
                        scale = max(0.12, min(0.42, scale))
                        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                        self.page_ready.emit(source, pno, pix.tobytes("png"))
                    self.status.emit(f"Loaded {Path(source).name}")
                except Exception as ex:
                    self.source_failed.emit(source, f"{type(ex).__name__}: {ex}")
                finally:
                    if doc is not None:
                        try: doc.close()
                        except Exception: pass
        finally:
            self.all_done.emit()


class VisualMergeSaveWorker(QThread):
    progress = pyqtSignal(int, str)
    finished_ok = pyqtSignal(str)
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, output_pdf: str, page_plan: list[dict], max_open_docs: int = 6):
        super().__init__()
        self.output_pdf = output_pdf
        self.page_plan = list(page_plan)
        self.max_open_docs = max(2, int(max_open_docs))

    def run(self):
        cache = OrderedDict()
        dst = None
        try:
            if not self.page_plan:
                raise ValueError("The merged output cannot be empty.")
            dst = fitz.open()
            total = len(self.page_plan)
            for i, entry in enumerate(self.page_plan):
                if self.isInterruptionRequested():
                    raise JobCancelled()
                source = entry["source"]
                pno = int(entry["page"])
                rotation = int(entry.get("rotation", 0)) % 360
                doc = cache.pop(source, None)
                if doc is None:
                    doc = fitz.open(source)
                    if doc.needs_pass:
                        doc.close()
                        raise RuntimeError(f"Password-protected source: {Path(source).name}")
                cache[source] = doc
                while len(cache) > self.max_open_docs:
                    _, old = cache.popitem(last=False)
                    old.close()
                dst.insert_pdf(doc, from_page=pno, to_page=pno)
                if rotation:
                    page = dst[-1]
                    page.set_rotation((int(page.rotation) + rotation) % 360)
                if i % 2 == 0 or i == total - 1:
                    self.progress.emit(int((i + 1) * 92 / total), f"Writing page {i+1}/{total} — {Path(source).name}")
            if self.isInterruptionRequested():
                raise JobCancelled()
            self.progress.emit(95, "Optimizing merged PDF…")
            save_pdf(dst, self.output_pdf, garbage=4, clean=True)
            dst.close(); dst = None
            for doc in cache.values():
                doc.close()
            cache.clear()
            self.progress.emit(100, "Done")
            self.finished_ok.emit(self.output_pdf)
        except JobCancelled:
            try:
                # Nothing was written yet (saving is atomic), so an existing file
                # chosen as the target is left untouched.
                if dst is not None: dst.close()
                for doc in cache.values(): doc.close()
            except Exception:
                pass
            self.cancelled.emit()
        except Exception:
            try:
                if dst is not None: dst.close()
                for doc in cache.values(): doc.close()
            except Exception:
                pass
            self.failed.emit(traceback.format_exc())


class VisualMergeDialog(QDialog):
    """Visual multi-PDF merger with background thumbnail generation."""
    def __init__(self, settings: dict, parent=None, initial_files: Optional[list[str]] = None):
        super().__init__(parent)
        self.settings = settings
        self.loader = None
        self.save_worker = None
        self.result_path = None
        self.source_paths = []
        self.source_set = set()
        self.source_errors = []
        self.pending_imports = []    # files dropped/added while a batch is still loading
        self.thumb_base = {}         # (source, page) -> unrotated thumbnail
        self._preview_docs = OrderedDict()
        self.setWindowTitle(f"Visual Merge PDF — {APP_NAME}")
        self.resize(1380, 820)
        self.setMinimumSize(1080, 680)
        self.setAcceptDrops(True)

        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 14)
        root.setSpacing(9)

        top = QHBoxLayout()
        titles = QVBoxLayout()
        title = QLabel("🧩  Visual Merge PDF")
        title.setObjectName("heroTitle")
        sub = QLabel("Import as many PDFs as your system can handle. Drag page thumbnails into any order, rotate/delete pages, then save one merged PDF.")
        sub.setObjectName("heroSub"); sub.setWordWrap(True)
        titles.addWidget(title); titles.addWidget(sub); top.addLayout(titles, 1)
        self.add_btn = QPushButton("＋ Add PDFs")
        self.add_btn.clicked.connect(self.choose_files)
        top.addWidget(self.add_btn)
        self.add_folder_btn = QPushButton("＋ Add Folder")
        self.add_folder_btn.setToolTip("Add all PDFs from a folder and its subfolders.")
        self.add_folder_btn.clicked.connect(self.choose_folder)
        top.addWidget(self.add_folder_btn)
        root.addLayout(top)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Sources pane
        src_frame = QFrame(); src_frame.setObjectName("toolCard")
        sl = QVBoxLayout(src_frame); sl.setContentsMargins(12,12,12,12)
        sh = QLabel("Source PDFs"); sh.setObjectName("toolTitle"); sl.addWidget(sh)
        self.sources = QListWidget(); self.sources.setMinimumWidth(235); sl.addWidget(self.sources, 1)
        src_buttons = QHBoxLayout()
        self.remove_source_btn = QPushButton("Remove")
        self.clear_sources_btn = QPushButton("Clear")
        self.remove_source_btn.clicked.connect(self.remove_selected_sources)
        self.clear_sources_btn.clicked.connect(self.clear_all)
        src_buttons.addWidget(self.remove_source_btn); src_buttons.addWidget(self.clear_sources_btn)
        sl.addLayout(src_buttons)
        self.source_count = QLabel("0 PDFs"); self.source_count.setObjectName("heroSub"); sl.addWidget(self.source_count)
        splitter.addWidget(src_frame)

        # Unified page workspace
        page_frame = QFrame()
        pl = QVBoxLayout(page_frame); pl.setContentsMargins(0,0,0,0)
        toolbar = QHBoxLayout()
        self.workspace_info = QLabel("0 pages"); self.workspace_info.setStyleSheet("font-weight:700;")
        self.workspace_info.setMinimumWidth(170)
        toolbar.addWidget(self.workspace_info); toolbar.addStretch()
        self.merge_profile=QComboBox(); self.merge_profile.addItems(["Smart / Auto","Low memory","Maximum speed"]); self.merge_profile.setToolTip("Performance profile: thumbnail size and how many source PDFs stay open while saving."); self.merge_profile.setMinimumContentsLength(12); self.merge_profile.currentTextChanged.connect(self.apply_merge_profile)
        self.select_all_btn = QPushButton("Select all"); self.select_all_btn.setToolTip("Ctrl+A")
        self.dup_btn = QPushButton("⧉ Duplicate"); self.dup_btn.setToolTip("Ctrl+D")
        self.left_btn = QPushButton("↶"); self.left_btn.setToolTip("Rotate left (Ctrl+←)")
        self.right_btn = QPushButton("↷"); self.right_btn.setToolTip("Rotate right (Ctrl+→)")
        self.delete_btn = QPushButton("🗑 Remove"); self.delete_btn.setToolTip("Remove selected pages (Del)")
        self.select_all_btn.clicked.connect(self.select_all)
        self.dup_btn.clicked.connect(self.duplicate_selected)
        self.left_btn.clicked.connect(lambda: self.rotate_selected(-90))
        self.right_btn.clicked.connect(lambda: self.rotate_selected(90))
        self.delete_btn.clicked.connect(self.delete_selected)
        for b in (self.select_all_btn, self.dup_btn, self.left_btn, self.right_btn, self.delete_btn): toolbar.addWidget(b)
        pl.addLayout(toolbar)
        self.pages = QListWidget()
        self.pages.setObjectName("thumbView")
        self.pages.setViewMode(QListView.ViewMode.IconMode)
        self.pages.setResizeMode(QListView.ResizeMode.Adjust)
        self.pages.setMovement(QListView.Movement.Snap)
        self.pages.setWrapping(True)
        self.pages.setSpacing(8)
        self.pages.setIconSize(QSize(130,175))
        self.pages.setGridSize(QSize(165,225))
        self.pages.setWordWrap(True); self.pages.setTextElideMode(Qt.TextElideMode.ElideMiddle)
        self.pages.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.pages.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.pages.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.pages.setDropIndicatorShown(True)
        self.pages.itemSelectionChanged.connect(self.update_preview)
        self.pages.model().rowsMoved.connect(lambda *_: self.refresh_labels())
        pl.addWidget(self.pages, 1)
        splitter.addWidget(page_frame)

        # Preview pane
        prev_frame = QFrame(); prev_frame.setObjectName("toolCard")
        pr = QVBoxLayout(prev_frame); pr.setContentsMargins(12,12,12,12)
        ph = QLabel("Page preview"); ph.setObjectName("toolTitle"); pr.addWidget(ph)
        self.preview = QLabel("Select a page")
        self.preview.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview.setMinimumWidth(300)
        self.preview.setStyleSheet("border:1px solid #2a3240; border-radius:10px; padding:8px;")
        pr.addWidget(self.preview, 1)
        self.preview_meta = QLabel(""); self.preview_meta.setObjectName("heroSub"); self.preview_meta.setWordWrap(True); self.preview_meta.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pr.addWidget(self.preview_meta)
        splitter.addWidget(prev_frame)
        splitter.setStretchFactor(0, 1); splitter.setStretchFactor(1, 4); splitter.setStretchFactor(2, 2)
        root.addWidget(splitter, 1)

        self.progress = QProgressBar(); self.progress.setRange(0,100); self.progress.setValue(0); root.addWidget(self.progress)
        self.status = QLabel("Drop PDFs/folders here or use Add PDFs / Add Folder. Del removes · Ctrl+D duplicates · Ctrl+←/→ rotates.")
        self.status.setObjectName("heroSub"); root.addWidget(self.status)
        for keys, fn in (("Delete", self.delete_selected), ("Ctrl+A", self.select_all), ("Ctrl+D", self.duplicate_selected),
                         ("Ctrl+Left", lambda: self.rotate_selected(-90)), ("Ctrl+Right", lambda: self.rotate_selected(90)),
                         ("Ctrl+S", self.save_pdf)):
            QShortcut(QKeySequence(keys), self, activated=fn)

        bottom = QHBoxLayout()
        self.open_result_btn = QPushButton("Open output folder"); self.open_result_btn.setEnabled(False)
        self.open_result_btn.clicked.connect(lambda: open_folder(self.result_path) if self.result_path else None)
        self.stop_btn = QPushButton("Cancel current job"); self.stop_btn.setEnabled(False); self.stop_btn.clicked.connect(self.cancel_current)
        close = QPushButton("Close"); close.clicked.connect(self.request_close)
        self.save_btn = QPushButton("Merge && Save PDF"); self.save_btn.setObjectName("primary"); self.save_btn.setEnabled(False); self.save_btn.clicked.connect(self.save_pdf)
        bottom.addWidget(self.open_result_btn); bottom.addWidget(self.stop_btn); bottom.addStretch()
        bottom.addWidget(QLabel("Performance:")); bottom.addWidget(self.merge_profile); bottom.addSpacing(12)
        bottom.addWidget(close); bottom.addWidget(self.save_btn)
        root.addLayout(bottom)

        if initial_files:
            self.queue_files(initial_files)

    def dragEnterEvent(self, e: QDragEnterEvent):
        if any(u.isLocalFile() and (Path(u.toLocalFile()).is_dir() or Path(u.toLocalFile()).suffix.lower() == ".pdf") for u in e.mimeData().urls()):
            e.acceptProposedAction()

    def dropEvent(self, e: QDropEvent):
        files = []
        for u in e.mimeData().urls():
            if not u.isLocalFile():
                continue
            p = Path(u.toLocalFile())
            if p.is_dir():
                files.extend(collect_folder_files(str(p), "PDF files (*.pdf)", recursive=True))
            elif p.suffix.lower() == ".pdf":
                files.append(str(p))
        if files: self.queue_files(files)
        e.acceptProposedAction()

    def choose_files(self):
        start = self.settings.get("last_dir", str(Path.home()))
        files, _ = QFileDialog.getOpenFileNames(self, "Add PDF files", start, "PDF files (*.pdf)")
        if files:
            self.settings["last_dir"] = str(Path(files[0]).parent); save_settings(self.settings)
            self.queue_files(files)

    def choose_folder(self):
        start = self.settings.get("last_dir", str(Path.home()))
        folder = QFileDialog.getExistingDirectory(self, "Add folder of PDFs", start)
        if not folder:
            return
        self.settings["last_dir"] = folder; save_settings(self.settings)
        files = collect_folder_files(folder, "PDF files (*.pdf)", recursive=True)
        if not files:
            QMessageBox.information(self, "Add Folder", "No PDF files were found in this folder or its subfolders.")
            return
        self.status.setText(f"Found {len(files)} PDF(s) in folder. Preparing import…")
        self.queue_files(files)

    def queue_files(self, files: list[str]):
        if self.loader is not None and self.loader.isRunning():
            # Queue instead of refusing: the next batch starts automatically.
            self.pending_imports.extend(files)
            self.status.setText(f"Queued {len(files)} more PDF(s); they will load after the current batch.")
            return
        new_files = []
        existing = set(self.source_set)
        for f in files:
            if f and Path(f).suffix.lower() == ".pdf" and f not in existing:
                new_files.append(f); existing.add(f)
        if not new_files:
            return
        self.source_errors.clear()
        for f in new_files:
            self.source_paths.append(f); self.source_set.add(f)
            item = QListWidgetItem(f"⏳  {Path(f).name}\nLoading…")
            item.setData(Qt.ItemDataRole.UserRole, f)
            item.setToolTip(f)
            self.sources.addItem(item)
        self._set_import_busy(True)
        self.progress.setRange(0,0)
        self.status.setText(f"Loading thumbnails from {len(new_files)} PDF(s) in background…")
        tw,th=(100,135) if self.merge_profile.currentText()=="Low memory" else (160,215) if self.merge_profile.currentText()=="Maximum speed" else (130,175)
        self.loader = MergeThumbnailWorker(new_files,tw,th)
        self.loader.source_info.connect(self.on_source_info)
        self.loader.page_ready.connect(self.on_page_ready)
        self.loader.source_failed.connect(self.on_source_failed)
        self.loader.status.connect(self.status.setText)
        self.loader.all_done.connect(self.on_load_done)
        self.loader.start()
        self.update_stats()

    def apply_merge_profile(self, profile: str):
        if profile=="Low memory": self.pages.setIconSize(QSize(105,140)); self.pages.setGridSize(QSize(140,190))
        elif profile=="Maximum speed": self.pages.setIconSize(QSize(150,200)); self.pages.setGridSize(QSize(185,248))
        else: self.pages.setIconSize(QSize(130,175)); self.pages.setGridSize(QSize(165,225))
        self.status.setText("Profile changed. New imports use the selected thumbnail size; merge writing remains ordered and background-safe.")

    def _source_item(self, path: str):
        for i in range(self.sources.count()):
            it = self.sources.item(i)
            if it.data(Qt.ItemDataRole.UserRole) == path:
                return it
        return None

    def on_source_info(self, path: str, count: int):
        item = self._source_item(path)
        if item: item.setText(f"📄  {Path(path).name}\n{count} pages")

    def on_page_ready(self, path: str, pno: int, png: bytes):
        if path not in self.source_set:
            return
        pm = QPixmap(); pm.loadFromData(png, "PNG")
        self.thumb_base[(path, pno)] = pm
        item = QListWidgetItem(QIcon(pm), "")
        item.setData(Qt.ItemDataRole.UserRole, {"source": path, "page": pno, "rotation": 0})
        item.setToolTip(f"{Path(path).name} — source page {pno+1}")
        item.setTextAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.pages.addItem(item)
        if self.pages.count() % 20 == 0:
            self.refresh_labels()

    def on_source_failed(self, path: str, error: str):
        self.source_errors.append(f"{Path(path).name}: {error}")
        item = self._source_item(path)
        if item: item.setText(f"⚠  {Path(path).name}\nNot loaded")

    def on_load_done(self):
        self.loader = None
        self.progress.setRange(0,100); self.progress.setValue(0)
        self._set_import_busy(False)
        self.refresh_labels(); self.update_stats()
        if self.pending_imports:
            nxt, self.pending_imports = self.pending_imports, []
            self.queue_files(nxt)
            return
        if self.source_errors:
            self.status.setText(f"Loaded with {len(self.source_errors)} warning(s).")
            QMessageBox.warning(self, "Some PDFs were not loaded", "\n".join(self.source_errors[:12]))
        else:
            self.status.setText("Ready. Drag thumbnails to mix pages from any source PDF.")

    def _set_import_busy(self, busy: bool):
        # Add buttons stay enabled: new files are queued while a batch loads.
        self.add_btn.setEnabled(True); self.add_folder_btn.setEnabled(True)
        self.remove_source_btn.setEnabled(not busy)
        self.clear_sources_btn.setEnabled(not busy)
        self.pages.setEnabled(not busy)
        for b in (self.select_all_btn, self.dup_btn, self.left_btn, self.right_btn, self.delete_btn):
            b.setEnabled(not busy)
        self.stop_btn.setEnabled(busy or (self.save_worker is not None and self.save_worker.isRunning()))

    def update_stats(self):
        self.source_count.setText(f"{len(self.source_paths)} PDF(s)")
        self.workspace_info.setText(f"{self.pages.count()} output page(s)  •  {len(self.pages.selectedItems())} selected")
        self.save_btn.setEnabled(self.pages.count() > 0 and self.loader is None and self.save_worker is None)

    def refresh_labels(self):
        for i in range(self.pages.count()):
            item = self.pages.item(i); d = item.data(Qt.ItemDataRole.UserRole) or {}
            source = Path(d.get("source", "")).name
            pno = int(d.get("page", 0)) + 1
            rot = int(d.get("rotation", 0)) % 360
            suffix = f" ↻{rot}°" if rot else ""
            stem = Path(source).stem
            short = stem if len(stem) <= 14 else stem[:6] + "…" + stem[-6:]
            item.setText(f"{i+1}  ·  {short} p{pno}{suffix}")
        self.update_stats()

    def select_all(self):
        self.pages.selectAll(); self.update_stats()

    def duplicate_selected(self):
        rows = sorted(self.pages.row(i) for i in self.pages.selectedItems())
        if not rows: return
        # Insert copies after the final selected row, preserving selected order.
        pos = rows[-1] + 1
        copies = []
        for r in rows:
            src = self.pages.item(r)
            cp = QListWidgetItem(src.icon(), src.text())
            cp.setData(Qt.ItemDataRole.UserRole, dict(src.data(Qt.ItemDataRole.UserRole) or {}))
            cp.setToolTip(src.toolTip()); cp.setTextAlignment(Qt.AlignmentFlag.AlignHCenter)
            copies.append(cp)
        for cp in copies:
            self.pages.insertItem(pos, cp); pos += 1
        self.refresh_labels()

    def rotate_selected(self, delta: int):
        for item in self.pages.selectedItems():
            d = dict(item.data(Qt.ItemDataRole.UserRole) or {})
            d["rotation"] = (int(d.get("rotation", 0)) + delta) % 360
            item.setData(Qt.ItemDataRole.UserRole, d)
            # Re-derive from the unrotated base thumbnail: no blur after repeated rotations.
            base = self.thumb_base.get((d.get("source"), int(d.get("page", 0))))
            if base is not None and not base.isNull():
                pm = base.transformed(QTransform().rotate(d["rotation"]), Qt.TransformationMode.SmoothTransformation) if d["rotation"] else base
                item.setIcon(QIcon(pm.scaled(self.pages.iconSize(), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)))
        self.refresh_labels(); self.update_preview()

    def delete_selected(self):
        rows = sorted((self.pages.row(i) for i in self.pages.selectedItems()), reverse=True)
        for r in rows: self.pages.takeItem(r)
        self.refresh_labels(); self.preview.setText("Select a page"); self.preview_meta.clear()

    def remove_selected_sources(self):
        selected = list(self.sources.selectedItems())
        if not selected: return
        remove_paths = {it.data(Qt.ItemDataRole.UserRole) for it in selected}
        for i in range(self.pages.count()-1, -1, -1):
            d = self.pages.item(i).data(Qt.ItemDataRole.UserRole) or {}
            if d.get("source") in remove_paths: self.pages.takeItem(i)
        for it in selected:
            path = it.data(Qt.ItemDataRole.UserRole)
            if path in self.source_set: self.source_paths.remove(path); self.source_set.discard(path)
            self.thumb_base = {k: v for k, v in self.thumb_base.items() if k[0] != path}
            self._close_preview_doc(path)
            self.sources.takeItem(self.sources.row(it))
        self.refresh_labels(); self.update_stats()

    def clear_all(self):
        if self.pages.count() and QMessageBox.question(self, "Clear merge workspace", "Remove all imported PDFs and pages from this merge workspace?") != QMessageBox.StandardButton.Yes:
            return
        self.pages.clear(); self.sources.clear(); self.source_paths.clear(); self.source_set.clear(); self.thumb_base.clear()
        for path in list(self._preview_docs): self._close_preview_doc(path)
        self.preview.setText("Select a page"); self.preview_meta.clear(); self.update_stats()

    def page_plan(self):
        return [dict(self.pages.item(i).data(Qt.ItemDataRole.UserRole) or {}) for i in range(self.pages.count())]

    def update_preview(self):
        self.update_stats()
        items = self.pages.selectedItems()
        if not items:
            return
        d = items[0].data(Qt.ItemDataRole.UserRole) or {}
        source = d.get("source", ""); pno = int(d.get("page", 0)); rot = int(d.get("rotation", 0)) % 360
        try:
            doc = self._preview_doc(source); page = doc[pno]
            target_w, target_h = 500, 620
            rect = page.rect
            scale = min(target_w/max(1.0,rect.width), target_h/max(1.0,rect.height))
            scale = max(0.2, min(1.2, scale))
            pix = page.get_pixmap(matrix=fitz.Matrix(scale,scale).prerotate(rot), alpha=False)
            fmt = QImage.Format.Format_RGB888 if pix.n == 3 else QImage.Format.Format_RGBA8888
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()
            pm = QPixmap.fromImage(img)
            self.preview.setPixmap(pm.scaled(self.preview.size()-QSize(20,20), Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
            row = self.pages.row(items[0]) + 1
            self.preview_meta.setText(f"Output page {row} of {self.pages.count()}\n{Path(source).name} — source page {pno+1} — extra rotation {rot}°")
        except Exception as ex:
            self.preview.setText("Preview unavailable"); self.preview_meta.setText(str(ex))

    def _preview_doc(self, source):
        """Small LRU of open documents so clicking through pages doesn't reopen files."""
        doc = self._preview_docs.pop(source, None)
        if doc is None:
            doc = fitz.open(source)
        self._preview_docs[source] = doc
        while len(self._preview_docs) > 4:
            _, old = self._preview_docs.popitem(last=False)
            try: old.close()
            except Exception: pass
        return doc

    def _close_preview_doc(self, source):
        doc = self._preview_docs.pop(source, None)
        if doc is not None:
            try: doc.close()
            except Exception: pass

    def save_pdf(self):
        if self.save_worker is not None or (self.loader is not None and self.loader.isRunning()):
            return
        if not self.pages.count():
            QMessageBox.warning(self, "Visual Merge", "Add at least one PDF page first."); return
        start = self.settings.get("last_dir", str(Path.home()))
        default = str(Path(start) / "merged_visual.pdf")
        out, _ = QFileDialog.getSaveFileName(self, "Save merged PDF", default, "PDF files (*.pdf)")
        if not out: return
        if any(_same_file(out, x) for x in self.source_paths):
            QMessageBox.warning(self, "Visual Merge", "Choose a new output file. A source PDF will not be overwritten in-place."); return
        self.result_path = None; self.save_btn.setEnabled(False); self.add_btn.setEnabled(False); self.remove_source_btn.setEnabled(False); self.clear_sources_btn.setEnabled(False); self.pages.setEnabled(False); self.stop_btn.setEnabled(True); self.progress.setRange(0,100); self.progress.setValue(0)
        for b in (self.select_all_btn, self.dup_btn, self.left_btn, self.right_btn, self.delete_btn): b.setEnabled(False)
        self.status.setText(f"Merging {self.pages.count()} pages from {len(self.source_paths)} PDF(s)…")
        max_docs=2 if self.merge_profile.currentText()=="Low memory" else 10 if self.merge_profile.currentText()=="Maximum speed" else 6
        self.save_worker = VisualMergeSaveWorker(out, self.page_plan(), max_open_docs=max_docs)
        self.save_worker.progress.connect(self.on_save_progress)
        self.save_worker.finished_ok.connect(self.on_save_ok)
        self.save_worker.failed.connect(self.on_save_fail)
        self.save_worker.cancelled.connect(self.on_save_cancelled)
        self.save_worker.start()

    def on_save_progress(self, p, msg):
        self.progress.setValue(p); self.status.setText(msg)

    def on_save_ok(self, out):
        self.save_worker = None; self.result_path = out; self.progress.setValue(100); self.status.setText(f"✓ Saved {out}")
        self.add_btn.setEnabled(True); self.remove_source_btn.setEnabled(True); self.clear_sources_btn.setEnabled(True); self.pages.setEnabled(True); self.stop_btn.setEnabled(False); self.save_btn.setEnabled(True); self.open_result_btn.setEnabled(True)
        for b in (self.select_all_btn, self.dup_btn, self.left_btn, self.right_btn, self.delete_btn): b.setEnabled(True)
        add_recent(self.settings, "Visual Merge PDF", out)
        QMessageBox.information(self, "Visual Merge PDF", f"Merged PDF saved successfully.\n\n{out}")

    def on_save_fail(self, err):
        self.save_worker = None; self.add_btn.setEnabled(True); self.remove_source_btn.setEnabled(True); self.clear_sources_btn.setEnabled(True); self.pages.setEnabled(True); self.stop_btn.setEnabled(False); self.save_btn.setEnabled(self.pages.count()>0)
        for b in (self.select_all_btn, self.dup_btn, self.left_btn, self.right_btn, self.delete_btn): b.setEnabled(True)
        self.status.setText("✗ Merge failed"); QMessageBox.critical(self, "Visual Merge PDF", friendly_error(err))

    def on_save_cancelled(self):
        self.save_worker = None; self.add_btn.setEnabled(True); self.remove_source_btn.setEnabled(True); self.clear_sources_btn.setEnabled(True); self.pages.setEnabled(True); self.stop_btn.setEnabled(False); self.save_btn.setEnabled(self.pages.count()>0); self.progress.setValue(0)
        for b in (self.select_all_btn, self.dup_btn, self.left_btn, self.right_btn, self.delete_btn): b.setEnabled(True)
        self.status.setText("Merge cancelled. No output was written.")

    def cancel_current(self):
        if self.loader is not None and self.loader.isRunning():
            self.loader.requestInterruption(); self.status.setText("Stopping thumbnail import after the current page…")
        if self.save_worker is not None and self.save_worker.isRunning():
            self.save_worker.requestInterruption(); self.status.setText("Cancelling merge after the current page…")

    def request_close(self):
        self.reject()

    def reject(self):
        if not stop_threads_for_close(self, [self.loader, self.save_worker], "Importing or merging"):
            return
        self.pending_imports.clear()
        for path in list(self._preview_docs): self._close_preview_doc(path)
        super().reject()



# ----------------------------- Visual Split PDF -----------------------------
class RangeRow(QWidget):
    changed = pyqtSignal()
    remove_requested = pyqtSignal(object)
    def __init__(self, index:int, page_count:int, start:int=1, end:int=1, parent=None):
        super().__init__(parent); self.index=index
        lay=QHBoxLayout(self); lay.setContentsMargins(0,0,0,0); lay.setSpacing(8)
        self.label=QLabel(f"Range {index}"); self.label.setMinimumWidth(72); lay.addWidget(self.label)
        self.start=QSpinBox(); self.start.setRange(1,max(1,page_count)); self.start.setValue(max(1,min(page_count,start))); self.start.setPrefix("from ")
        self.end=QSpinBox(); self.end.setRange(1,max(1,page_count)); self.end.setValue(max(1,min(page_count,end))); self.end.setPrefix("to ")
        rm=QPushButton("×"); rm.setFixedWidth(34); rm.setToolTip("Remove range")
        lay.addWidget(self.start,1); lay.addWidget(self.end,1); lay.addWidget(rm)
        self.start.valueChanged.connect(self.changed); self.end.valueChanged.connect(self.changed); rm.clicked.connect(lambda:self.remove_requested.emit(self))
    def value(self):
        a,b=self.start.value(),self.end.value(); return (min(a,b),max(a,b))
    def set_index(self,i): self.index=i; self.label.setText(f"Range {i}")


class SplitPreviewCard(QFrame):
    def __init__(self, title:str, first_pm:QPixmap, last_pm:QPixmap, a:int,b:int,parent=None):
        super().__init__(parent); self.setObjectName("splitRangeCard")
        lay=QVBoxLayout(self); lay.setContentsMargins(14,12,14,12); lay.setSpacing(8)
        titlelab=QLabel(title); titlelab.setAlignment(Qt.AlignmentFlag.AlignCenter); titlelab.setStyleSheet("font-size:11pt;font-weight:700;"); lay.addWidget(titlelab)
        row=QHBoxLayout(); row.setSpacing(12); row.addStretch()
        for pm,num in ([(first_pm,a)] if a==b else [(first_pm,a),(last_pm,b)]):
            box=QFrame(); bl=QVBoxLayout(box); bl.setContentsMargins(8,8,8,6)
            paper=QLabel(); paper.setObjectName("pagePaper"); paper.setAlignment(Qt.AlignmentFlag.AlignCenter); paper.setFixedSize(150,195)
            if not pm.isNull(): paper.setPixmap(pm.scaled(136,180,Qt.AspectRatioMode.KeepAspectRatio,Qt.TransformationMode.SmoothTransformation))
            bl.addWidget(paper); n=QLabel(str(num)); n.setAlignment(Qt.AlignmentFlag.AlignCenter); bl.addWidget(n); row.addWidget(box)
            if num==a and a!=b:
                dots=QLabel("•••"); dots.setStyleSheet("font-size:20pt;font-weight:800;"); dots.setAlignment(Qt.AlignmentFlag.AlignCenter); row.addWidget(dots)
        row.addStretch(); lay.addLayout(row)
        meta=QLabel(f"Pages {a}–{b}  •  {b-a+1} page(s)"); meta.setObjectName("heroSub"); meta.setAlignment(Qt.AlignmentFlag.AlignCenter); lay.addWidget(meta)


class VisualSplitDialog(QDialog):
    """iLovePDF-like visual split workspace with batch support and deterministic smart splitting."""
    def __init__(self, settings:dict, parent=None, initial_files=None):
        super().__init__(parent); self.settings=settings; self.worker=None; self.result_path=None; self.doc=None; self.range_rows=[]; self._updating=False; self.source_passwords={}; self.thumb_cache={}
        self.setWindowTitle(f"Visual / Batch Split PDF — {APP_NAME}"); self.resize(1360,820); self.setMinimumSize(1080,700); self.setAcceptDrops(True)
        root=QHBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)

        # Left workspace
        left=QWidget(); ll=QVBoxLayout(left); ll.setContentsMargins(22,18,22,18); ll.setSpacing(12)
        top=QHBoxLayout(); titlebox=QVBoxLayout(); title=QLabel("✂  Visual / Batch Split PDF"); title.setObjectName("heroTitle"); sub=QLabel("Preview ranges visually, then split one PDF or an unlimited batch with isolated background workers. Page ranges accept 'end', e.g. 5-end."); sub.setObjectName("heroSub"); sub.setWordWrap(True); titlebox.addWidget(title); titlebox.addWidget(sub); top.addLayout(titlebox,1)
        self.add_btn=QPushButton("＋ Add PDFs"); self.add_btn.clicked.connect(self.add_pdfs); top.addWidget(self.add_btn)
        self.add_folder_btn=QPushButton("＋ Add Folder"); self.add_folder_btn.setToolTip("Add all PDFs from a folder and its subfolders."); self.add_folder_btn.clicked.connect(self.add_folder); top.addWidget(self.add_folder_btn); ll.addLayout(top)
        srcrow=QHBoxLayout(); srcrow.addWidget(QLabel("Preview source:")); self.source_combo=QComboBox(); self.source_combo.currentIndexChanged.connect(self.source_changed); srcrow.addWidget(self.source_combo,1)
        self.remove_btn=QPushButton("Remove current"); self.remove_btn.clicked.connect(self.remove_current); srcrow.addWidget(self.remove_btn); self.batch_badge=QLabel("0 PDFs"); self.batch_badge.setObjectName("badge"); srcrow.addWidget(self.batch_badge); ll.addLayout(srcrow)
        self.preview_scroll=QScrollArea(); self.preview_scroll.setWidgetResizable(True); self.preview_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.preview_host=QWidget(); self.preview_layout=QVBoxLayout(self.preview_host); self.preview_layout.setContentsMargins(8,8,8,8); self.preview_layout.setSpacing(14); self.preview_layout.addStretch(); self.preview_scroll.setWidget(self.preview_host); ll.addWidget(self.preview_scroll,1)
        self.preview_status=QLabel("Add one or more PDFs to start."); self.preview_status.setObjectName("heroSub"); ll.addWidget(self.preview_status)
        root.addWidget(left,1)

        # Right settings panel
        panel=QFrame(); panel.setObjectName("previewPanel"); panel.setFixedWidth(440); pr=QVBoxLayout(panel); pr.setContentsMargins(18,18,18,18); pr.setSpacing(12)
        h=QLabel("Split"); h.setAlignment(Qt.AlignmentFlag.AlignCenter); h.setStyleSheet("font-size:20pt;font-weight:800;"); pr.addWidget(h)
        tabrow=QHBoxLayout(); self.mode_buttons=[]
        for key,text in [("range","▣\nRange"),("pages","▦\nPages"),("size","◫\nSize")]:
            b=QPushButton(text); b.setObjectName("modeTab"); b.setCheckable(False); b.clicked.connect(lambda _=False,k=key:self.set_split_mode(k)); tabrow.addWidget(b); self.mode_buttons.append((key,b))
        pr.addLayout(tabrow); self.mode_stack=QStackedWidget(); pr.addWidget(self.mode_stack,1)

        # Range page
        range_page=QWidget(); rl=QVBoxLayout(range_page); rl.setContentsMargins(0,8,0,0); rl.setSpacing(10); rl.addWidget(QLabel("Range mode:"))
        mode_row=QHBoxLayout(); self.range_mode_buttons=[]
        for key,text in [("custom","Custom"),("fixed","Fixed"),("smart","Smart ✨")]:
            b=QPushButton(text); b.setObjectName("modeTab"); b.clicked.connect(lambda _=False,k=key:self.set_range_mode(k)); mode_row.addWidget(b); self.range_mode_buttons.append((key,b))
        rl.addLayout(mode_row)
        self.range_stack=QStackedWidget(); rl.addWidget(self.range_stack)
        # custom
        cust=QWidget(); cl=QVBoxLayout(cust); cl.setContentsMargins(0,0,0,0); self.range_rows_host=QWidget(); self.range_rows_layout=QVBoxLayout(self.range_rows_host); self.range_rows_layout.setContentsMargins(0,0,0,0); cl.addWidget(self.range_rows_host); add_range=QPushButton("＋ Add range"); add_range.clicked.connect(self.add_range_row); cl.addWidget(add_range); self.range_stack.addWidget(cust)
        # fixed
        fixed=QWidget(); fl=QFormLayout(fixed); self.fixed_pages=QSpinBox(); self.fixed_pages.setRange(1,100000); self.fixed_pages.setValue(10); self.fixed_pages.valueChanged.connect(self.refresh_preview); fl.addRow("Pages per output:",self.fixed_pages); self.range_stack.addWidget(fixed)
        # smart
        smart=QWidget(); sl=QVBoxLayout(smart); sl.setContentsMargins(0,0,0,0); self.smart_info=QLabel("Smart uses level-1 bookmarks first, then blank separator pages, then a conservative fixed fallback."); self.smart_info.setWordWrap(True); self.smart_info.setObjectName("heroSub"); sl.addWidget(self.smart_info); analyze=QPushButton("✨ Analyze current PDF"); analyze.clicked.connect(self.analyze_smart); sl.addWidget(analyze); self.range_stack.addWidget(smart)
        self.merge_ranges=QCheckBox("Merge all ranges into one PDF file"); self.merge_ranges.toggled.connect(self.refresh_preview); rl.addWidget(self.merge_ranges); rl.addStretch(); self.mode_stack.addWidget(range_page)

        # Pages page
        pages_page=QWidget(); pgl=QVBoxLayout(pages_page); pgl.setContentsMargins(0,8,0,0); pgf=QFormLayout(); self.pages_spec=QLineEdit(); self.pages_spec.setPlaceholderText("Example: 1-3,5,8"); self.pages_spec.textChanged.connect(self.refresh_preview); pgf.addRow("Pages:",self.pages_spec); self.each_page=QCheckBox("Export each selected page as its own PDF"); self.each_page.setChecked(True); self.each_page.toggled.connect(self.refresh_preview); pgf.addRow("",self.each_page); pgl.addLayout(pgf); pgl.addStretch(); self.mode_stack.addWidget(pages_page)

        # Size page
        size_page=QWidget(); szl=QVBoxLayout(size_page); szl.setContentsMargins(0,8,0,0); szf=QFormLayout(); self.target_mb=QDoubleSpinBox(); self.target_mb.setRange(0.1,2048.0); self.target_mb.setDecimals(1); self.target_mb.setValue(5.0); self.target_mb.setSuffix(" MB"); szf.addRow("Approx. max part size:",self.target_mb); szl.addLayout(szf); note=QLabel("Size mode estimates compressed page weight before grouping. Final files can vary because PDF resources are shared between pages."); note.setWordWrap(True); note.setObjectName("heroSub"); szl.addWidget(note); szl.addStretch(); self.mode_stack.addWidget(size_page)

        perf=QGroupBox("Performance"); pf=QFormLayout(perf); self.profile=QComboBox(); self.profile.addItems(["Smart / Auto","Balanced","Maximum speed","Low memory","Single worker"]); self.profile.currentTextChanged.connect(self.apply_profile); pf.addRow("Profile:",self.profile); self.workers=QSpinBox(); self.workers.setRange(1,max(1,min(16,os.cpu_count() or 4))); self.workers.setValue(min(4,os.cpu_count() or 2)); pf.addRow("Parallel PDF workers:",self.workers); pr.addWidget(perf)
        self.progress=QProgressBar(); self.progress.setRange(0,100); pr.addWidget(self.progress); self.status=QLabel("Ready"); self.status.setWordWrap(True); self.status.setObjectName("heroSub"); pr.addWidget(self.status)
        self.open_result_btn=QPushButton("Open output folder"); self.open_result_btn.setEnabled(False); self.open_result_btn.clicked.connect(lambda:open_folder(self.result_path) if self.result_path else None); pr.addWidget(self.open_result_btn)
        bottom=QHBoxLayout(); close=QPushButton("Close"); close.clicked.connect(self.request_close); self.cancel_btn=QPushButton("Cancel"); self.cancel_btn.setEnabled(False); self.cancel_btn.clicked.connect(self.cancel_job); self.run_btn=QPushButton("Split PDF"); self.run_btn.setObjectName("primary"); self.run_btn.clicked.connect(self.run_split); bottom.addWidget(close); bottom.addWidget(self.cancel_btn); bottom.addWidget(self.run_btn); pr.addLayout(bottom)
        root.addWidget(panel)

        self.split_mode="range"; self.range_mode="custom"; self.set_split_mode("range"); self.set_range_mode("custom")
        if initial_files: self.load_sources(initial_files)

    def dragEnterEvent(self,e):
        if any(u.isLocalFile() and (Path(u.toLocalFile()).is_dir() or Path(u.toLocalFile()).suffix.lower()==".pdf") for u in e.mimeData().urls()): e.acceptProposedAction()
    def dropEvent(self,e):
        files=[]
        for u in e.mimeData().urls():
            if not u.isLocalFile():
                continue
            p=Path(u.toLocalFile())
            if p.is_dir(): files.extend(collect_folder_files(str(p),"PDF files (*.pdf)",recursive=True))
            elif p.suffix.lower()==".pdf": files.append(str(p))
        self.load_sources(files); e.acceptProposedAction()
    def add_pdfs(self):
        fs,_=QFileDialog.getOpenFileNames(self,"Select PDFs",self.settings.get("last_dir",str(Path.home())),"PDF files (*.pdf)")
        if fs: self.settings["last_dir"]=str(Path(fs[0]).parent); save_settings(self.settings); self.load_sources(fs)
    def add_folder(self):
        folder=QFileDialog.getExistingDirectory(self,"Add folder of PDFs",self.settings.get("last_dir",str(Path.home())))
        if not folder: return
        self.settings["last_dir"]=folder; save_settings(self.settings)
        fs=collect_folder_files(folder,"PDF files (*.pdf)",recursive=True)
        if not fs:
            QMessageBox.information(self,"Add Folder","No PDF files were found in this folder or its subfolders.")
            return
        self.load_sources(fs)
        self.preview_status.setText(f"Added {len(fs)} PDF(s) from folder and subfolders.")
    def load_sources(self,files):
        existing={self.source_combo.itemData(i) for i in range(self.source_combo.count())}
        had_items=self.source_combo.count()>0
        added=0
        self.source_combo.blockSignals(True)
        for f in files:
            if f and Path(f).suffix.lower()==".pdf" and f not in existing:
                self.source_combo.addItem(Path(f).name,f); self.source_combo.setItemData(self.source_combo.count()-1,f,Qt.ItemDataRole.ToolTipRole); existing.add(f); added+=1
        self.source_combo.blockSignals(False)
        self.batch_badge.setText(f"{self.source_combo.count()} PDF(s)")
        # Only (re)load the preview when nothing was shown yet — adding more files to a
        # batch must not reopen the current PDF or ask for its password again.
        if self.source_combo.count() and (not had_items or self.doc is None):
            if self.source_combo.currentIndex()<0: self.source_combo.setCurrentIndex(0)
            self.source_changed(self.source_combo.currentIndex())
        if added: self.status.setText(f"{self.source_combo.count()} PDF(s) in batch. Settings apply to every PDF.")
    def remove_current(self):
        i=self.source_combo.currentIndex()
        if i<0: return
        path=self.source_combo.itemData(i); self.source_passwords.pop(path,None)
        self.source_combo.blockSignals(True); self.source_combo.removeItem(i); self.source_combo.blockSignals(False)
        self.batch_badge.setText(f"{self.source_combo.count()} PDF(s)"); self.source_changed(self.source_combo.currentIndex())
    def source_changed(self,index):
        if self.doc is not None:
            try:self.doc.close()
            except:pass
            self.doc=None
        if index<0: self.clear_preview("Add one or more PDFs to start."); return
        path=self.source_combo.itemData(index); self.thumb_cache={}
        if hasattr(self,"_smart_ranges"): delattr(self,"_smart_ranges")
        try:
            self.doc=fitz.open(path)
            if self.doc.needs_pass and not self.doc.authenticate(self.source_passwords.get(path,"\0")):
                pw,ok=QInputDialog.getText(self,"Encrypted PDF",f"Password for {Path(path).name}:",QLineEdit.EchoMode.Password)
                if not ok or not self.doc.authenticate(pw):
                    self.doc.close(); self.doc=None
                    raise ValueError("Wrong password or cancelled.")
                self.source_passwords[path]=pw
            self.pages_spec.setPlaceholderText(f"1-{len(self.doc)}, e.g. 1-3,5")
            if not self.range_rows: self.add_range_row(1,len(self.doc))
            else:
                for r in self.range_rows:
                    r.start.setMaximum(max(1,len(self.doc))); r.end.setMaximum(max(1,len(self.doc)))
            if self.range_mode=="smart": self.analyze_smart(silent=True)
            else: self.refresh_preview()
        except Exception as ex:
            self.clear_preview(str(ex)); QMessageBox.critical(self,"Open PDF",str(ex))
    def set_split_mode(self,mode):
        self.split_mode=mode; idx={"range":0,"pages":1,"size":2}[mode]; self.mode_stack.setCurrentIndex(idx)
        for k,b in self.mode_buttons: b.setProperty("active",k==mode); b.style().unpolish(b); b.style().polish(b)
        self.refresh_preview()
    def set_range_mode(self,mode):
        self.range_mode=mode; idx={"custom":0,"fixed":1,"smart":2}[mode]; self.range_stack.setCurrentIndex(idx)
        for k,b in self.range_mode_buttons: b.setProperty("active",k==mode); b.style().unpolish(b); b.style().polish(b)
        if mode=="smart": self.analyze_smart(silent=True)
        else:self.refresh_preview()
    def add_range_row(self,start=1,end=None):
        n=len(self.doc) if self.doc is not None else 1; end=n if end is None else end
        row=RangeRow(len(self.range_rows)+1,n,start,end); row.changed.connect(self.refresh_preview); row.remove_requested.connect(self.remove_range_row); self.range_rows.append(row); self.range_rows_layout.addWidget(row); self.refresh_preview()
    def remove_range_row(self,row):
        if row in self.range_rows: self.range_rows.remove(row); row.setParent(None); row.deleteLater()
        for i,r in enumerate(self.range_rows,1): r.set_index(i)
        self.refresh_preview()
    def apply_profile(self,text):
        maxcpu=max(1,min(16,os.cpu_count() or 4))
        if text=="Low memory" or text=="Single worker": self.workers.setValue(1)
        elif text=="Balanced": self.workers.setValue(min(2,maxcpu))
        elif text=="Maximum speed": self.workers.setValue(maxcpu)
        else:self.workers.setValue(min(4,maxcpu))
    def analyze_smart(self,silent=False):
        if self.doc is None:return
        ranges,why=_smart_ranges_for_doc(self.doc); self.smart_info.setText(f"Smart analysis: {why}.")
        self._smart_ranges=ranges; self.refresh_preview()
        if not silent:self.status.setText(f"Smart analysis complete: {why}.")
    def current_ranges(self):
        if self.doc is None:return []
        n=len(self.doc)
        if self.range_mode=="fixed":
            c=max(1,self.fixed_pages.value()); return [(a,min(n,a+c-1)) for a in range(1,n+1,c)]
        if self.range_mode=="smart":
            if not hasattr(self,"_smart_ranges"): self._smart_ranges,_=_smart_ranges_for_doc(self.doc)
            return list(self._smart_ranges)
        return [r.value() for r in self.range_rows] or [(1,n)]
    def render_thumb(self,pno):
        if self.doc is None or not (0<=pno<len(self.doc)):return QPixmap()
        if pno in self.thumb_cache:return self.thumb_cache[pno]
        p=self.doc[pno]; rect=p.rect; scale=min(0.9,150/max(1,rect.width),190/max(1,rect.height)); scale=max(0.15,scale)
        pix=p.get_pixmap(matrix=fitz.Matrix(scale,scale),alpha=False); fmt=QImage.Format.Format_RGB888 if pix.n==3 else QImage.Format.Format_RGBA8888; img=QImage(pix.samples,pix.width,pix.height,pix.stride,fmt).copy(); pm=QPixmap.fromImage(img); self.thumb_cache[pno]=pm
        # Bound cache size for very large documents.
        if len(self.thumb_cache)>160:self.thumb_cache.pop(next(iter(self.thumb_cache)))
        return pm
    def clear_preview(self,msg=""):
        while self.preview_layout.count()>1:
            it=self.preview_layout.takeAt(0); w=it.widget();
            if w: w.hide(); w.setParent(None); w.deleteLater()  # hide now: no ghost cards until deletion
        self.preview_status.setText(msg)
    def refresh_preview(self,*_):
        if self.doc is None:return
        while self.preview_layout.count()>1:
            it=self.preview_layout.takeAt(0); w=it.widget();
            if w: w.hide(); w.setParent(None); w.deleteLater()  # hide now: no ghost cards until deletion
        try:
            if self.split_mode=="range":
                ranges=self.current_ranges(); shown=ranges[:24]
                for i,(a,b) in enumerate(shown,1):
                    a=max(1,min(len(self.doc),a)); b=max(1,min(len(self.doc),b));
                    if a>b:a,b=b,a
                    card=SplitPreviewCard(f"Range {i}",self.render_thumb(a-1),self.render_thumb(b-1),a,b); self.preview_layout.insertWidget(self.preview_layout.count()-1,card)
                self.preview_status.setText(f"{len(ranges)} output range(s) previewed"+(" · showing first 24" if len(ranges)>24 else ""))
            elif self.split_mode=="pages":
                spec=self.pages_spec.text().strip()
                pages=page_set(spec,len(self.doc)) if spec else list(range(len(self.doc)))
                preview_pages=pages[:24]
                for i,p in enumerate(preview_pages,1):
                    card=SplitPreviewCard(f"Selected page {p+1}",self.render_thumb(p),self.render_thumb(p),p+1,p+1); self.preview_layout.insertWidget(self.preview_layout.count()-1,card)
                self.preview_status.setText(f"{len(pages)} selected page(s)"+(" · showing first 24" if len(pages)>24 else ""))
            else:
                # Size mode preview intentionally lightweight; exact grouping happens in the worker.
                first=self.render_thumb(0); last=self.render_thumb(len(self.doc)-1)
                card=SplitPreviewCard(f"Target ≈ {self.target_mb.value():.1f} MB per part",first,last,1,len(self.doc)); self.preview_layout.insertWidget(self.preview_layout.count()-1,card)
                self.preview_status.setText("Exact size grouping is calculated in the background when you run the split.")
        except Exception as ex:
            self.preview_status.setText(f"Preview: {ex}")
    def options(self):
        opts={"split_mode":self.split_mode,"workers":self.workers.value(),"performance_profile":self.profile.currentText(),"passwords":dict(self.source_passwords)}
        if self.split_mode=="range": opts.update({"range_mode":self.range_mode,"ranges":[r.value() for r in self.range_rows],"fixed_pages":self.fixed_pages.value(),"merge_ranges":self.merge_ranges.isChecked()})
        elif self.split_mode=="pages": opts.update({"pages":self.pages_spec.text(),"each_page":self.each_page.isChecked()})
        else:opts["target_mb"]=self.target_mb.value()
        return opts
    def run_split(self):
        files=[self.source_combo.itemData(i) for i in range(self.source_combo.count())]
        if not files:QMessageBox.warning(self,"Split PDF","Add at least one PDF.");return
        # Resolve passwords before launching the batch so encrypted files do not fail silently in child processes.
        for f in files:
            try:
                d=fitz.open(f)
                if d.needs_pass and f not in self.source_passwords:
                    pw,ok=QInputDialog.getText(self,"Encrypted PDF",f"Password for {Path(f).name}:",QLineEdit.EchoMode.Password)
                    if not ok or not d.authenticate(pw): d.close(); QMessageBox.warning(self,"Split PDF",f"Could not authenticate {Path(f).name}."); return
                    self.source_passwords[f]=pw
                d.close()
            except Exception as ex:
                QMessageBox.warning(self,"Split PDF",f"Could not inspect {Path(f).name}: {ex}"); return
        out=QFileDialog.getExistingDirectory(self,"Select output folder",self.settings.get("last_dir",str(Path.home())))
        if not out:return
        self.settings["last_dir"]=out; save_settings(self.settings)
        self.result_path=None; self.run_btn.setEnabled(False); self.add_btn.setEnabled(False); self.add_folder_btn.setEnabled(False); self.remove_btn.setEnabled(False); self.cancel_btn.setEnabled(True); self.open_result_btn.setEnabled(False); self.progress.setValue(0); self._t0=time.monotonic()
        self.status.setText(f"Starting split for {len(files)} PDF(s)…")
        self.worker=Worker(eng_visual_split,files,out,self.options()); self.worker.progress.connect(self.on_progress); self.worker.finished_ok.connect(self.on_ok); self.worker.failed.connect(self.on_fail); self.worker.cancelled.connect(self.on_cancelled); self.worker.start()
    def on_progress(self,p,msg):self.progress.setValue(p);self.status.setText(msg)
    def _unlock_ui(self):self.run_btn.setEnabled(True);self.add_btn.setEnabled(True);self.add_folder_btn.setEnabled(True);self.remove_btn.setEnabled(True);self.cancel_btn.setEnabled(False)
    def on_ok(self,result):
        last=self.status.text()
        self.worker=None;self._unlock_ui();self.result_path=str(result);self.open_result_btn.setEnabled(True);self.progress.setValue(100)
        took=fmt_elapsed(time.monotonic()-getattr(self,"_t0",time.monotonic()))
        self.status.setText(f"✓ {last}  ·  {took}")
        add_recent(self.settings,"Visual / Batch Split PDF",self.result_path)
        box=QMessageBox(self); box.setIcon(QMessageBox.Icon.Information); box.setWindowTitle("Split PDF"); box.setText(f"{last}\nTime: {took}\n\n{self.result_path}")
        ob=box.addButton("Open folder",QMessageBox.ButtonRole.ActionRole); box.addButton(QMessageBox.StandardButton.Ok); box.exec()
        if box.clickedButton()==ob: open_folder(self.result_path)
    def on_fail(self,error):self.worker=None;self._unlock_ui();self.status.setText("✗ Split failed");QMessageBox.critical(self,"Split PDF",friendly_error(error))
    def on_cancelled(self):self.worker=None;self._unlock_ui();self.progress.setValue(0);self.status.setText("Cancellation requested / job stopped safely.")
    def cancel_job(self):
        if self.worker and self.worker.isRunning():self.worker.requestInterruption();self.cancel_btn.setEnabled(False);self.status.setText("Cancelling… worker processes are being stopped.")
    def request_close(self):
        self.reject()
    def reject(self):
        if not stop_threads_for_close(self,[self.worker],"The split job"): return
        if self.doc is not None:
            try:self.doc.close()
            except Exception:pass
            self.doc=None
        super().reject()



class OcrLanguageSelector(QWidget):
    """Smart/manual Tesseract language selector with search and checkbox multi-select."""
    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(8)

        self.mode = QComboBox()
        self.mode.addItems(["Smart / Auto detect", "Manual multi-select"])
        outer.addWidget(self.mode)

        self.manual_panel = QWidget()
        mp = QVBoxLayout(self.manual_panel)
        mp.setContentsMargins(0, 2, 0, 0)
        mp.setSpacing(7)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Search language or code…  e.g. Greek, ell, German, deu")
        self.search.setClearButtonEnabled(True)
        mp.addWidget(self.search)

        self.list = QListWidget()
        self.list.setMinimumHeight(210)
        self.list.setMaximumHeight(260)
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        for code, name in sorted(TESSERACT_LANGUAGE_NAMES.items(), key=lambda kv: (kv[1].casefold(), kv[0])):
            item = QListWidgetItem(f"{name}   ({code})")
            item.setData(Qt.ItemDataRole.UserRole, code)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if code == "eng" else Qt.CheckState.Unchecked)
            self.list.addItem(item)
        mp.addWidget(self.list)

        row = QHBoxLayout()
        greek_eng = QPushButton("Greek + English")
        english = QPushButton("English only")
        clear = QPushButton("Clear")
        greek_eng.clicked.connect(lambda: self.set_codes(["ell", "eng"]))
        english.clicked.connect(lambda: self.set_codes(["eng"]))
        clear.clicked.connect(lambda: self.set_codes([]))
        row.addWidget(greek_eng); row.addWidget(english); row.addWidget(clear); row.addStretch()
        mp.addLayout(row)

        note = QLabel(
            "All standard Tesseract language codes are listed. You can tick several languages; "
            "PDF Mule sends them to Tesseract as e.g. ell+eng. Selected language packs are "
            "validated when OCR starts."
        )
        note.setWordWrap(True)
        note.setObjectName("heroSub")
        mp.addWidget(note)
        outer.addWidget(self.manual_panel)

        smart_note = QLabel(
            "Smart mode decides separately for every PDF. If ≥ the selected threshold of alphabetic "
            "characters is Greek, the document is OCRed as Greek (ell). For scans without embedded "
            "text it samples up to 3 pages with Tesseract OSD."
        )
        smart_note.setWordWrap(True)
        smart_note.setObjectName("heroSub")
        outer.addWidget(smart_note)

        self.search.textChanged.connect(self._filter)
        self.mode.currentTextChanged.connect(self._sync_mode)
        self._sync_mode(self.mode.currentText())

    def _sync_mode(self, text):
        self.manual_panel.setEnabled(text.startswith("Manual"))

    def _filter(self, text):
        q = text.strip().casefold()
        for i in range(self.list.count()):
            item = self.list.item(i)
            item.setHidden(bool(q and q not in item.text().casefold()))

    def set_codes(self, codes):
        wanted = set(codes)
        for i in range(self.list.count()):
            item = self.list.item(i)
            code = item.data(Qt.ItemDataRole.UserRole)
            item.setCheckState(Qt.CheckState.Checked if code in wanted else Qt.CheckState.Unchecked)

    def selected_codes(self):
        out = []
        for i in range(self.list.count()):
            item = self.list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                out.append(str(item.data(Qt.ItemDataRole.UserRole)))
        return out

    def mode_text(self):
        return self.mode.currentText()


class TranslateLanguageSelector(QWidget):
    """Smart/manual source + searchable target language selector for Translate PDF."""
    def __init__(self, parent=None):
        super().__init__(parent)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(8)

        self.source_mode = QComboBox()
        self.source_mode.addItems(["Smart / Auto detect", "Manual source language"])
        outer.addWidget(self.source_mode)

        self.source = QComboBox()
        self.target = QComboBox()
        for combo in (self.source, self.target):
            combo.setEditable(True)
            combo.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
            for code, name in sorted(TRANSLATE_LANGUAGE_NAMES.items(), key=lambda kv: kv[1].casefold()):
                combo.addItem(f"{name}   ({code})", code)
            if combo.completer() is not None:
                combo.completer().setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                combo.completer().setFilterMode(Qt.MatchFlag.MatchContains)

        self._set_combo_code(self.source, "en")
        self._set_combo_code(self.target, "el")

        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.addRow("Manual source:", self.source)
        form.addRow("Translate to:", self.target)
        outer.addLayout(form)

        quick = QHBoxLayout()
        to_gr = QPushButton("→ Greek")
        to_en = QPushButton("→ English")
        swap = QPushButton("Swap")
        to_gr.clicked.connect(lambda: self._set_combo_code(self.target, "el"))
        to_en.clicked.connect(lambda: self._set_combo_code(self.target, "en"))
        swap.clicked.connect(self._swap)
        quick.addWidget(to_gr); quick.addWidget(to_en); quick.addWidget(swap); quick.addStretch()
        outer.addLayout(quick)

        self.ocr_scans = QCheckBox("Automatically OCR scanned / textless pages before translation")
        self.ocr_scans.setChecked(True)
        outer.addWidget(self.ocr_scans)

        note = QLabel(
            "Smart source uses the translation provider's automatic language detection. "
            "For scanned pages PDF Mule first performs Smart OCR, then translates the extracted text. "
            "Translation uses GoogleTranslator via deep-translator and therefore requires Internet access."
        )
        note.setWordWrap(True)
        note.setObjectName("heroSub")
        outer.addWidget(note)

        self.source_mode.currentTextChanged.connect(self._sync_source)
        self._sync_source(self.source_mode.currentText())

    def _set_combo_code(self, combo, code):
        for i in range(combo.count()):
            if combo.itemData(i) == code:
                combo.setCurrentIndex(i)
                return

    def _combo_code(self, combo):
        data = combo.currentData()
        if data:
            return str(data)
        txt = combo.currentText().strip()
        m = re.search(r"\(([^()]+)\)\s*$", txt)
        if m and m.group(1) in TRANSLATE_LANGUAGE_NAMES:
            return m.group(1)
        if txt in TRANSLATE_LANGUAGE_NAMES:
            return txt
        return ""

    def _sync_source(self, value):
        self.source.setEnabled(value.startswith("Manual"))

    def _swap(self):
        src = self._combo_code(self.source)
        tgt = self._combo_code(self.target)
        if tgt:
            self._set_combo_code(self.source, tgt)
        if src:
            self._set_combo_code(self.target, src)
        self.source_mode.setCurrentText("Manual source language")

    def source_mode_text(self):
        return self.source_mode.currentText()

    def source_code(self):
        return self._combo_code(self.source)

    def target_code(self):
        return self._combo_code(self.target)


# ----------------------------- Tool Dialog -----------------------------
class ToolDialog(QDialog):
    def __init__(self, tool: ToolDef, settings: dict, parent=None):
        super().__init__(parent); self.tool=tool; self.settings=settings; self.worker=None; self.result_path=None; self.ocr_language_selector=None; self.translate_language_selector=None
        self.setWindowTitle(f"{tool.title} — {APP_NAME}"); self.setModal(True); self.setAcceptDrops(True)
        # Fit small laptop screens: the middle section scrolls instead of overflowing the display.
        avail = QApplication.primaryScreen().availableGeometry() if QApplication.primaryScreen() else None
        want_w, want_h = (900, 880) if tool.key in {"ocr","translate"} else (820, 760)
        if avail is not None: want_w, want_h = min(want_w, avail.width()-40), min(want_h, avail.height()-60)
        self.resize(want_w, want_h)
        root=QVBoxLayout(self); root.setContentsMargins(22,20,22,20); root.setSpacing(12)
        title=QLabel(f"{tool.icon}  {tool.title}"); title.setObjectName("heroTitle"); root.addWidget(title)
        sub=QLabel(tool.desc); sub.setObjectName("heroSub"); sub.setWordWrap(True); root.addWidget(sub)
        body=QWidget(); body_l=QVBoxLayout(body); body_l.setContentsMargins(0,0,6,0); body_l.setSpacing(12)
        body_scroll=QScrollArea(); body_scroll.setWidgetResizable(True); body_scroll.setFrameShape(QFrame.Shape.NoFrame); body_scroll.setWidget(body)

        file_box=QGroupBox("Input"); fl=QVBoxLayout(file_box)
        self.files=QListWidget(); self.files.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); self.files.setMinimumHeight(90); self.files.setMaximumHeight(150); fl.addWidget(self.files)
        drop_hint=QLabel("Tip: drag & drop files or folders here." + ("" if tool.multi else " This tool uses one file.")); drop_hint.setObjectName("heroSub"); fl.addWidget(drop_hint)
        QShortcut(QKeySequence("Delete"), self.files, activated=self.remove_selected_inputs, context=Qt.ShortcutContext.WidgetShortcut)
        btnrow=QHBoxLayout(); self.add_btn=QPushButton("＋ Add files"); self.add_folder_btn=QPushButton("＋ Add Folder"); self.rem_btn=QPushButton("Remove selected"); self.clear_btn=QPushButton("Clear")
        self.add_btn.clicked.connect(self.add_files); self.add_folder_btn.clicked.connect(self.add_folder_files); self.rem_btn.clicked.connect(self.remove_selected_inputs); self.clear_btn.clicked.connect(self.files.clear)
        self.files.setToolTip("Processing order follows this list. Drag files here, press Delete to remove the selection.")
        self.add_folder_btn.setToolTip("Add all supported files from a folder and its subfolders."); self.add_folder_btn.setVisible(bool(tool.multi))
        btnrow.addWidget(self.add_btn); btnrow.addWidget(self.add_folder_btn); btnrow.addWidget(self.rem_btn); btnrow.addWidget(self.clear_btn); btnrow.addStretch();
        self.input_badge=QLabel("0 files"); self.input_badge.setObjectName("badge"); btnrow.addWidget(self.input_badge); fl.addLayout(btnrow); self.files.model().rowsInserted.connect(lambda *_:self.update_input_badge()); self.files.model().rowsRemoved.connect(lambda *_:self.update_input_badge())
        if tool.key == "html2pdf": file_box.setVisible(False)
        body_l.addWidget(file_box)

        self.form_group=QGroupBox("Options"); self.form=QFormLayout(self.form_group); self.widgets={}; self._build_options(); body_l.addWidget(self.form_group)

        if tool.batchable:
            perf=QGroupBox("Smart batch / performance"); pf=QFormLayout(perf)
            self.profile=QComboBox(); self.profile.addItems(["Smart / Auto","Balanced","Maximum speed","Low memory","Single worker"]); self.profile.currentTextChanged.connect(self.apply_profile); pf.addRow("Processing profile:",self.profile)
            self.worker_spin=QSpinBox(); self.worker_spin.setRange(1,max(1,min(16,os.cpu_count() or 4))); self.worker_spin.setValue(min(4,os.cpu_count() or 2)); pf.addRow("Parallel file workers:",self.worker_spin)
            self.continue_errors=QCheckBox("Continue batch when one file fails"); self.continue_errors.setChecked(True); pf.addRow("",self.continue_errors)
            hint=QLabel("For 2+ inputs each PDF is isolated in its own worker process. Smart/Auto limits concurrency to reduce memory pressure."); hint.setWordWrap(True); hint.setObjectName("heroSub"); pf.addRow("",hint); body_l.addWidget(perf)
        else:self.profile=self.worker_spin=self.continue_errors=None
        body_l.addStretch(1); root.addWidget(body_scroll,1)

        self.log=QPlainTextEdit(); self.log.setReadOnly(True); self.log.setMaximumHeight(105); self.log.setMaximumBlockCount(3000); self.log.setPlaceholderText("Processing log…"); root.addWidget(self.log)
        self.progress=QProgressBar(); self.progress.setRange(0,100); self.progress.setValue(0); root.addWidget(self.progress)
        row=QHBoxLayout(); self.open_btn=QPushButton("Open output folder"); self.open_btn.setEnabled(False); self.open_btn.clicked.connect(self.open_result); self.cancel_btn=QPushButton("Cancel job"); self.cancel_btn.setEnabled(False); self.cancel_btn.clicked.connect(self.cancel_job)
        row.addWidget(self.open_btn); row.addWidget(self.cancel_btn); row.addStretch(); close=QPushButton("Close"); close.clicked.connect(self.request_close); self.run_btn=QPushButton(f"Run {tool.title}"); self.run_btn.setObjectName("primary"); self.run_btn.clicked.connect(self.run_tool); row.addWidget(close); row.addWidget(self.run_btn); root.addLayout(row)

    def update_input_badge(self):
        n=self.files.count(); self.input_badge.setText(f"{n} file(s)")
        if hasattr(self,"run_btn"):
            self.run_btn.setText(f"Run {self.tool.title}" + (f"  ·  {n} files" if n>1 and self.tool.batchable else ""))
    def _add_paths(self,paths):
        """Add files (and folders, scanned recursively) honouring this tool's file filter."""
        suffixes=filter_suffixes(self.tool.extensions)
        found=[]
        for pth in paths:
            P=Path(pth)
            if P.is_dir():
                if self.tool.multi: found.extend(collect_folder_files(str(P),self.tool.extensions,True))
            elif P.is_file() and (not suffixes or P.suffix.lower() in suffixes):
                found.append(str(P))
        if not found: return 0
        if not self.tool.multi:
            self.files.clear(); found=found[:1]
        existing={self.files.item(i).text() for i in range(self.files.count())}
        added=0
        for f in found:
            if f not in existing: self.files.addItem(f); existing.add(f); added+=1
        self.update_input_badge(); return added
    def dragEnterEvent(self,e):
        if e.mimeData().hasUrls() and (self.worker is None): e.acceptProposedAction()
    def dropEvent(self,e):
        paths=[u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        added=self._add_paths(paths)
        if not added: self.log.appendPlainText(f"[INPUT] Nothing added — expected {self.tool.extensions}")
        e.acceptProposedAction()
    def remove_selected_inputs(self):
        for i in sorted((self.files.row(x) for x in self.files.selectedItems()),reverse=True): self.files.takeItem(i)
    def line(self,key,label,placeholder="",password=False):
        w=QLineEdit(); w.setPlaceholderText(placeholder); w.setEchoMode(QLineEdit.EchoMode.Password if password else QLineEdit.EchoMode.Normal); self.widgets[key]=w; self.form.addRow(label,w); return w
    def combo(self,key,label,items):
        w=QComboBox(); w.addItems(items); self.widgets[key]=w; self.form.addRow(label,w); return w
    def spin(self,key,label,mn,mx,val):
        w=QSpinBox(); w.setRange(mn,mx); w.setValue(val); self.widgets[key]=w; self.form.addRow(label,w); return w
    def dspin(self,key,label,mn,mx,val):
        w=QDoubleSpinBox(); w.setRange(mn,mx); w.setValue(val); w.setDecimals(1); self.widgets[key]=w; self.form.addRow(label,w); return w
    def check(self,key,label,val=False):
        w=QCheckBox(label); w.setChecked(val); self.widgets[key]=w; self.form.addRow("",w); return w
    def _build_options(self):
        k=self.tool.key
        if k in {"remove","extract"}:
            self.line("pages","Pages / sequence", "Example: 1-3,5,8  ·  10-end  ·  -3")
            self._note("Ranges are 1-based. 'end' means the last page, so 5-end works for any length.")
        if k in {"scan","img2pdf"}:
            self.combo("page_size","Page size",["Image size (DPI)","A4 (fit image)","Letter (fit image)","A3 (fit image)"])
            self.dspin("margin_mm","Margin (mm)",0,50,0)
            self._note("Multi-page TIFFs keep every page; phone photos are auto-rotated using EXIF orientation.")
        if k=="compress": self.combo("compression_profile","Compression",["Smart / Auto","Lossless","Balanced (Ghostscript if available)","Strong (Ghostscript if available)"])
        if k=="ocr":
            self.ocr_language_selector=OcrLanguageSelector(self)
            self.form.addRow("OCR language mode:", self.ocr_language_selector)
            self.spin("smart_threshold","Smart script threshold (%)",50,95,70)
            self.spin("dpi","OCR render DPI",150,600,300)
            self.check("skip_text_pages","Skip pages that already contain text (keeps them untouched)",True)
            self.check("auto_rotate","Auto-straighten upside-down / sideways scans (Tesseract OSD)",True)
            self._note("The output keeps the original page images and adds an invisible, searchable text layer.")
        if k=="html2pdf": self.line("source","URL or HTML path","https://example.com or C:\\page.html")
        if k=="pdf2jpg": self.spin("dpi","DPI",72,600,180); self.spin("quality","JPG quality",40,100,90)
        if k=="pdfa": self.combo("pdfa","PDF/A level",["1","2","3"])
        if k=="rotate":
            self.combo("angle","Rotation",["90° clockwise","180°","270° (90° counter-clockwise)"])
            self.line("pages","Pages (blank = all)","Example: 2,4-6")
        if k=="numbers":
            self.combo("number_format","Format",["1","Page 1","1 / N","Page 1 of N","Σελίδα 1 από N"])
            self.spin("start","Start number",0,99999,1); self.spin("font_size","Font size",6,48,11)
            self.combo("position","Position",["Bottom center","Bottom right","Bottom left","Top center","Top right","Top left"])
            self.check("skip_first","Don't number the first page (cover)",False)
        if k=="watermark":
            self.line("text","Watermark text","CONFIDENTIAL").setText("CONFIDENTIAL"); self.combo("layout","Layout",["Diagonal","Horizontal"]); self.spin("font_size","Max font size",10,200,60); self.dspin("opacity","Opacity",0.05,1.0,0.18).setSingleStep(0.05); self.widgets["opacity"].setDecimals(2)
            img=self.line("image","Optional image", "Leave blank for text"); b=QPushButton("Browse image…"); b.clicked.connect(lambda:self._pick_to_line(img,"Images (*.png *.jpg *.jpeg)")); self.form.addRow("",b)
        if k=="crop":
            for key,label in [("left","Left mm"),("right","Right mm"),("top","Top mm"),("bottom","Bottom mm")]: self.dspin(key,label,0,100,0)
        if k=="edit":
            self.line("text","Text","Text to place (Greek and other scripts supported)"); self.spin("page","Page",1,99999,1); self.check("all_pages","Apply to all pages",False); self.dspin("x","X from left (mm)",0,500,20); self.dspin("y","Y from top (mm)",0,500,20); self.spin("font_size","Font size",6,96,18)
        if k in {"unlock","protect"}: self.line("password","Password","PDF password",True)
        if k=="protect":
            self.line("password_confirm","Confirm password","Type the password again",True)
            self.line("owner_password","Owner password (optional)","Defaults to user password",True)
            self.check("allow_copy","Allow copying text",True)
        if k=="sign":
            img=self.line("image","Signature image","PNG/JPG signature (transparent PNG looks best)"); b=QPushButton("Browse signature…"); b.clicked.connect(lambda:self._pick_to_line(img,"Images (*.png *.jpg *.jpeg)")); self.form.addRow("",b); self.dspin("width","Signature width (mm)",10,120,45)
            self.combo("sign_pages","Place on",["Last page","First page","All pages"]); self.combo("position","Position",["Bottom right","Bottom left","Bottom center","Top right","Top left"])
        if k=="redact":
            self.line("text","Text to redact","e.g. John Smith; 6912345678; IBAN")
            self._note("Separate several terms with ';'. Matching is case-insensitive. Redaction is permanent: the text is removed, not just covered.")
        if k=="summarize": self.spin("sentences","Summary sentences",3,40,10)
        if k=="translate":
            self.translate_language_selector=TranslateLanguageSelector(self)
            self.form.addRow("Translation language:", self.translate_language_selector)
        if k=="workflow": self.line("password","Optional output password","Leave blank for no encryption",True)
        if self.form.rowCount()==0:self.form.addRow(QLabel("No extra options required."))
    def _note(self,text):
        n=QLabel(text); n.setWordWrap(True); n.setObjectName("heroSub"); self.form.addRow("",n); return n
    def _pick_to_line(self,line,flt):
        f,_=QFileDialog.getOpenFileName(self,"Choose file",self.settings.get("last_dir",str(Path.home())),flt)
        if f: line.setText(f)
    def add_files(self):
        start=self.settings.get("last_dir",str(Path.home()))
        if self.tool.multi: files,_=QFileDialog.getOpenFileNames(self,"Select input files",start,self.tool.extensions)
        else:
            f,_=QFileDialog.getOpenFileName(self,"Select input file",start,self.tool.extensions); files=[f] if f else []
        if files:
            self._add_paths(files)
            self.settings["last_dir"]=str(Path(files[0]).parent);save_settings(self.settings)
    def add_folder_files(self):
        if not self.tool.multi:
            return
        start=self.settings.get("last_dir",str(Path.home()))
        folder=QFileDialog.getExistingDirectory(self,"Add folder",start)
        if not folder:
            return
        files=collect_folder_files(folder,self.tool.extensions,recursive=True)
        self.settings["last_dir"]=folder; save_settings(self.settings)
        if not files:
            QMessageBox.information(self,"Add Folder",f"No supported files were found in this folder or its subfolders.\n\nExpected: {self.tool.extensions}")
            return
        existing={self.files.item(i).text() for i in range(self.files.count())}
        added=0
        for f in files:
            if f not in existing:
                self.files.addItem(f); existing.add(f); added+=1
        self.update_input_badge()
        self.log.appendPlainText(f"[INPUT] Folder scan: {len(files)} matching file(s), {added} new file(s) added from {folder}")
    def apply_profile(self,text):
        if not self.worker_spin:return
        mx=self.worker_spin.maximum()
        self.worker_spin.setValue(1 if text in {"Low memory","Single worker"} else min(2,mx) if text=="Balanced" else mx if text=="Maximum speed" else min(4,mx))
    def options(self):
        out={}
        for k,w in self.widgets.items():
            if isinstance(w,QLineEdit):out[k]=w.text()
            elif isinstance(w,QComboBox):out[k]=w.currentText()
            elif isinstance(w,QSpinBox):out[k]=w.value()
            elif isinstance(w,QDoubleSpinBox):out[k]=w.value()
            elif isinstance(w,QCheckBox):out[k]=w.isChecked()
        if self.tool.key=="ocr" and self.ocr_language_selector is not None:
            codes=self.ocr_language_selector.selected_codes()
            out["ocr_mode"]=self.ocr_language_selector.mode_text()
            out["lang"]="+".join(codes)
        if self.tool.key=="translate" and self.translate_language_selector is not None:
            out["translate_source_mode"]=self.translate_language_selector.source_mode_text()
            out["translate_source"]=self.translate_language_selector.source_code()
            out["translate_target"]=self.translate_language_selector.target_code()
            out["translate_ocr_scans"]=self.translate_language_selector.ocr_scans.isChecked()
        if self.tool.batchable:
            out.update({"workers":self.worker_spin.value(),"performance_profile":self.profile.currentText(),"continue_on_error":self.continue_errors.isChecked(),"_engine_key":self.tool.key})
        return out
    def _choose_output(self,files):
        k=self.tool.key;start=self.settings.get("last_dir",str(Path.home()))
        if self.tool.batchable and len(files)>1:return QFileDialog.getExistingDirectory(self,"Select batch output folder",start)
        dirs={"pdf2jpg","word2pdf","ppt2pdf","xls2pdf","extract_text"}
        if k in dirs:return QFileDialog.getExistingDirectory(self,"Select output folder",start)
        ext=".pdf";flt="PDF files (*.pdf)"
        if k=="pdf2word":ext=".docx";flt="Word document (*.docx)"
        elif k=="pdf2ppt":ext=".pptx";flt="PowerPoint (*.pptx)"
        elif k=="pdf2xls":ext=".xlsx";flt="Excel workbook (*.xlsx)"
        elif k in {"forms","summarize","translate"}:ext=".txt";flt="Text files (*.txt)"
        elif k=="compare":ext=".html";flt="HTML report (*.html)"
        elif k=="markdown":ext=".md";flt="Markdown (*.md)"
        if files: start=self.settings.get("last_out_dir") or str(Path(files[0]).parent)
        base=Path(files[0]).stem if files else "output"; default=Path(start)/(base+"_"+k+ext)
        while True:
            f,_=QFileDialog.getSaveFileName(self,"Save output",str(default),flt)
            if not f: return f
            if not Path(f).suffix: f+=ext
            if any(_same_file(f,x) for x in files):
                QMessageBox.warning(self,"Save output","Choose a different file name — the input file cannot be overwritten while it is being read.")
                continue
            return f
    def _lock_ui(self,busy):
        self.run_btn.setEnabled(not busy);self.cancel_btn.setEnabled(busy);self.add_btn.setEnabled(not busy);self.add_folder_btn.setEnabled(not busy);self.rem_btn.setEnabled(not busy);self.clear_btn.setEnabled(not busy);self.files.setEnabled(not busy);self.form_group.setEnabled(not busy)
        if self.profile:self.profile.setEnabled(not busy);self.worker_spin.setEnabled(not busy);self.continue_errors.setEnabled(not busy)
    def run_tool(self):
        files=[self.files.item(i).text() for i in range(self.files.count())]
        if self.tool.key!="html2pdf" and not files:QMessageBox.warning(self,"No input","Add at least one input file first.");return
        if self.tool.key=="compare" and len(files)!=2:QMessageBox.warning(self,"Compare PDF","Select exactly two PDF files.");return
        if self.tool.key=="ocr" and self.ocr_language_selector is not None and self.ocr_language_selector.mode_text().startswith("Manual") and not self.ocr_language_selector.selected_codes():
            QMessageBox.warning(self,"OCR languages","Select at least one language, or switch to Smart / Auto detect.");return
        if self.tool.key=="translate" and self.translate_language_selector is not None:
            target=self.translate_language_selector.target_code()
            if not target:
                QMessageBox.warning(self,"Translate PDF","Choose a valid target language from the list.");return
            if self.translate_language_selector.source_mode_text().startswith("Manual"):
                source=self.translate_language_selector.source_code()
                if not source:
                    QMessageBox.warning(self,"Translate PDF","Choose a valid manual source language from the list.");return
                if source==target:
                    QMessageBox.warning(self,"Translate PDF","Source and target language are the same.");return
        if not self.tool.multi and len(files)>1:files=files[:1]
        missing=[f for f in files if not Path(f).exists()]
        if missing:
            QMessageBox.warning(self,"Missing input",f"{len(missing)} input file(s) no longer exist:\n\n"+"\n".join(missing[:8])); return
        passwords=self._collect_passwords(files)
        if passwords is None: return
        out=self._choose_output(files)
        if not out:return
        self.settings["last_out_dir"]=out if Path(out).is_dir() else str(Path(out).parent); save_settings(self.settings)
        self.result_path=None;self.open_btn.setEnabled(False);self._lock_ui(True);self.progress.setValue(0);self.log.clear();self.log.appendPlainText(f"Starting… {len(files)} input file(s)")
        opts=self.options(); opts["passwords"]=passwords; self._t0=time.monotonic()
        if self.tool.key=="pdf2word" and files: opts["input_password"]=passwords.get(files[0],"")
        fn=eng_batch if self.tool.batchable and len(files)>1 else ENGINE[self.tool.key]
        self.worker=Worker(fn,files,out,opts);self.worker.progress.connect(self.on_progress);self.worker.finished_ok.connect(self.on_ok);self.worker.failed.connect(self.on_fail);self.worker.cancelled.connect(self.on_cancelled);self.worker.start()
    def _collect_passwords(self,files):
        """Ask once per encrypted PDF before starting (the last good password is tried first)."""
        if "*.pdf" not in self.tool.extensions or self.tool.key=="unlock": return {}
        out={}; last=""
        for f in files:
            try:
                d=fitz.open(f)
            except Exception:
                continue  # the engine reports unreadable files per item
            try:
                if not d.needs_pass: continue
                if last and d.authenticate(last): out[f]=last; continue
                while True:
                    pw,ok=QInputDialog.getText(self,"Encrypted PDF",f"{Path(f).name} is password-protected.\nPassword:",QLineEdit.EchoMode.Password)
                    if not ok: return None
                    if d.authenticate(pw): out[f]=pw; last=pw; break
                    QMessageBox.warning(self,"Encrypted PDF","Wrong password, try again.")
            finally:
                d.close()
        return out
    def on_progress(self,p,msg):
        self.progress.setValue(p)
        if msg: self.log.appendPlainText(msg)
    def on_ok(self,result):
        self.result_path=str(result);self.worker=None;self.progress.setValue(100)
        took=fmt_elapsed(time.monotonic()-getattr(self,"_t0",time.monotonic()))
        self.log.appendPlainText(f"✓ Completed in {took}");self._lock_ui(False);self.open_btn.setEnabled(True)
        add_recent(self.settings,self.tool.title,self.result_path)
        box=QMessageBox(self); box.setIcon(QMessageBox.Icon.Information); box.setWindowTitle("Completed")
        box.setText(f"{self.tool.title} finished in {took}.\n\nOutput:\n{self.result_path}")
        open_folder_btn=box.addButton("Open folder",QMessageBox.ButtonRole.ActionRole)
        open_file_btn=box.addButton("Open file",QMessageBox.ButtonRole.ActionRole) if Path(self.result_path).is_file() else None
        box.addButton(QMessageBox.StandardButton.Ok); box.exec()
        if box.clickedButton()==open_folder_btn: open_folder(self.result_path)
        elif open_file_btn is not None and box.clickedButton()==open_file_btn: QDesktopServices.openUrl(QUrl.fromLocalFile(self.result_path))
    def on_fail(self,error):
        self.worker=None;self._lock_ui(False);self.progress.setValue(0)
        self.log.appendPlainText("✗ Failed: "+friendly_error(error)); self.log.appendPlainText(error.strip())
        QMessageBox.critical(self,f"{self.tool.title} failed",friendly_error(error))
    def on_cancelled(self):self.worker=None;self._lock_ui(False);self.progress.setValue(0);self.log.appendPlainText("■ Cancellation requested / job stopped safely.")
    def cancel_job(self):
        if self.worker is not None and self.worker.isRunning():self.worker.requestInterruption();self.cancel_btn.setEnabled(False);self.log.appendPlainText("Cancelling… (a single long step such as a Word conversion finishes first)")
    def request_close(self):
        self.reject()
    def reject(self):
        if not stop_threads_for_close(self,[self.worker],f"{self.tool.title}"): return
        super().reject()
    def open_result(self):
        if self.result_path:open_folder(self.result_path)


# ----------------------------- Cards & Main UI -----------------------------
class ToolCard(QFrame):
    clicked = pyqtSignal(object)
    def __init__(self,tool):
        super().__init__(); self.tool=tool; self.setObjectName("toolCard"); self.setCursor(Qt.CursorShape.PointingHandCursor); self.setMinimumSize(210,155); self.setMaximumHeight(175); self.setToolTip(f"{tool.title}\n{tool.desc}")
        lay=QVBoxLayout(self); lay.setContentsMargins(15,14,15,14); lay.setSpacing(8)
        row=QHBoxLayout(); icon=QLabel(tool.icon); icon.setObjectName("toolIcon"); icon.setFixedSize(46,46); icon.setAlignment(Qt.AlignmentFlag.AlignCenter); row.addWidget(icon); row.addStretch()
        cat=QLabel(tool.category); cat.setStyleSheet("color:#7f8998;font-size:8.5pt;"); row.addWidget(cat); lay.addLayout(row)
        t=QLabel(tool.title); t.setObjectName("toolTitle"); lay.addWidget(t)
        d=QLabel(tool.desc); d.setObjectName("toolDesc"); d.setWordWrap(True); lay.addWidget(d); lay.addStretch()
    def mousePressEvent(self,e):
        if e.button()==Qt.MouseButton.LeftButton: self.clicked.emit(self.tool)
        super().mousePressEvent(e)


class DependencyStatusDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Dependency status — {APP_NAME}")
        self.setWindowIcon(load_app_icon())
        self.resize(850, 650)
        self.setMinimumSize(720, 540)
        self.scan_worker = None
        self.proc = None
        self.last_scan = None

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        root.setSpacing(12)

        head = QHBoxLayout()
        titles = QVBoxLayout()
        t = QLabel("Dependency status")
        t.setStyleSheet("font-size:18pt;font-weight:900;")
        s = QLabel("Fast background scan — modules are detected without importing heavy libraries.")
        s.setObjectName("heroSub")
        titles.addWidget(t); titles.addWidget(s)
        head.addLayout(titles, 1)
        self.scan_btn = QPushButton("↻  Refresh")
        self.scan_btn.clicked.connect(self.start_scan)
        head.addWidget(self.scan_btn)
        root.addLayout(head)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["Component", "Type", "Status"])
        self.table.verticalHeader().setVisible(False)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        root.addWidget(self.table, 1)

        status_row = QHBoxLayout()
        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("heroSub")
        status_row.addWidget(self.status_label, 1)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFixedWidth(250)
        status_row.addWidget(self.progress)
        root.addLayout(status_row)

        log_title = QLabel("Terminal log")
        log_title.setStyleSheet("font-weight:800;")
        root.addWidget(log_title)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(5000)
        self.log.setMinimumHeight(180)
        self.log.setStyleSheet(
            "font-family:Consolas,'Courier New',monospace;font-size:9.5pt;"
        )
        root.addWidget(self.log, 1)

        buttons = QHBoxLayout()
        self.install_btn = QPushButton("Install / upgrade Python add-ons")
        self.install_btn.setObjectName("primary")
        self.install_btn.clicked.connect(self.install_addons)
        buttons.addWidget(self.install_btn)
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_process)
        buttons.addWidget(self.stop_btn)
        buttons.addStretch(1)
        self.close_btn = QPushButton("Close")
        self.close_btn.clicked.connect(self.close)
        buttons.addWidget(self.close_btn)
        root.addLayout(buttons)

        if getattr(sys, "frozen", False):
            self.install_btn.setEnabled(False)
            self.install_btn.setToolTip("This is a packaged EXE: Python add-ons must be bundled at build time (see BUILD_EXE.txt).")
            self.append_log("[INFO] Packaged EXE build: pip installs are disabled. External programs (Tesseract, "
                            "LibreOffice, Ghostscript) are still detected and used when installed.")
        self.start_scan()

    def append_log(self, message):
        text = str(message).rstrip()
        if text:
            self.log.appendPlainText(text)
            bar = self.log.verticalScrollBar()
            bar.setValue(bar.maximum())

    def start_scan(self):
        if self.scan_worker and self.scan_worker.isRunning():
            return
        self.scan_btn.setEnabled(False)
        self.install_btn.setEnabled(False)
        self.status_label.setText("Scanning dependencies…")
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.append_log("[SCAN] Starting non-blocking dependency scan…")
        self.scan_worker = Worker(scan_optional_dependencies)
        self.scan_worker.progress.connect(self.on_scan_progress)
        self.scan_worker.finished_ok.connect(self.on_scan_done)
        self.scan_worker.failed.connect(self.on_scan_failed)
        self.scan_worker.start()

    def on_scan_progress(self, value, message):
        self.progress.setValue(value)
        self.status_label.setText(message)

    def on_scan_done(self, result):
        self.last_scan = result
        self.table.setRowCount(0)
        missing = 0
        for item in result.get("python", []):
            row = self.table.rowCount(); self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(item["label"]))
            self.table.setItem(row, 1, QTableWidgetItem("Python"))
            status = "✓ Installed" if item["found"] else "✗ Missing"
            if not item["found"]: missing += 1
            cell = QTableWidgetItem(f"{status}  —  {item['detail']}")
            self.table.setItem(row, 2, cell)
            self.append_log(f"[SCAN] {item['label']}: {status}")
        for item in result.get("external", []):
            row = self.table.rowCount(); self.table.insertRow(row)
            self.table.setItem(row, 0, QTableWidgetItem(item["label"]))
            self.table.setItem(row, 1, QTableWidgetItem("External app"))
            status = "✓ Found" if item["found"] else "✗ Not found"
            self.table.setItem(row, 2, QTableWidgetItem(f"{status}  —  {item['detail']}"))
            self.append_log(f"[SCAN] {item['label']}: {status} {item['detail'] if item['found'] else ''}")
        self.progress.setValue(100)
        self.status_label.setText(f"Scan complete — {missing} Python add-on(s) missing")
        self.append_log(f"[SCAN] Complete. Missing Python add-ons: {missing}")
        self.scan_btn.setEnabled(True)
        self.install_btn.setEnabled(not getattr(sys, "frozen", False))
        _find_exe_cached.cache_clear()

    def on_scan_failed(self, error):
        self.status_label.setText("Dependency scan failed")
        self.append_log("[SCAN ERROR] " + str(error))
        self.scan_btn.setEnabled(True)
        self.install_btn.setEnabled(True)
        QMessageBox.critical(self, "Dependency scan", str(error))

    def packages_to_install(self):
        if not self.last_scan:
            return [x[2] for x in PYTHON_OPTIONAL_DEPS]
        missing = [x["package"] for x in self.last_scan.get("python", []) if not x["found"]]
        # If everything is present, the button acts as an upgrade-all action.
        return missing or [x[2] for x in PYTHON_OPTIONAL_DEPS]

    def install_addons(self):
        if self.proc and self.proc.state() != QProcess.ProcessState.NotRunning:
            QMessageBox.information(self, "Installer running", "An installation is already running.")
            return
        packages = self.packages_to_install()
        action = "install the missing" if self.last_scan and any(not x["found"] for x in self.last_scan.get("python", [])) else "upgrade the"
        ans = QMessageBox.question(
            self,
            "Python add-ons",
            f"Run pip to {action} Python add-ons?\n\n" + "\n".join(packages)
        )
        if ans != QMessageBox.StandardButton.Yes:
            return

        self.log.clear()
        py = python_console_executable()
        self.append_log(f"$ {py} -m pip install --upgrade " + " ".join(packages))
        self.append_log("[INSTALL] Starting pip…")
        self.status_label.setText("Installing Python add-ons…")
        self.progress.setRange(0, 0)  # indeterminate while pip works
        self.install_btn.setEnabled(False)
        self.scan_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.close_btn.setEnabled(False)

        self.proc = QProcess(self)
        self.proc.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self.proc.readyReadStandardOutput.connect(self.read_process_output)
        self.proc.started.connect(lambda: self.append_log("[INSTALL] pip process started."))
        self.proc.errorOccurred.connect(self.process_error)
        self.proc.finished.connect(self.process_finished)
        args = ["-m", "pip", "install", "--upgrade", "--disable-pip-version-check", *packages]
        self.proc.start(py, args)

    def read_process_output(self):
        if not self.proc:
            return
        data = bytes(self.proc.readAllStandardOutput()).decode(errors="replace")
        if data:
            for line in data.splitlines():
                self.append_log(line)

    def process_error(self, error):
        self.append_log(f"[PROCESS ERROR] {error}")

    def process_finished(self, exit_code, exit_status):
        self.read_process_output()
        success = (exit_code == 0 and exit_status == QProcess.ExitStatus.NormalExit)
        self.progress.setRange(0, 100)
        self.progress.setValue(100 if success else 0)
        self.stop_btn.setEnabled(False)
        self.close_btn.setEnabled(True)
        self.scan_btn.setEnabled(True)
        self.install_btn.setEnabled(True)

        if success:
            self.status_label.setText("Installation completed successfully ✓")
            self.append_log("[OK] Installation completed successfully.")
            QMessageBox.information(
                self,
                "Installation complete",
                "Python add-ons were installed successfully.\n\nDependency status will now be refreshed."
            )
            self.start_scan()
        else:
            self.status_label.setText(f"Installation failed — exit code {exit_code}")
            self.append_log(f"[FAILED] pip exited with code {exit_code}.")
            QMessageBox.warning(
                self,
                "Installation failed",
                f"pip finished with exit code {exit_code}.\n\nCheck the terminal log for the exact error."
            )

    def stop_process(self):
        if not self.proc or self.proc.state() == QProcess.ProcessState.NotRunning:
            return
        ans = QMessageBox.question(self, "Stop installer", "Stop the running pip installation?")
        if ans == QMessageBox.StandardButton.Yes:
            self.append_log("[STOP] Terminating pip…")
            self.proc.terminate()

    def closeEvent(self, event):
        if self.proc and self.proc.state() != QProcess.ProcessState.NotRunning:
            ans = QMessageBox.question(
                self, "Installer running", "pip is still running. Stop it and close?"
            )
            if ans != QMessageBox.StandardButton.Yes:
                event.ignore(); return
            self.proc.kill()
        if self.scan_worker and self.scan_worker.isRunning():
            self.scan_worker.requestInterruption()
            self.scan_worker.wait(800)
        event.accept()



# ----------------------------------------------------------------------------------------------
# App updates (OTA): the newest PDF Mule from GitHub, saved next to this one and started in its place
# ----------------------------------------------------------------------------------------------
APP_PAGE = "https://github.com/nikkpap/PDFMule"
APP_UPDATE_URL = "https://raw.githubusercontent.com/nikkpap/PDFMule/main/PDFMule.pyw"
APP_CHANGES_URL = "https://raw.githubusercontent.com/nikkpap/PDFMule/main/CHANGELOG.md"


def script_dir() -> Path:
    """Folder of the running .pyw (new versions are saved next to it)."""
    return Path(os.path.abspath(sys.argv[0] if sys.argv and sys.argv[0] else __file__)).resolve().parent


def _gui_python() -> str:
    """pythonw.exe next to the running interpreter: the same Python, without a console window."""
    exe = Path(sys.executable)
    if os.name == "nt" and exe.name.lower() in ("python.exe", "py.exe"):
        w = exe.with_name("pythonw.exe")
        if w.exists():
            return str(w)
    return str(exe)


def _http(url, timeout=15):
    import urllib.request
    req = urllib.request.Request(url, headers={"User-Agent": f"PDFMule/{APP_VERSION}", "Cache-Control": "no-cache"})
    return urllib.request.urlopen(req, timeout=timeout)


def _ver(v) -> tuple:
    try:
        return tuple(int(x) for x in str(v).strip().lstrip("vV").split("."))
    except ValueError:
        return (0,)


def app_update_check(progress=None):
    """(version, text, notes) of a newer PDF Mule on GitHub, or None when this one is the newest."""
    import urllib.error
    try:
        with _http(APP_UPDATE_URL, timeout=25) as r:
            text = r.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise RuntimeError("the update file was not found on GitHub (404). The repository must be public "
                               f"for updates to work: {APP_PAGE}")
        raise RuntimeError(f"GitHub answered HTTP {e.code}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"no connection to GitHub ({e.reason})")
    m = re.search(r'^APP_VERSION = "([0-9.]+)"', text, re.M)
    if not m:
        raise RuntimeError("the file on GitHub has no version")
    ver = m.group(1)
    if _ver(ver) <= _ver(APP_VERSION):
        return None
    compile(text, f"PDFMule_v{ver}.pyw", "exec")   # a download cut short is never started
    notes = ""
    try:
        with _http(APP_CHANGES_URL, timeout=10) as r:
            ch = r.read().decode("utf-8")
        parts = re.split(r"^## ", ch, flags=re.M)
        news = [x for x in parts[1:] if x.split() and _ver(x.split()[0]) > _ver(APP_VERSION)]
        # Only the heading and the bullet lines, kept short for the dialog.
        out = []
        for x in news[:4]:
            lines = [l.strip() for l in x.strip().splitlines()]
            head = lines[0]
            bullets = [l for l in lines[1:] if l.startswith("- ")][:6]
            out.append(head + "\n" + "\n".join(bullets))
        notes = "\n\n".join(out)
    except Exception:
        pass
    return ver, text, notes


def app_update_install(ver: str, text: str) -> Path:
    """Saves PDFMule_v<ver>.pyw next to this file and starts it; the newer copy makes this one close."""
    dest = script_dir() / f"PDFMule_v{ver}.pyw"
    tmp = dest.with_suffix(".pyw.part")
    tmp.write_text(text, encoding="utf-8", newline="\n")
    os.replace(tmp, dest)
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([_gui_python(), str(dest)], cwd=str(dest.parent), close_fds=True,
                     creationflags=flags if os.name == "nt" else 0)
    return dest


def offer_app_update(parent, res, quiet=False):
    """Asks to update when res names a newer version; with quiet, says nothing when there is none."""
    if not res:
        if not quiet:
            QMessageBox.information(parent, "No update", f"{APP_NAME} {APP_VERSION} is the newest version.")
        return
    ver, text, notes = res
    esc = lambda t: t.replace("&", "&amp;").replace("<", "&lt;").replace("\n", "<br>")
    box = QMessageBox(parent)
    box.setWindowTitle(f"Update {APP_NAME}?")
    box.setIcon(QMessageBox.Icon.Information)
    box.setTextFormat(Qt.TextFormat.RichText)
    box.setText(f"<b>{APP_NAME} {ver}</b> is out (you have {APP_VERSION}).<br><br>"
                + (f"<small>{esc(notes)}</small><br><br>" if notes else "")
                + f"Update now? It is saved next to this file as <b>PDFMule_v{ver}.pyw</b> and opens in place of "
                  "this window. Your settings and recent outputs stay; the old file stays too, in case you want it back.")
    go = box.addButton("Update now", QMessageBox.ButtonRole.AcceptRole)
    box.addButton("Later", QMessageBox.ButtonRole.RejectRole)
    box.exec()
    if box.clickedButton() is not go:
        return
    try:
        app_update_install(ver, text)
    except Exception as e:
        QMessageBox.critical(parent, "Update failed", f"Could not save the new version: {e}")
        return
    QTimer.singleShot(10000, QApplication.quit)   # in case the new copy cannot ask this one to close


def auto_install_addons(win):
    """First start of each version: quietly install missing optional Python add-ons (OCR, Word, Excel,
    PowerPoint, Markdown, translation, HTML) in the background. Nothing to click; Dependency status
    shows the result. External programs (Tesseract, LibreOffice, Ghostscript) are not pip packages."""
    if getattr(sys, "frozen", False) or not win.settings.get("addons_auto", True):
        return
    if win.settings.get("addons_tried") == APP_VERSION:
        return
    missing = [pkg for _, mod, pkg in PYTHON_OPTIONAL_DEPS if importlib.util.find_spec(mod) is None]
    win.settings["addons_tried"] = APP_VERSION; save_settings(win.settings)
    if not missing:
        return
    w = Worker(lambda progress: _pip_install(missing))
    win._addons_worker = w
    w.finished_ok.connect(lambda r: importlib.invalidate_caches())
    w.start()


def ota_supported() -> bool:
    """A packaged EXE cannot replace itself with a .pyw; it points to the GitHub page instead."""
    return not getattr(sys, "frozen", False)


def run_update_check(parent, quiet: bool):
    """Background check (never blocks the window); the result is shown on the GUI thread."""
    if not ota_supported():
        if not quiet:
            QMessageBox.information(parent, "Updates", f"This is a packaged EXE. Download new versions from\n{APP_PAGE}")
        return
    w = Worker(app_update_check)
    parent._update_worker = w   # keep a reference while it runs
    w.finished_ok.connect(lambda r: offer_app_update(parent, r, quiet))
    if not quiet:
        w.failed.connect(lambda e: QMessageBox.warning(parent, "Updates", "Could not reach GitHub:\n" + friendly_error(e)))
    w.start()


# ---- one running copy: a newer start takes over, a second start of the same version brings the window up ----
_INSTANCE = None


def _instance_name() -> str:
    import getpass
    try:
        user = getpass.getuser()
    except Exception:
        user = "user"
    return "PDFMule-" + re.sub(r"[^A-Za-z0-9]", "_", user)


def single_instance() -> bool:
    """True when this copy should run. A running OLDER copy steps aside ("yield") for a newer one;
    the same or a newer one comes to the front and this start quits. A copy busy with a job answers
    "busy" and keeps running."""
    global _INSTANCE
    from PyQt6.QtNetwork import QLocalServer, QLocalSocket
    name = _instance_name()
    probe = QLocalSocket()
    probe.connectToServer(name)
    if probe.waitForConnected(500):
        probe.write(f"show {APP_VERSION}\n".encode())
        probe.flush()
        probe.waitForBytesWritten(500)
        answer = bytes(probe.readAll()) if probe.waitForReadyRead(1500) else b""
        probe.disconnectFromServer()
        if b"busy" in answer:
            QMessageBox.information(None, APP_NAME, f"{APP_NAME} is still running a job.\n\n"
                                    "Let it finish (or close it), then start this version again.")
            return False
        if b"yield" not in answer:
            return False
        for _ in range(60):   # up to 15 s for the older copy to close
            time.sleep(0.25)
            p2 = QLocalSocket()
            p2.connectToServer(name)
            if not p2.waitForConnected(200):
                break
            p2.disconnectFromServer()
        else:
            return False
    QLocalServer.removeServer(name)   # a socket file left by a crash (Linux/macOS)
    _INSTANCE = QLocalServer()
    _INSTANCE.listen(name)
    return True


def attach_instance_server(win):
    """Answers later starts: bring the window up, or close for a newer version when idle."""
    if _INSTANCE is None:
        return

    def handle():
        while _INSTANCE.hasPendingConnections():
            c = _INSTANCE.nextPendingConnection()
            c.disconnected.connect(c.deleteLater)
            line = bytes(c.readAll()) if (c.bytesAvailable() or c.waitForReadyRead(300)) else b""
            parts = line.decode("utf-8", "replace").split()
            newer = len(parts) >= 2 and parts[0] == "show" and _ver(parts[1]) > _ver(APP_VERSION)
            if newer:
                busy = QApplication.activeModalWidget() is not None
                c.write(b"busy\n" if busy else b"yield\n"); c.flush(); c.waitForBytesWritten(300)
                if busy:
                    win.showNormal(); win.raise_(); win.activateWindow()
                else:
                    _INSTANCE.close()
                    QTimer.singleShot(100, QApplication.quit)
                return
            c.write(b"shown\n"); c.flush()
            win.showNormal(); win.raise_(); win.activateWindow()

    _INSTANCE.newConnection.connect(handle)


class StartupSplash(QDialog):
    """Channelizer-style fake startup splash (~3 seconds)."""

    finishedStartup = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(load_app_icon())
        self.setWindowFlags(
            Qt.WindowType.Dialog
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.setModal(False)
        self.setFixedSize(650, 400)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(18, 18, 18, 18)

        card = QFrame()
        card.setObjectName("SplashCard")
        card.setStyleSheet(
            "QFrame#SplashCard {"
            "background:#151922; border:1px solid #2a3240; border-radius:20px;"
            "}"
        )
        outer.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(30, 26, 30, 24)
        lay.setSpacing(10)

        logo_box = QFrame()
        logo_box.setStyleSheet("background:#ffffff;border-radius:14px;")
        logo_lay = QVBoxLayout(logo_box)
        logo_lay.setContentsMargins(12, 8, 12, 8)
        logo = QLabel()
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        pm = QPixmap(str(app_logo_path()))
        if not pm.isNull():
            logo.setPixmap(pm.scaled(520, 178, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            logo.setText("PDF Mule")
            logo.setStyleSheet("color:#20242c;font-size:28pt;font-weight:900;")
        logo_lay.addWidget(logo)
        lay.addWidget(logo_box)

        subtitle = QLabel("Heavy-Duty PDF Toolbox")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color:#e9edf3;font-size:11pt;font-weight:700;")
        lay.addWidget(subtitle)

        self.status = QLabel("Initializing UI…")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status.setStyleSheet("color:#9aa5b5;font-size:9.5pt;")
        lay.addWidget(self.status)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(12)
        self.progress.setStyleSheet(
            "QProgressBar {background:#202632;border:1px solid #2b3443;"
            "border-radius:6px;}"
            "QProgressBar::chunk {background:#ff4d5f;border-radius:5px;}"
        )
        lay.addWidget(self.progress)

        meta = QLabel(f"v{APP_VERSION}   •   Build {BUILD_DATE}   •   ALU DEV TEAM @ 2026   •   click to skip")
        meta.setAlignment(Qt.AlignmentFlag.AlignCenter)
        meta.setStyleSheet("color:#707b8d;font-size:8.5pt;")
        lay.addWidget(meta)

        self._value = 0
        self._timer = QTimer(self)
        # Channelizer-style staged progress, tuned to ~3 seconds total.
        self._timer.setInterval(38)
        self._timer.timeout.connect(self._tick)

    def start(self):
        self._value = 0
        self.progress.setValue(0)
        self.status.setText("Initializing UI…")
        self._timer.start()

    def _tick(self):
        # Slightly uneven fake progress feels more natural than a linear bar.
        if self._value < 28:
            step = 2
        elif self._value < 70:
            step = 1
        else:
            step = 2

        self._value = min(100, self._value + step)
        self.progress.setValue(self._value)

        if self._value < 28:
            self.status.setText("Initializing UI…")
        elif self._value < 55:
            self.status.setText("Loading PDF engine…")
        elif self._value < 82:
            self.status.setText("Preparing tools & workers…")
        elif self._value < 100:
            self.status.setText("Applying theme & settings…")
        else:
            self.status.setText("Ready")
            self._timer.stop()
            QTimer.singleShot(260, self._finish)

    def _finish(self):
        if getattr(self, "_finished", False):
            return
        self._finished = True
        self._timer.stop()
        self.hide()
        self.finishedStartup.emit()

    def mousePressEvent(self, event):
        # Click anywhere to skip the intro.
        self._finish()

    def showEvent(self, event):
        super().showEvent(event)
        try:
            screen = QApplication.primaryScreen()
            if screen is not None:
                rect = self.frameGeometry()
                rect.moveCenter(screen.availableGeometry().center())
                self.move(rect.topLeft())
        except Exception:
            pass

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.settings=load_settings(); self.category="All tools"; self.query=""
        self.setWindowIcon(load_app_icon())
        self.setWindowTitle(f"{APP_NAME} {APP_VERSION} — Heavy-Duty PDF Toolbox"); self.resize(1320,850); self.setMinimumSize(1000,680); self.setAcceptDrops(True)
        central=QWidget(); self.setCentralWidget(central); root=QHBoxLayout(central); root.setContentsMargins(0,0,0,0); root.setSpacing(0)
        self.sidebar=self._build_sidebar(); root.addWidget(self.sidebar)
        self.right=QWidget(); rr=QVBoxLayout(self.right); rr.setContentsMargins(0,0,0,0); rr.setSpacing(0); rr.addWidget(self._build_topbar())
        self.scroll=QScrollArea(); self.scroll.setWidgetResizable(True); self.scroll.setFrameShape(QFrame.Shape.NoFrame); rr.addWidget(self.scroll,1); root.addWidget(self.right,1)
        self._cols=0
        self.apply_theme(self.settings.get("theme","dark")); self.refresh_tools()
        QShortcut(QKeySequence("Ctrl+F"), self, activated=lambda: (self.search.setFocus(), self.search.selectAll()))
        QShortcut(QKeySequence("Escape"), self.search, activated=self.search.clear, context=Qt.ShortcutContext.WidgetShortcut)

    def _build_sidebar(self):
        f=QFrame(); f.setObjectName("sidebar"); f.setFixedWidth(230); lay=QVBoxLayout(f); lay.setContentsMargins(16,20,16,16); lay.setSpacing(5)
        brand=QLabel("PDF<span style='color:#ff4d5f'>Mule</span>"); brand.setTextFormat(Qt.TextFormat.RichText); brand.setObjectName("brand"); lay.addWidget(brand)
        v=QLabel(f"Heavy-Duty PDF Toolbox  •  v{APP_VERSION}"); v.setStyleSheet("color:#7f8998;font-size:8.5pt;margin-bottom:10px;"); lay.addWidget(v)
        self.navs={}
        icons={"All tools":"▦","Organize PDF":"🧩","Optimize PDF":"🗜","Convert to PDF":"⬇","Convert from PDF":"⬆","Edit PDF":"✎","PDF Security":"🔒","PDF Intelligence":"✨","Workflows":"⚙"}
        for c in CATEGORIES:
            b=QPushButton(f"{icons.get(c,'•')}   {c}"); b.setObjectName("nav"); b.setProperty("active",c==self.category); b.clicked.connect(lambda _=False,x=c:self.set_category(x)); lay.addWidget(b); self.navs[c]=b
        lay.addStretch()
        dep=QPushButton("🧰   Dependency status"); dep.setObjectName("nav"); dep.clicked.connect(self.show_dependencies); lay.addWidget(dep)
        theme=QPushButton("☼ / ☾   Toggle theme"); theme.setObjectName("nav"); theme.clicked.connect(self.toggle_theme); lay.addWidget(theme)
        about=QPushButton("ⓘ   About"); about.setObjectName("nav"); about.clicked.connect(self.about); lay.addWidget(about)
        return f

    def _build_topbar(self):
        f=QFrame(); f.setObjectName("topbar"); f.setFixedHeight(74); lay=QHBoxLayout(f); lay.setContentsMargins(24,14,24,14)
        self.search=QLineEdit(); self.search.setPlaceholderText("Search tools…  (Ctrl+F)  e.g. merge, OCR, protect, Word, συμπίεση"); self.search.setClearButtonEnabled(True); self.search.textChanged.connect(self.on_search); self.search.returnPressed.connect(self.open_first_match); lay.addWidget(self.search,1)
        recent=QPushButton("Recent outputs"); recent.clicked.connect(self.show_recent); lay.addWidget(recent)
        return f

    def set_category(self,c):
        self.category=c
        for k,b in self.navs.items(): b.setProperty("active",k==c); b.style().unpolish(b); b.style().polish(b)
        self.refresh_tools()
    def on_search(self,t): self.query=t.strip().lower(); self.refresh_tools()

    def _columns_for_width(self):
        w=self.scroll.viewport().width() if self.scroll.viewport().width()>50 else self.width()-260
        return max(1,min(5,(w-56+15)//(210+15)))

    def visible_tools(self):
        q=self.query
        out=[]
        for t in TOOLS:
            if self.category!="All tools" and t.category!=self.category: continue
            hay=(t.title+" "+t.desc+" "+t.category+" "+t.key+" "+TOOL_KEYWORDS.get(t.key,"")).lower()
            if q and not all(word in hay for word in q.split()): continue
            out.append(t)
        return out

    def refresh_tools(self):
        page=QWidget(); outer=QVBoxLayout(page); outer.setContentsMargins(28,25,28,30); outer.setSpacing(18)
        title=QLabel(self.category if self.category!="All tools" else "Every PDF tool in one desktop app"); title.setObjectName("heroTitle"); outer.addWidget(title)
        sub=QLabel("Private-first local processing. Drag PDFs, images, Office files or folders onto the window, or open any tool card below."); sub.setObjectName("heroSub"); sub.setWordWrap(True); outer.addWidget(sub)
        gridw=QWidget(); grid=QGridLayout(gridw); grid.setContentsMargins(0,8,0,0); grid.setHorizontalSpacing(15); grid.setVerticalSpacing(15)
        visible=self.visible_tools()
        cols=self._columns_for_width(); self._cols=cols
        for i,t in enumerate(visible):
            c=ToolCard(t); c.clicked.connect(self.open_tool); grid.addWidget(c,i//cols,i%cols)
        for col in range(cols): grid.setColumnStretch(col,1)
        if not visible:
            empty=QLabel(f"No tools match “{self.query}”. Try another word, e.g. merge, OCR, compress, Word, sign.")
            empty.setObjectName("heroSub"); empty.setWordWrap(True); outer.addWidget(empty)
        elif self.query:
            hint=QLabel(f"{len(visible)} tool(s) found — press Enter to open the first one."); hint.setObjectName("heroSub"); outer.addWidget(hint)
        outer.addWidget(gridw); outer.addStretch(); self.scroll.setWidget(page)

    def open_first_match(self):
        v=self.visible_tools()
        if v: self.open_tool(v[0])

    def resizeEvent(self,e):
        super().resizeEvent(e)
        # Re-check after Qt has laid out the scroll viewport for the new size.
        QTimer.singleShot(0, self._maybe_reflow)

    def _maybe_reflow(self):
        if self._cols and self._columns_for_width()!=self._cols:
            self.refresh_tools()

    def showEvent(self,e):
        super().showEvent(e)
        QTimer.singleShot(0, self._maybe_reflow)

    def open_tool(self,tool,files=None):
        files=list(files or [])
        if tool.key == "organize":
            dlg = OrganizeDialog(self.settings, self)
            if files: QTimer.singleShot(0, lambda: dlg.load_pdf(files[0]))
        elif tool.key == "merge":
            dlg = VisualMergeDialog(self.settings, self, initial_files=files or None)
        elif tool.key == "split":
            dlg = VisualSplitDialog(self.settings, self, initial_files=files or None)
        else:
            dlg = ToolDialog(tool,self.settings,self)
            if files: dlg._add_paths(files)
        dlg.exec()
        dlg.deleteLater()

    def apply_theme(self,theme):
        self.settings["theme"]=theme; save_settings(self.settings); QApplication.instance().setStyleSheet(themed_qss(theme))
    def toggle_theme(self): self.apply_theme("light" if self.settings.get("theme")=="dark" else "dark")

    def show_recent(self):
        rec=[r for r in self.settings.get("recent",[]) if r.get("path") and Path(r["path"]).exists()]
        if len(rec)!=len(self.settings.get("recent",[])): self.settings["recent"]=rec; save_settings(self.settings)
        if not rec: QMessageBox.information(self,"Recent outputs","No completed jobs yet (outputs that were moved or deleted are hidden)."); return
        dlg=QDialog(self); dlg.setWindowTitle("Recent outputs"); dlg.resize(760,420); lay=QVBoxLayout(dlg); lw=QListWidget(); lay.addWidget(lw)
        for r in rec:
            when=r.get("when","")
            item=QListWidgetItem(f"{r.get('tool','')}   ·   {Path(r['path']).name}" + (f"   ·   {when}" if when else "")); item.setToolTip(r["path"]); item.setData(Qt.ItemDataRole.UserRole,r.get("path")); lw.addItem(item)
        lw.setCurrentRow(0)
        cur=lambda: lw.currentItem().data(Qt.ItemDataRole.UserRole) if lw.currentItem() else None
        lw.itemDoubleClicked.connect(lambda it: open_folder(it.data(Qt.ItemDataRole.UserRole)))
        row=QHBoxLayout(); lab=QLabel("Double-click opens the folder."); lab.setObjectName("heroSub"); row.addWidget(lab); row.addStretch()
        of=QPushButton("Open file"); of.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(cur())) if cur() and Path(cur()).is_file() else None)
        op=QPushButton("Open folder"); op.clicked.connect(lambda: open_folder(cur()) if cur() else None)
        clr=QPushButton("Clear list"); clr.clicked.connect(lambda: (self.settings.__setitem__("recent",[]), save_settings(self.settings), dlg.accept()))
        row.addWidget(clr); row.addWidget(of); row.addWidget(op); lay.addLayout(row); dlg.exec()

    def show_dependencies(self):
        dlg = DependencyStatusDialog(self)
        dlg.exec()


    def about(self):
        dlg = QDialog(self)
        dlg.setWindowTitle(f"About {APP_NAME}")
        dlg.setWindowIcon(load_app_icon())
        dlg.setModal(True)
        dlg.resize(760, 660)
        dlg.setMinimumSize(700, 630)

        root = QVBoxLayout(dlg)
        root.setContentsMargins(24, 22, 24, 20)
        root.setSpacing(16)

        hero = QFrame()
        hero.setObjectName("previewPanel")
        hero_l = QHBoxLayout(hero)
        hero_l.setContentsMargins(18, 18, 18, 18)
        hero_l.setSpacing(18)

        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo_label.setMinimumSize(260, 150)
        logo_label.setStyleSheet("background:#ffffff;border-radius:14px;padding:10px;")
        pm = QPixmap(str(app_logo_path()))
        if not pm.isNull():
            logo_label.setPixmap(pm.scaled(300, 180, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            logo_label.setText("PDF Mule")
            logo_label.setStyleSheet("background:#ffffff;color:#20242c;border-radius:14px;font-size:22pt;font-weight:900;padding:10px;")
        hero_l.addWidget(logo_label, 0)

        info_box = QVBoxLayout()
        title = QLabel(APP_NAME)
        title.setStyleSheet("font-size:25pt;font-weight:900;")
        subtitle = QLabel(f"Heavy-Duty PDF Toolbox  •  Version {APP_VERSION}")
        subtitle.setStyleSheet("color:#8d98a8;font-size:10.5pt;")
        info_box.addWidget(title)
        info_box.addWidget(subtitle)
        info_box.addSpacing(6)

        desc = QLabel(
            "A modern local PDF toolbox for organizing, splitting, merging, converting, "
            "optimizing, editing and securing PDF files. Built for large batch jobs, smart "
            "worker profiles and visual page workflows."
        )
        desc.setWordWrap(True)
        desc.setStyleSheet("font-size:10.5pt;line-height:1.35;")
        info_box.addWidget(desc)

        badges = QLabel(
            "<b>Core:</b> PyQt6 + PyMuPDF + Pillow<br>"
            "<b>Optional:</b> Tesseract OCR, LibreOffice, Ghostscript, pdf2docx, "
            "python-pptx, openpyxl, pymupdf4llm and deep-translator"
        )
        badges.setWordWrap(True)
        badges.setTextFormat(Qt.TextFormat.RichText)
        badges.setStyleSheet("color:#8d98a8;")
        info_box.addWidget(badges)
        info_box.addStretch(1)
        hero_l.addLayout(info_box, 1)
        root.addWidget(hero)

        card = QFrame()
        card.setObjectName("previewPanel")
        card_lay = QVBoxLayout(card)
        card_lay.setContentsMargins(18, 16, 18, 16)
        card_lay.setSpacing(6)

        team = QLabel("ALU DEV TEAM @ 2026")
        team.setStyleSheet("font-size:12pt;font-weight:900;color:#ff4d5f;")
        card_lay.addWidget(team)

        developer = QLabel("Nikolaos K. Paridis シ  (nikkpap)")
        developer.setStyleSheet("font-size:11pt;font-weight:700;")
        card_lay.addWidget(developer)

        links = QLabel(
            "<a href='mailto:nikkpap@gmail.com'>nikkpap@gmail.com</a><br>"
            "GitHub: <a href='https://github.com/nikkpap'>github.com/nikkpap</a><br>"
            "LinkedIn: <a href='https://www.linkedin.com/in/nickparidis/'>"
            "linkedin.com/in/nickparidis</a>"
        )
        links.setTextFormat(Qt.TextFormat.RichText)
        links.setOpenExternalLinks(True)
        links.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        links.setStyleSheet("font-size:10.5pt;")
        card_lay.addWidget(links)

        motto = QLabel(
            "Open for discussions and ideas. Public brainstorming is better "
            "than no brainstorming :) Cheers!"
        )
        motto.setWordWrap(True)
        motto.setStyleSheet("margin-top:8px;font-style:italic;color:#8d98a8;")
        card_lay.addWidget(motto)
        root.addWidget(card)

        note = QLabel(
            "PDF Mule is an independent desktop utility. It is not affiliated "
            "with or endorsed by iLovePDF or its owners."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color:#7f8998;font-size:8.5pt;")
        root.addWidget(note)

        upd_row = QHBoxLayout()
        chk = QCheckBox("Check for updates when PDF Mule starts")
        chk.setChecked(bool(self.settings.get("update_check", True)) and ota_supported())
        chk.setEnabled(ota_supported())
        chk.toggled.connect(lambda v: (self.settings.__setitem__("update_check", bool(v)), save_settings(self.settings)))
        upd_row.addWidget(chk); upd_row.addStretch()
        gh = QPushButton("GitHub page"); gh.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(APP_PAGE)))
        cu = QPushButton("Check for updates"); cu.setObjectName("primary")
        cu.clicked.connect(lambda: run_update_check(dlg, quiet=False))
        upd_row.addWidget(gh); upd_row.addWidget(cu)
        root.addLayout(upd_row)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(dlg.reject)
        root.addWidget(buttons)

        dlg.exec()

    def dragEnterEvent(self,e:QDragEnterEvent):
        if e.mimeData().hasUrls(): e.acceptProposedAction()

    def dropEvent(self,e:QDropEvent):
        """Smart drop: suggests the tools that fit what was dropped (PDFs, images,
        Office files, folders) instead of always opening one fixed tool."""
        paths=[u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        e.acceptProposedAction()
        self.handle_paths(paths)

    def handle_paths(self,paths):
        """Shared by drag & drop and command-line arguments ("Open with PDF Mule")."""
        files=[]
        for p in paths:
            if Path(p).is_dir(): files.extend(collect_folder_files(p,"*.pdf *.jpg *.jpeg *.png *.bmp *.tif *.tiff *.doc *.docx *.ppt *.pptx *.xls *.xlsx",True))
            elif Path(p).is_file(): files.append(p)
        if not files: return
        exts={Path(f).suffix.lower() for f in files}
        fitting=[t for t in TOOLS if t.key!="html2pdf" and exts <= filter_suffixes(t.extensions) and (t.multi or len(files)==1)]
        if not fitting:
            QMessageBox.information(self,APP_NAME,"No tool accepts this combination of files. Drop PDFs, images or Office files of one kind.")
            return
        preferred=["merge","split","compress","ocr","organize","img2pdf","word2pdf","ppt2pdf","xls2pdf","pdf2word","protect","rotate"]
        fitting.sort(key=lambda t: preferred.index(t.key) if t.key in preferred else 99)
        dlg=QDialog(self); dlg.setWindowTitle(f"{APP_NAME} — {len(files)} file(s) dropped"); dlg.resize(480,520)
        lay=QVBoxLayout(dlg); lab=QLabel(f"<b>{len(files)} file(s)</b> dropped. Choose a tool:"); lay.addWidget(lab)
        lw=QListWidget()
        for t in fitting:
            it=QListWidgetItem(f"{t.icon}   {t.title}"); it.setToolTip(t.desc); it.setData(Qt.ItemDataRole.UserRole,t.key); lw.addItem(it)
        lw.setCurrentRow(0); lay.addWidget(lw,1)
        bb=QDialogButtonBox(QDialogButtonBox.StandardButton.Open|QDialogButtonBox.StandardButton.Cancel); lay.addWidget(bb)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject); lw.itemDoubleClicked.connect(lambda *_: dlg.accept())
        if dlg.exec()!=QDialog.DialogCode.Accepted or lw.currentItem() is None: return
        key=lw.currentItem().data(Qt.ItemDataRole.UserRole)
        self.open_tool(next(t for t in TOOLS if t.key==key), files)


def main():
    multiprocessing.freeze_support()
    app=QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("NikkPap Tools")
    app.setWindowIcon(load_app_icon())
    app.setStyle("Fusion")

    # One running copy: a newer version takes over from an older one (OTA), a second
    # start of the same version just brings the open window to the front.
    if not single_instance():
        sys.exit(0)

    # Channelizer-style fake startup presentation: logo, staged status,
    # version/build metadata and ~3-second progress before the main UI appears.
    splash = StartupSplash()
    splash.show()
    app.processEvents()

    w = MainWindow()
    w.hide()

    splash.finishedStartup.connect(w.show)
    splash.finishedStartup.connect(splash.deleteLater)
    attach_instance_server(w)
    splash.finishedStartup.connect(lambda: QTimer.singleShot(3000, lambda: auto_install_addons(w)))
    if ota_supported() and w.settings.get("update_check", True):
        # Silent check a few seconds after the window opens; asks only when there is something newer.
        splash.finishedStartup.connect(lambda: QTimer.singleShot(5000, lambda: run_update_check(w, quiet=True)))
    cli_paths = [a for a in sys.argv[1:] if Path(a).exists()]
    if cli_paths:
        splash.finishedStartup.connect(lambda: QTimer.singleShot(150, lambda: w.handle_paths(cli_paths)))
    splash.start()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()
