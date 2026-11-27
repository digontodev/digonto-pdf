#!/usr/bin/env python3
"""
Digonto PDF - a PDF viewer with tabs, and tools to create, shrink, organize, sign and edit PDF files.

Runs on Ubuntu/Linux, macOS and Windows 11.
Uses only two libraries (see requirements.txt, pinned with SHA-256 hashes):
    PySide6-Essentials  (the window, by The Qt Company)
    PyMuPDF             (the PDF engine, by Artifex)
Word/Excel/text/image conversion is Digonto's own code (digonto_convert.py).
The app never connects to the internet.
Optional: Ghostscript, if you install it, is also tried when shrinking.

Original files are never changed unless you choose to save over them.
"""
import glob
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time

try:
    import pymupdf as fitz  # PyMuPDF >= 1.24
except ImportError:
    import fitz  # older PyMuPDF

import digonto_convert as conv
import digonto_ui as ui
import pdf_ops as ops
import pdf_panel as panel
import pdf_safe
from PySide6.QtCore import (QByteArray, QBuffer, QIODevice, QItemSelectionModel, QPointF, QRect, QTimer, QUrl,
                            QRectF, QSettings, QSize, QStandardPaths, Qt,
                            QThread, Signal)
from PySide6.QtGui import (QAction, QColor, QFont, QIcon, QImage, QKeySequence,
                           QPainter, QPainterPath, QPen, QPixmap, QShortcut, QDesktopServices)
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QButtonGroup, QCheckBox, QColorDialog,
    QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout,
    QGraphicsItem, QGraphicsPixmapItem, QGraphicsRectItem, QGraphicsScene,
    QFrame, QGraphicsView, QHBoxLayout, QHeaderView, QInputDialog, QLabel,
    QLineEdit, QListView, QProgressBar, QListWidget, QListWidgetItem, QMainWindow,
    QMessageBox, QPlainTextEdit, QPushButton, QRadioButton, QSlider,
    QSpinBox, QSplitter, QStackedWidget, QMenu, QStyle, QStyledItemDelegate, QTableWidget,
    QTableWidgetItem, QTabBar, QTabWidget, QToolButton, QVBoxLayout, QWidget)

APP_NAME = "Digonto PDF"
APP_VERSION = "4.0.0"
HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
PDF_FILTER = "PDF files (*.pdf *.PDF)"
IMG_FILTER = "Images (*.png *.jpg *.jpeg *.PNG *.JPG *.JPEG)"


# --------------------------------------------------------------------------
# small helpers
# --------------------------------------------------------------------------
def human(n):
    n = float(n)
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024


def settings():
    return QSettings("DigontoPDF", "DigontoPDF")


def last_dir():
    d = settings().value("last_dir", "")
    if d and os.path.isdir(d):
        return d
    docs = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
    return docs or os.path.expanduser("~")


def remember_dir(path):
    if path:
        settings().setValue("last_dir", path if os.path.isdir(path) else os.path.dirname(path))


def unique_path(path):
    """file.pdf -> file.pdf, or file_2.pdf, file_3.pdf ... if it exists."""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    i = 2
    while os.path.exists(f"{base}_{i}{ext}"):
        i += 1
    return f"{base}_{i}{ext}"


def default_name(src, suffix):
    base = os.path.splitext(os.path.basename(src))[0] if src else "document"
    folder = os.path.dirname(src) if src else last_dir()
    return os.path.join(folder, f"{base}_{suffix}.pdf")


def ensure_pdf_ext(path):
    return path if path.lower().endswith(".pdf") else path + ".pdf"


def open_pdf(parent, path):
    """Open a PDF from bytes (no file lock on Windows). Asks for a password
    when needed. Returns a fitz.Document or None."""
    try:
        with open(path, "rb") as f:
            data = f.read()
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception as e:
        QMessageBox.critical(parent, APP_NAME, f"Could not open this file:\n{path}\n\n{e}")
        return None
    if doc.needs_pass:
        while True:
            pw, ok = QInputDialog.getText(parent, APP_NAME, f"{os.path.basename(path)} is protected.\nPassword:",
                                          QLineEdit.Password)
            if not ok:
                return None
            if doc.authenticate(pw):
                doc._dg_pw = pw            # kept, so saving does not silently drop the protection
                break
    if doc.page_count == 0:
        QMessageBox.warning(parent, APP_NAME, "This PDF has no pages.")
        return None
    return doc


def office_mode():
    """Settings: every PDF Digonto saves is made upload-ready (on unless switched off)"""
    return settings().value("office_ready", True, type=bool)


def office_bytes(data):
    """bytes of a PDF about to be written: made upload-ready (the same steps as Save for upload) when that setting is on"""
    if office_mode() and isinstance(data, (bytes, bytearray)) and data[:5] == b"%PDF-":
        return pdf_safe.office_ready(bytes(data))
    return data


def save_doc(parent, doc, path, **kw):
    """Save to a temp file first, then move into place (safe even when
    overwriting the file that was opened)."""
    path = ensure_pdf_ext(path)
    folder = os.path.dirname(os.path.abspath(path))
    fd, tmp = tempfile.mkstemp(suffix=".pdf", dir=folder)
    os.close(fd)
    try:
        if doc.is_encrypted or kw.get("encryption"):
            doc.save(tmp, garbage=3, deflate=True, **kw)
        elif office_mode():
            # exactly what Save for upload does (same input, same steps): fonts are trimmed only once
            data, _pics = pdf_safe.upload_copy(doc.tobytes())
            with open(tmp, "wb") as fh:
                fh.write(data)
        else:
            copy = fitz.open(stream=doc.tobytes(), filetype="pdf")    # a copy: the open document keeps its full fonts
            pdf_safe.finish(copy)                                      # only the letters used go into the file
            copy.save(tmp, garbage=3, deflate=True, **kw)
            copy.close()
        os.replace(tmp, path)
    except Exception as e:
        try:
            os.remove(tmp)
        except OSError:
            pass
        QMessageBox.critical(parent, APP_NAME, f"Could not save:\n{path}\n\n{e}")
        return False
    remember_dir(path)
    ui.add_recent("digonto-pdf", os.path.abspath(path))
    return True


def render_page(page, zoom, dpr=1.0, matrix_extra=None):
    """Render a page to a QPixmap. Scene units stay in PDF points:
    the pixmap is rendered at zoom*dpr and carries that pixel ratio."""
    scale = zoom * dpr
    m = fitz.Matrix(scale, scale)
    if matrix_extra is not None:
        m = m * matrix_extra
    pix = page.get_pixmap(matrix=m, alpha=False)
    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
    pm = QPixmap.fromImage(img)
    pm.setDevicePixelRatio(dpr)
    return pm


def qimage_png_bytes(img):
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "PNG")
    buf.close()
    return bytes(ba)


def parse_ranges(text, count):
    """'1-3, 5, 8-' -> [0,1,2,4,7,...]. Empty or 'all' -> every page."""
    text = (text or "").strip().lower()
    if text in ("", "all", "*"):
        return list(range(count))
    out = []
    for part in re.split(r"[,\s;]+", text):
        if not part:
            continue
        m = re.fullmatch(r"(\d*)-(\d*)", part)
        if m:
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else count
        elif part.isdigit():
            a = b = int(part)
        else:
            raise ValueError(f"Not understood: {part}")
        if a < 1 or b > count or a > b:
            raise ValueError(f"Pages must be between 1 and {count}: {part}")
        out.extend(range(a - 1, b))
    return out


def make_app_icon():
    icon = QIcon()
    for size in (16, 32, 48, 64, 128, 256, 512):
        f = os.path.join(ASSETS, f"digonto-pdf-{size}.png")
        if os.path.exists(f):
            icon.addFile(f, QSize(size, size))
    if not icon.isNull():
        return icon
    pm = QPixmap(64, 64)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QColor("#FFFFFF"))
    p.setPen(QPen(QColor("#555555"), 2))
    p.drawRoundedRect(10, 4, 44, 56, 4, 4)
    p.setBrush(QColor("#D93025"))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(4, 30, 46, 20, 3, 3)
    p.setPen(QColor("white"))
    f = QFont()
    f.setBold(True)
    f.setPixelSize(14)
    p.setFont(f)
    p.drawText(QRect(4, 30, 46, 20), Qt.AlignCenter, "PDF")
    p.end()
    return QIcon(pm)


def hline():
    w = QWidget()
    w.setFixedHeight(1)
    w.setStyleSheet("background:#CCCCCC")
    return w


# --------------------------------------------------------------------------
# Tab: Create PDF (images, svg, Word, Excel, text, pdf)
# --------------------------------------------------------------------------
KIND_STYLE = {   # label, badge text, colour
    "image": ("Picture", "IMG", "#C2395C"),
    "svg": ("SVG drawing", "SVG", "#C2395C"),
    "word": ("Word document", "DOC", "#2B5797"),
    "excel": ("Excel workbook", "XLS", "#C2395C"),
    "text": ("Text file", "TXT", "#6D6D6D"),
    "pdf": ("PDF", "PDF", "#D93025"),
}
CREATE_FILTER = ("All supported (*.png *.jpg *.jpeg *.jfif *.bmp *.gif *.tif *.tiff *.webp *.svg *.docx *.docm "
                 "*.xlsx *.xlsm *.txt *.pdf *.PNG *.JPG *.JPEG *.SVG *.DOCX *.XLSX *.PDF);;"
                 "Pictures (*.png *.jpg *.jpeg *.jfif *.bmp *.gif *.tif *.tiff *.webp *.svg);;"
                 "Word (*.docx *.docm);;Excel (*.xlsx *.xlsm);;Text (*.txt);;PDF (*.pdf)")


def badge_icon(kind):
    _, text, color = KIND_STYLE.get(kind, ("", "?", "#888888"))
    pm = QPixmap(40, 40)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QColor(color))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(2, 2, 36, 36, 7, 7)
    p.setPen(QColor("white"))
    f = QFont()
    f.setBold(True)
    f.setPixelSize(12)
    p.setFont(f)
    p.drawText(QRect(2, 2, 36, 36), Qt.AlignCenter, text)
    p.end()
    return QIcon(pm)




# --------------------------------------------------------------------------
# Create PDF (pictures, SVG, Word, Excel, text, PDF) - the result opens in new tabs
# --------------------------------------------------------------------------
class CreateWorker(QThread):
    progress = Signal(int, str)            # index, text
    done = Signal(list, list)              # [(name, pdf bytes)], [errors]

    def __init__(self, files, opts):
        super().__init__()
        self.files, self.opts = files, opts

    def run(self):
        o = self.opts
        out, errors = [], []
        combined = fitz.open() if o["combine"] else None
        bases = [os.path.splitext(os.path.basename(f))[0].lower() for f in self.files]
        for i, f in enumerate(self.files):
            self.progress.emit(i, os.path.basename(f))
            try:
                doc = conv.convert_file(f, o["page"], o["margin"], o["office"], o["grid"])
            except Exception as e:
                errors.append(f"{os.path.basename(f)}: {e}")
                continue
            if combined is not None:
                combined.insert_pdf(doc)
            else:
                base, ext = os.path.splitext(os.path.basename(f))
                if bases.count(base.lower()) > 1:      # photo.jpg + photo.docx -> photo_jpg, photo_docx
                    base += "_" + ext[1:].lower()
                try:
                    out.append((base + ".pdf", doc.tobytes(garbage=3, deflate=True)))
                except Exception as e:
                    errors.append(f"{os.path.basename(f)}: {e}")
        if combined is not None and combined.page_count:
            name = os.path.splitext(os.path.basename(self.files[0]))[0] if len(self.files) == 1 else "combined"
            try:
                out.append((name + ".pdf", combined.tobytes(garbage=3, deflate=True)))
            except Exception as e:
                errors.append(f"Combining: {e}")
        self.progress.emit(len(self.files), "")
        self.done.emit(out, errors)


class CreateDialog(QDialog):
    def __init__(self, parent=None, files=()):
        super().__init__(parent)
        self.setWindowTitle("Create PDF")
        self.resize(720, 560)
        self.worker = None
        self.results = []
        self.setAcceptDrops(True)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("Add pictures (PNG, JPG, SVG ...), Word (.docx), Excel (.xlsx), text or PDF files. "
                             "Drag files here, or drag in the list to change the order."))
        self.list = QListWidget()
        self.list.setIconSize(QSize(32, 32))
        self.list.setSpacing(2)
        self.list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list.setDragDropMode(QAbstractItemView.InternalMove)
        self.list.setDefaultDropAction(Qt.MoveAction)
        lay.addWidget(self.list, 1)
        row = QHBoxLayout()
        for text, slot in (("Add files...", self.add_dialog), ("Remove selected", self.remove_selected),
                           ("Clear list", self.list.clear), ("Move up", lambda: self.move(-1)),
                           ("Move down", lambda: self.move(1))):
            b = QPushButton(text)
            b.clicked.connect(slot)
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        lay.addWidget(hline())
        form = QFormLayout()
        self.page = QComboBox()
        for label, key in (("A4", "A4"), ("Letter (US)", "Letter"), ("A5", "A5"), ("A3", "A3"),
                           ("Legal (US)", "Legal"), ("Same size as the picture", "fit")):
            self.page.addItem(label, key)
        form.addRow("Page size for pictures:", self.page)
        self.margin = QSpinBox()
        self.margin.setRange(0, 50)
        self.margin.setValue(7)
        self.margin.setSuffix(" mm")
        form.addRow("Margin around pictures:", self.margin)
        lay.addLayout(form)
        self.grid = QCheckBox("Excel: show grid lines")
        self.grid.setChecked(True)
        lay.addWidget(self.grid)
        self.office = QCheckBox("Word/Excel: exact copy with LibreOffice / MS Office")
        if conv.office_available():
            self.office.setToolTip("Uses the office program installed on this computer for a 1:1 layout. "
                                   "Unticked = Digonto's own converter.")
        else:
            self.office.setEnabled(False)
            self.office.setToolTip("Not installed - Digonto's own converter is used (no extra software needed).")
        lay.addWidget(self.office)
        self.combine = QRadioButton("All files together in ONE PDF (in the order of the list)")
        self.separate = QRadioButton("One PDF for each file (each opens in its own tab)")
        self.combine.setChecked(True)
        lay.addWidget(self.combine)
        lay.addWidget(self.separate)
        row = QHBoxLayout()
        self.bar = QProgressBar()
        self.bar.setVisible(False)
        self.status = QLabel("")
        row.addWidget(self.bar, 1)
        row.addWidget(self.status)
        row.addStretch()
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        self.go = QPushButton("Create PDF")
        ui.primary(self.go)
        self.go.clicked.connect(self.start)
        row.addWidget(cancel)
        row.addWidget(self.go)
        lay.addLayout(row)
        if files:
            self.add_files(files)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        self.add_files([u.toLocalFile() for u in e.mimeData().urls()])

    def add_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Choose files", last_dir(), CREATE_FILTER)
        if files:
            remember_dir(files[0])
            self.add_files(files)

    def add_files(self, files):
        skipped = []
        for f in files:
            if os.path.isdir(f):
                self.add_files(sorted(os.path.join(f, x) for x in os.listdir(f)))
                continue
            kind = conv.kind_of(f)
            if not kind or not os.path.isfile(f):
                skipped.append(os.path.basename(f))
                continue
            it = QListWidgetItem(badge_icon(kind), f"{os.path.basename(f)}      ({KIND_STYLE[kind][0]})")
            it.setData(Qt.UserRole, f)
            it.setToolTip(f)
            self.list.addItem(it)
        if skipped:
            old = [s for s in skipped if s.lower().endswith((".doc", ".xls"))]
            msg = "These files are not supported and were skipped:\n" + "\n".join(skipped[:15])
            if old:
                msg += ("\n\nOld .doc/.xls files: open them in Word, Excel or LibreOffice and "
                        "'Save as' .docx/.xlsx first.")
            QMessageBox.information(self, APP_NAME, msg)

    def remove_selected(self):
        for it in self.list.selectedItems():
            self.list.takeItem(self.list.row(it))

    def move(self, step):
        rows = sorted(self.list.row(i) for i in self.list.selectedItems())
        if not rows or (step < 0 and rows[0] == 0) or (step > 0 and rows[-1] == self.list.count() - 1):
            return
        for r in (rows if step < 0 else reversed(rows)):
            it = self.list.takeItem(r)
            self.list.insertItem(r + step, it)
        self.list.clearSelection()
        for r in rows:
            self.list.item(r + step).setSelected(True)

    def start(self):
        if self.worker and self.worker.isRunning():
            return
        if self.list.count() == 0:
            self.add_dialog()
            if self.list.count() == 0:
                return
        files = [self.list.item(i).data(Qt.UserRole) for i in range(self.list.count())]
        opts = dict(page=self.page.currentData(), margin=self.margin.value() * 72 / 25.4,
                    office=self.office.isChecked() and self.office.isEnabled(), grid=self.grid.isChecked(),
                    combine=self.combine.isChecked())
        self.go.setEnabled(False)
        self.bar.setRange(0, len(files))
        self.bar.setValue(0)
        self.bar.setVisible(True)
        self.worker = CreateWorker(files, opts)
        self.worker.progress.connect(self.on_progress)
        self.worker.done.connect(self.on_done)
        self.worker.start()

    def on_progress(self, i, name):
        self.bar.setValue(i)
        self.status.setText(f"Converting {name}..." if name else "")

    def on_done(self, out, errors):
        self.go.setEnabled(True)
        self.bar.setVisible(False)
        self.status.setText("")
        if errors:
            QMessageBox.warning(self, APP_NAME, "Problems:\n" + "\n".join(errors[:10]))
        if out:
            self.results = out
            self.accept()

    def reject(self):
        if self.worker and self.worker.isRunning():
            self.worker.wait(3000)
        super().reject()


PRESETS = [
    # label, ghostscript setting, image dpi, jpeg quality
    ("Smallest - reading on screen, WhatsApp, email", "screen", 72, 45),
    ("Small - good quality, most documents (recommended)", "ebook", 150, 70),
    ("Larger - for printing", "printer", 300, 85),
]


def find_ghostscript():
    for name in ("gs", "gswin64c", "gswin32c"):
        p = shutil.which(name)
        if p:
            return p
    cands = []
    if sys.platform == "win32":
        for root in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            cands += sorted(glob.glob(os.path.join(root, "gs", "*", "bin", "gswin64c.exe")), reverse=True)
            cands += sorted(glob.glob(os.path.join(root, "gs", "*", "bin", "gswin32c.exe")), reverse=True)
    else:
        cands += ["/opt/homebrew/bin/gs", "/usr/local/bin/gs", "/usr/bin/gs"]
    for c in cands:
        if os.path.isfile(c):
            return c
    return None


def shrink_with_pymupdf(src, dst, dpi, quality):
    doc = fitz.open(src)
    if doc.needs_pass:
        raise RuntimeError("password protected")
    if hasattr(doc, "rewrite_images"):
        doc.rewrite_images(dpi_threshold=int(dpi * 1.25), dpi_target=dpi, quality=quality,
                           lossy=True, lossless=True, bitonal=False, color=True, gray=True)
    try:
        doc.subset_fonts()
    except Exception:
        pass
    doc.save(dst, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True,
             clean=True, use_objstms=1)
    doc.close()


def shrink_with_gs(gs, src, dst, setting):
    kw = {}
    if sys.platform == "win32":
        kw["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
    r = subprocess.run([gs, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.5",
                        f"-dPDFSETTINGS=/{setting}", "-dNOPAUSE", "-dBATCH", "-dQUIET",
                        f"-sOutputFile={dst}", src],
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, **kw)
    if r.returncode != 0:
        raise RuntimeError(r.stderr.decode(errors="ignore")[-300:] or "ghostscript failed")




class PageView(QGraphicsView):
    """Shows one page. Scene units = PDF points of the displayed page."""
    clicked = Signal(QPointF)
    doubleClicked = Signal(QPointF)
    rectSelected = Signal(QRectF)
    hovered = Signal(QPointF)
    flip = Signal(int)          # wheel at the end of a page: +1 next page, -1 previous page

    def __init__(self):
        super().__init__()
        self.setScene(QGraphicsScene(self))
        self.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        self.setBackgroundBrush(QColor("#7B7390"))
        self.setAlignment(Qt.AlignCenter)
        self.setMouseTracking(True)
        self.page_item = None
        self.zoom = 1.25
        self.mode = None          # None | "click" | "rect"
        self._start = None
        self._band = None
        self._flip_t = 0.0

    def show_page(self, page):
        if self.page_item is not None:
            self.scene().removeItem(self.page_item)
        pm = render_page(page, self.zoom, self.devicePixelRatioF() or 1.0)
        self.page_item = QGraphicsPixmapItem(pm)
        # pixmap logical size = points*zoom; scale back to points
        self.page_item.setScale(1.0 / self.zoom)
        self.page_item.setZValue(-10)
        self.scene().addItem(self.page_item)
        r = page.rect
        self.scene().setSceneRect(QRectF(0, 0, r.width, r.height))
        self.setTransform(self.transform().fromScale(self.zoom, self.zoom))

    def set_zoom(self, z, page):
        self.zoom = max(0.3, min(5.0, z))
        if page is not None:
            self.show_page(page)

    def fit_width(self, page):
        if page is None:
            return
        avail = self.viewport().width() - 30
        self.set_zoom(avail / page.rect.width, page)

    def fit_page(self, page):
        if page is None:
            return
        aw = self.viewport().width() - 30
        ah = self.viewport().height() - 30
        self.set_zoom(min(aw / page.rect.width, ah / page.rect.height), page)

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier and hasattr(self, "zoom_request"):
            self.zoom_request(1.15 if e.angleDelta().y() > 0 else 1 / 1.15)
            e.accept()
        else:
            sb = self.verticalScrollBar()
            d = e.angleDelta().y()
            at_edge = sb.maximum() == sb.minimum() or (d < 0 and sb.value() >= sb.maximum()) or (d > 0 and sb.value() <= sb.minimum())
            if at_edge and d != 0:
                now = time.time()
                if now - self._flip_t > 0.35:
                    self._flip_t = now
                    self.flip.emit(1 if d < 0 else -1)
                e.accept()
                return
            super().wheelEvent(e)

    def mousePressEvent(self, e):
        if self.mode and e.button() == Qt.LeftButton:
            p = self.mapToScene(e.position().toPoint())
            if self.mode == "click":
                self.clicked.emit(p)
                return
            if self.mode == "rect":
                self._start = p
                self._band = QGraphicsRectItem(QRectF(p, p))
                self._band.setPen(QPen(QColor("#D93025"), 0, Qt.DashLine))
                self._band.setBrush(QColor(217, 48, 37, 40))
                self.scene().addItem(self._band)
                return
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e):
        if self.mode == "click" and e.button() == Qt.LeftButton:
            self.doubleClicked.emit(self.mapToScene(e.position().toPoint()))
            return
        super().mouseDoubleClickEvent(e)

    def mouseMoveEvent(self, e):
        p = self.mapToScene(e.position().toPoint())
        if self._band is not None:
            self._band.setRect(QRectF(self._start, p).normalized())
            return
        self.hovered.emit(p)
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        if self._band is not None:
            r = self._band.rect()
            self.scene().removeItem(self._band)
            self._band = None
            if r.width() > 2 and r.height() > 2:
                self.rectSelected.emit(r)
            return
        super().mouseReleaseEvent(e)




def signatures_dir():
    base = QStandardPaths.writableLocation(QStandardPaths.AppDataLocation) or os.path.expanduser("~/.digonto-pdf")
    d = os.path.join(base, "signatures")
    os.makedirs(d, exist_ok=True)
    return d


class SignaturePad(QWidget):
    def __init__(self):
        super().__init__()
        self.setMinimumSize(640, 240)
        self.strokes = []          # list of (color, width, [QPointF])
        self.color = QColor("#000000")
        self.pen_width = 3
        self.setCursor(Qt.CrossCursor)

    def clear(self):
        self.strokes = []
        self.update()

    def undo(self):
        if self.strokes:
            self.strokes.pop()
            self.update()

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.strokes.append((QColor(self.color), self.pen_width, [e.position()]))
            self.update()

    def mouseMoveEvent(self, e):
        if e.buttons() & Qt.LeftButton and self.strokes:
            self.strokes[-1][2].append(e.position())
            self.update()

    @staticmethod
    def path_for(pts):
        path = QPainterPath(pts[0])
        if len(pts) == 1:
            path.lineTo(pts[0] + QPointF(0.1, 0.1))
            return path
        for i in range(1, len(pts) - 1):
            mid = (pts[i] + pts[i + 1]) / 2
            path.quadTo(pts[i], mid)
        path.lineTo(pts[-1])
        return path

    def paint_strokes(self, p, scale=1.0, offset=QPointF(0, 0)):
        p.setRenderHint(QPainter.Antialiasing)
        for color, width, pts in self.strokes:
            p.setPen(QPen(color, width * scale, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            p.setBrush(Qt.NoBrush)
            path = self.path_for([(q - offset) * scale for q in pts])
            p.drawPath(path)

    def paintEvent(self, e):
        p = QPainter(self)
        p.fillRect(self.rect(), Qt.white)
        p.setPen(QPen(QColor("#BBBBBB"), 1, Qt.DashLine))
        y = int(self.height() * 0.72)
        p.drawLine(30, y, self.width() - 30, y)
        p.setPen(QColor("#AAAAAA"))
        if not self.strokes:
            p.drawText(self.rect(), Qt.AlignCenter, "Sign here with the mouse, touchpad or pen")
        self.paint_strokes(p)
        p.end()

    def to_image(self, scale=4):
        """Transparent PNG, cropped to the ink."""
        if not self.strokes:
            return None
        bounds = QRectF()
        for _, w, pts in self.strokes:
            for q in pts:
                bounds = bounds.united(QRectF(q.x() - w, q.y() - w, 2 * w, 2 * w))
        bounds = bounds.adjusted(-4, -4, 4, 4)
        img = QImage(int(bounds.width() * scale) + 1, int(bounds.height() * scale) + 1, QImage.Format_ARGB32)
        img.fill(Qt.transparent)
        p = QPainter(img)
        self.paint_strokes(p, scale, bounds.topLeft())
        p.end()
        return img


class DrawSignatureDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Draw your signature")
        lay = QVBoxLayout(self)
        self.pad = SignaturePad()
        lay.addWidget(self.pad, 1)
        row = QHBoxLayout()
        row.addWidget(QLabel("Color:"))
        self.color = QComboBox()
        for name, c in (("Black", "#000000"), ("Blue", "#1A3FA8"), ("Dark blue", "#0B2161"), ("Red", "#B00020")):
            self.color.addItem(name, c)
        self.color.currentIndexChanged.connect(lambda: setattr(self.pad, "color", QColor(self.color.currentData())))
        row.addWidget(self.color)
        row.addWidget(QLabel("  Pen size:"))
        self.size = QSlider(Qt.Horizontal)
        self.size.setRange(1, 10)
        self.size.setValue(3)
        self.size.setFixedWidth(140)
        self.size.valueChanged.connect(lambda v: setattr(self.pad, "pen_width", v))
        row.addWidget(self.size)
        row.addStretch()
        u = QPushButton("Undo stroke")
        u.clicked.connect(self.pad.undo)
        c = QPushButton("Clear")
        c.clicked.connect(self.pad.clear)
        row.addWidget(u)
        row.addWidget(c)
        lay.addLayout(row)
        self.keep = QCheckBox("Keep this signature for next time")
        self.keep.setChecked(True)
        lay.addWidget(self.keep)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)


class SignatureItem(QGraphicsItem):
    HANDLE = 10

    def __init__(self, img, png, width_pt):
        super().__init__()
        self.img = img
        self.png = png
        self.aspect = img.height() / max(1, img.width())
        self.w = width_pt
        self.h = width_pt * self.aspect
        self.resizing = False
        self.setFlags(QGraphicsItem.ItemIsMovable | QGraphicsItem.ItemIsSelectable |
                      QGraphicsItem.ItemIsFocusable)
        self.setAcceptHoverEvents(True)
        self.setZValue(5)

    def boundingRect(self):
        return QRectF(-2, -2, self.w + 4, self.h + 4)

    def rect_pt(self):
        return QRectF(self.pos().x(), self.pos().y(), self.w, self.h)

    def paint(self, p, option, widget=None):
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        p.drawImage(QRectF(0, 0, self.w, self.h), self.img)
        if self.isSelected():
            p.setPen(QPen(QColor("#1A73E8"), 0, Qt.DashLine))
            p.setBrush(Qt.NoBrush)
            p.drawRect(QRectF(0, 0, self.w, self.h))
            p.setBrush(QColor("#1A73E8"))
            p.setPen(Qt.NoPen)
            h = self.HANDLE
            p.drawRect(QRectF(self.w - h / 2, self.h - h / 2, h, h))

    def near_handle(self, pos):
        return abs(pos.x() - self.w) < self.HANDLE and abs(pos.y() - self.h) < self.HANDLE

    def hoverMoveEvent(self, e):
        self.setCursor(Qt.SizeFDiagCursor if self.near_handle(e.pos()) else Qt.SizeAllCursor)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton and self.near_handle(e.pos()):
            self.resizing = True
            self.setSelected(True)
            e.accept()
            return
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if self.resizing:
            self.prepareGeometryChange()
            self.w = max(15.0, e.pos().x())
            self.h = self.w * self.aspect
            self.update()
            return
        super().mouseMoveEvent(e)

    def mouseReleaseEvent(self, e):
        self.resizing = False
        super().mouseReleaseEvent(e)


def place_image_on_page(page, rect_displayed, png):
    """rect_displayed is in the coordinates of the page as you see it
    (rotation applied). Converts to PDF coordinates and inserts the image
    upright."""
    pdf_safe.place_signature(page, (rect_displayed.x(), rect_displayed.y(), rect_displayed.width(), rect_displayed.height()), png)




# --------------------------------------------------------------------------
FONTS = [("Helvetica (Sans)", "helv", "hebo"), ("Times (Serif)", "tiro", "tibo"), ("Courier (Mono)", "cour", "cobo")]


def int_to_qcolor(c):
    return QColor((c >> 16) & 255, (c >> 8) & 255, c & 255)


def guess_font_index(fontname):
    n = (fontname or "").lower()
    if any(k in n for k in ("times", "serif", "roman", "georgia", "garamond", "cambria")) and "sans" not in n:
        return 1
    if any(k in n for k in ("courier", "mono", "consol")):
        return 2
    return 0


class TextDialog(QDialog):
    def __init__(self, parent, title, text="", size=11.0, font_idx=0, bold=False, color=QColor("black")):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.color = QColor(color)
        form = QFormLayout(self)
        self.text = QPlainTextEdit(text)
        self.text.setMinimumSize(420, 90)
        form.addRow("Text:", self.text)
        self.size = QSpinBox()
        self.size.setRange(4, 200)
        self.size.setValue(max(4, int(round(size))))
        form.addRow("Font size:", self.size)
        self.font = QComboBox()
        for f in FONTS:
            self.font.addItem(f[0])
        self.font.setCurrentIndex(font_idx)
        form.addRow("Font:", self.font)
        self.bold = QCheckBox("Bold")
        self.bold.setChecked(bold)
        form.addRow("", self.bold)
        self.color_btn = QPushButton()
        self.color_btn.clicked.connect(self.pick_color)
        self.paint_color()
        form.addRow("Color:", self.color_btn)
        note = QLabel("<span style='color:gray'>Built-in fonts support Latin letters incl. ä ö ü ß à è é.</span>")
        form.addRow(note)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    def paint_color(self):
        self.color_btn.setText(self.color.name())
        self.color_btn.setStyleSheet(f"background:{self.color.name()}; color:"
                                     f"{'white' if self.color.lightness() < 128 else 'black'}")

    def pick_color(self):
        c = QColorDialog.getColor(self.color, self, "Text color")
        if c.isValid():
            self.color = c
            self.paint_color()

    def values(self):
        f = FONTS[self.font.currentIndex()]
        return dict(text=self.text.toPlainText(), size=float(self.size.value()),
                    fontname=f[2] if self.bold.isChecked() else f[1],
                    color=(self.color.redF(), self.color.greenF(), self.color.blueF()))


def redact_text(page, rect_unrot, fill=False):
    page.add_redact_annot(rect_unrot, fill=fill)
    kw = {"images": fitz.PDF_REDACT_IMAGE_NONE}
    if hasattr(fitz, "PDF_REDACT_LINE_ART_NONE"):
        kw["graphics"] = fitz.PDF_REDACT_LINE_ART_NONE
    if fill is not False:
        kw = {}
    page.apply_redactions(**kw)


def insert_text_upright(page, point_displayed, v):
    """Insert text so it reads normally on the page as displayed."""
    p = fitz.Point(point_displayed.x(), point_displayed.y()) * page.derotation_matrix
    page.insert_text(p, v["text"], fontsize=v["size"], **pdf_safe.font_kw(v["fontname"]), color=v["color"],
                     rotate=page.rotation)




# --------------------------------------------------------------------------
# Recent files
# --------------------------------------------------------------------------
def recent_pdfs():
    v = settings().value("recent_pdfs", [])
    if isinstance(v, str):
        v = [v]
    return [p for p in (v or []) if os.path.isfile(p)][:30]


def add_recent_pdf(path):
    path = os.path.abspath(path)
    items = [p for p in recent_pdfs() if p != path]
    items.insert(0, path)
    settings().setValue("recent_pdfs", items[:30])


# --------------------------------------------------------------------------
# One open PDF = one tab
# --------------------------------------------------------------------------
class ShrinkJob(QThread):
    finished_job = Signal(object, int, int, str)     # bytes or None, before, after, message

    def __init__(self, data, preset, gs):
        super().__init__()
        self.data, self.preset, self.gs = data, preset, gs

    def run(self):
        _, gs_setting, dpi, quality = self.preset
        tmp = tempfile.mkdtemp(prefix="digonto_")
        try:
            src = os.path.join(tmp, "src.pdf")
            with open(src, "wb") as f:
                f.write(self.data)
            results, errors = [], []
            a = os.path.join(tmp, "a.pdf")
            try:
                shrink_with_pymupdf(src, a, dpi, quality)
                results.append(a)
            except Exception as e:
                errors.append(str(e))
            if self.gs:
                b = os.path.join(tmp, "b.pdf")
                try:
                    shrink_with_gs(self.gs, src, b, gs_setting)
                    results.append(b)
                except Exception as e:
                    errors.append(str(e))
            results = [r for r in results if os.path.exists(r) and os.path.getsize(r) > 0]
            if not results:
                self.finished_job.emit(None, len(self.data), 0, "; ".join(errors)[:200])
                return
            best = min(results, key=os.path.getsize)
            with open(best, "rb") as f:
                out = f.read()
            self.finished_job.emit(out, len(self.data), len(out), "")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


def thumb_icon(pm):
    """the same picture when selected (no pink tint over the page)"""
    ic = QIcon()
    ic.addPixmap(pm, QIcon.Normal)
    ic.addPixmap(pm, QIcon.Selected)
    return ic


class PageBox(QLineEdit):
    """the page number box in the bottom bar (type a number, Enter)"""
    valueChanged = Signal(int)

    def __init__(self):
        super().__init__()
        from PySide6.QtGui import QIntValidator
        self._max = 1
        self._val = 1
        self.validator_ = QIntValidator(1, 1, self)
        self.setValidator(self.validator_)
        self.setAlignment(Qt.AlignCenter)
        self.setFixedWidth(46)
        self.setStyleSheet("QLineEdit{border:1px solid #E7E1F0; border-radius:6px; padding:2px; color:#2A1241;"
                           " font-weight:600; background:white;}")
        self.editingFinished.connect(self._typed)

    def setMinimum(self, v):
        pass

    def setMaximum(self, v):
        self._max = max(1, v)
        self.validator_.setTop(self._max)

    def maximum(self):
        return self._max

    def value(self):
        return self._val

    def setValue(self, v):
        v = max(1, min(self._max, int(v)))
        changed = v != self._val
        self._val = v
        self.setText(str(v))
        if changed and not self.signalsBlocked():
            self.valueChanged.emit(v)

    def _typed(self):
        try:
            self.setValue(int(self.text()))
        except ValueError:
            self.setText(str(self._val))


class PageList(QListWidget):
    """page thumbnails: drag to change the order"""
    reordered = Signal(list)              # new order as a list of old page numbers

    def __init__(self):
        super().__init__()
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)

    def dropEvent(self, e):
        before = [self.item(i).data(Qt.UserRole) for i in range(self.count())]
        super().dropEvent(e)
        after = [self.item(i).data(Qt.UserRole) for i in range(self.count())]
        if after != before and sorted(after) == sorted(before):
            self.reordered.emit(after)


class DocTab(QWidget):
    changed = Signal()                    # name or modified state changed
    openBytes = Signal(bytes, str)        # open a new tab with these bytes (shrunk copy ...)

    EDIT_MODES = [("edit", "Edit text", "Click a line of text to change it"),
                  ("add", "Add text", "Click where new text should start - then click again for the next one. Double-click any text to edit it"),
                  ("delete", "Delete text", "Drag a box: text inside is removed (pictures stay)"),
                  ("whiteout", "White-out", "Drag a box: everything inside is covered with white")]

    def __init__(self, doc, path=None, name=None, modified=False):
        super().__init__()
        self.doc = doc
        self.path = path
        self.name = name or (os.path.basename(path) if path else "Untitled.pdf")
        self.modified = modified
        self.password = None
        self.undo_stack = []
        self.cur = 0
        self.tool = None
        self.pending = {}                 # page -> [SignatureItem] not yet written into the PDF
        self.lines = []
        self.hl = None
        self.edit_mode = "edit"
        self.job = None
        self._thumb_next = 0
        self._wheel_t = 0.0

        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # left: the page list (hidden until the Pages icon is pressed)
        self.pages_box = QFrame()
        self.pages_box.setObjectName("pagesBox")
        self.pages_box.setFixedWidth(170)
        self.pages_box.setStyleSheet("#pagesBox{background:#FBFAFD; border:none; border-right:1px solid #E7E1F0;}")
        pv = QVBoxLayout(self.pages_box)
        pv.setContentsMargins(10, 10, 6, 6)
        pv.setSpacing(4)
        t = QLabel("PAGES")
        t.setStyleSheet("color:#6B5F86; font-size:11px; font-weight:700; letter-spacing:1px;")
        pv.addWidget(t)
        h = QLabel("Drag to move · Shift/Ctrl to select more · right-click for more")
        h.setWordWrap(True)
        h.setStyleSheet("color:#6B5F86; font-size:10.5px;")
        pv.addWidget(h)
        self.thumbs = PageList()
        self.thumbs.setIconSize(QSize(110, 150))
        self.thumbs.setSpacing(3)
        self.thumbs.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.thumbs.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.thumbs.setStyleSheet("QListWidget{background:transparent; border:none; outline:none;}"
                                  "QListWidget::item{padding:3px; color:#6B5F86; border-radius:6px;}"
                                  "QListWidget::item:selected{background:#FCE4EA; color:#2A1241;}")
        self.thumbs.currentRowChanged.connect(self._thumb_clicked)
        self.thumbs.reordered.connect(self._reordered)
        self.thumbs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.thumbs.customContextMenuRequested.connect(self._page_menu)
        pv.addWidget(self.thumbs, 1)
        self.pages_box.hide()
        root.addWidget(self.pages_box)

        # centre: page (or the before/after preview) and the bottom bar
        centre = QWidget()
        cl = QVBoxLayout(centre)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        self.strip = QFrame()                 # old tool strip, kept for the tool code (it is never shown)
        self.strip_lay = QHBoxLayout(self.strip)
        self.strip.hide()
        self.view = PageView()
        self.view.setBackgroundBrush(QColor("#EEE9F4"))
        self.view.setFrameShape(QFrame.NoFrame)
        self.view.zoom_request = self.zoom_by
        self.view.flip.connect(self._flip)
        self.view.clicked.connect(self.on_click)
        self.view.doubleClicked.connect(self.on_double_click)
        self.view.rectSelected.connect(self.on_rect)
        self.view.hovered.connect(self.on_hover)
        self.compare = panel.BeforeAfter()
        self.compare.closed.connect(self.close_preview)
        self.centre_stack = QStackedWidget()
        self.centre_stack.addWidget(self.view)
        self.centre_stack.addWidget(self.compare)
        cl.addWidget(self.centre_stack, 1)
        bar = QFrame()
        bar.setObjectName("bar")
        bar.setFixedHeight(40)
        bar.setStyleSheet("#bar{background:white; border-top:1px solid #E7E1F0;}"
                          "#bar QToolButton{border:none; border-radius:7px; padding:3px 8px; color:#2A1241;"
                          " font-size:14px; font-weight:600;}"
                          "#bar QToolButton:hover{background:#F1EAF8;}"
                          "#bar QToolButton#fit{color:#6B5F86; font-size:13px;}"
                          "#bar QLabel{color:#2A1241; font-size:13px; font-weight:600;}"
                          "#bar QSpinBox{border:1px solid #E7E1F0; border-radius:6px; padding:2px 4px;}")
        bl = QHBoxLayout(bar)
        bl.setContentsMargins(12, 0, 12, 0)
        bl.setSpacing(4)

        def tb(text, tip, slot, name=None):
            b = QToolButton()
            b.setText(text)
            b.setToolTip(tip)
            b.setCursor(Qt.PointingHandCursor)
            if name:
                b.setObjectName(name)
            b.clicked.connect(slot)
            bl.addWidget(b)
            return b
        self.prev_btn = tb("‹", "Previous page (Page Up)", lambda: self.goto(self.cur - 1))
        self.page_spin = PageBox()
        bl.addWidget(self.page_spin)
        self.page_lbl = QLabel()
        bl.addWidget(self.page_lbl)
        self.next_btn = tb("›", "Next page (Page Down)", lambda: self.goto(self.cur + 1))
        bl.addStretch()
        self.status = QLabel("")
        self.status.setStyleSheet("color:#6B5F86; font-weight:400;")
        bl.addWidget(self.status)
        bl.addSpacing(14)
        tb("−", "Zoom out (Ctrl -)", lambda: self.zoom_by(1 / 1.2))
        self.zoom_lbl = QLabel("100 %")
        self.zoom_lbl.setFixedWidth(52)
        self.zoom_lbl.setAlignment(Qt.AlignCenter)
        bl.addWidget(self.zoom_lbl)
        tb("+", "Zoom in (Ctrl +)", lambda: self.zoom_by(1.2))
        bl.addSpacing(10)
        tb("Fit width", "The page fills the width (Ctrl 0)", self.fit_width, "fit")
        tb("Fit page", "The whole page is visible", self.fit_page, "fit")
        cl.addWidget(bar)
        root.addWidget(centre, 1)

        self.page_spin.valueChanged.connect(lambda v: self.goto(v - 1))
        QShortcut(QKeySequence.Delete, self.view, activated=self.remove_selected_sig)
        QShortcut(QKeySequence(Qt.Key_Backspace), self.view, activated=self.remove_selected_sig)
        QShortcut(QKeySequence.Delete, self.thumbs, activated=self.delete_pages)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self._fill_thumbs)
        self._fitted = False
        self.rebuild()

    def showEvent(self, e):
        super().showEvent(e)
        if not self._fitted:
            self._fitted = True
            QTimer.singleShot(0, self.fit_page if self.count() > 0 else self.fit_width)

    # ---- basics
    def title(self):
        return self.name + (" *" if self.modified else "")

    def set_modified(self, flag):
        self.modified = flag
        self.changed.emit()

    def page(self):
        return self.doc[self.cur]

    def count(self):
        return self.doc.page_count

    def push_undo(self):
        self.undo_stack.append(self.doc.tobytes())
        if len(self.undo_stack) > 25:
            self.undo_stack.pop(0)
        self.set_modified(True)

    def undo(self):
        for items in self.pending.values():          # first take back signatures that are only placed
            if items:
                it = items.pop()
                if it.scene():
                    it.scene().removeItem(it)
                return
        if not self.undo_stack:
            return
        data = self.undo_stack.pop()
        self.doc = fitz.open(stream=data, filetype="pdf")
        self.set_modified(True)
        self.rebuild(keep=True)

    # ---- page list and viewer
    def rebuild(self, keep=True, select=None):
        """after the PDF changed: page list, spin box, viewer"""
        n = self.count()
        self.cur = min(self.cur if keep else 0, n - 1)
        self.thumbs.blockSignals(True)
        self.thumbs.clear()
        for i in range(n):
            it = QListWidgetItem(str(i + 1))
            it.setData(Qt.UserRole, i)
            it.setTextAlignment(Qt.AlignHCenter)
            it.setSizeHint(QSize(136, 176))
            self.thumbs.addItem(it)
        self.thumbs.blockSignals(False)
        self._thumb_next = 0
        self.timer.start(0)
        self.page_spin.blockSignals(True)
        self.page_spin.setMaximum(n)
        self.page_spin.setValue(self.cur + 1)
        self.page_spin.blockSignals(False)
        self.page_lbl.setText(f"/ {n}")
        self.changed.emit()
        self.prev_btn.setEnabled(n > 1)
        self.next_btn.setEnabled(n > 1)
        self.page_spin.setEnabled(n > 1)
        self._select_rows(select if select else [self.cur])
        self.show_current()

    def _select_rows(self, rows):
        self.thumbs.blockSignals(True)
        self.thumbs.clearSelection()
        for r in rows:
            if 0 <= r < self.thumbs.count():
                self.thumbs.item(r).setSelected(True)
        self.thumbs.setCurrentRow(self.cur, QItemSelectionModel.NoUpdate)
        self.thumbs.blockSignals(False)

    def _fill_thumbs(self):
        t0 = time.time()
        while self._thumb_next < self.thumbs.count() and time.time() - t0 < 0.03:
            i = self._thumb_next
            self._thumb_next += 1
            try:
                page = self.doc[i]
                z = 110.0 / max(page.rect.width, 1)
                pm = render_page(page, min(z, 150.0 / max(page.rect.height, 1)), 1.0)
                self.thumbs.item(i).setIcon(thumb_icon(pm))
            except Exception:
                pass
        if self._thumb_next >= self.thumbs.count():
            self.timer.stop()

    def _thumb_clicked(self, row):
        self._last_sel = self.selected_pages()
        if row >= 0 and row != self.cur:
            self.goto(row, from_list=True)

    def goto(self, n, from_list=False, at_end=False):
        if n < 0 or n >= self.count() or n == self.cur:
            self.page_spin.blockSignals(True)
            self.page_spin.setValue(self.cur + 1)
            self.page_spin.blockSignals(False)
            return
        self._leave_page()
        self.cur = n
        self.page_spin.blockSignals(True)
        self.page_spin.setValue(n + 1)
        self.page_spin.blockSignals(False)
        if not from_list:
            self._select_rows([n])
            self.thumbs.scrollToItem(self.thumbs.item(n))
        self.show_current()
        sb = self.view.verticalScrollBar()
        sb.setValue(sb.maximum() if at_end else sb.minimum())

    def _flip(self, d):
        if d > 0 and self.cur < self.count() - 1:
            self.goto(self.cur + 1)
        elif d < 0 and self.cur > 0:
            self.goto(self.cur - 1, at_end=True)

    def _leave_page(self):
        sc = self.view.scene()
        for it in self.pending.get(self.cur, []):
            if it.scene() is sc:
                sc.removeItem(it)

    def show_current(self):
        self.view.show_page(self.page())
        sc = self.view.scene()
        for it in self.pending.get(self.cur, []):
            if it.scene() is None:
                sc.addItem(it)
        if self.tool == "edit":
            self.load_lines()

    def zoom_by(self, f):
        self.view.set_zoom(self.view.zoom * f, self.page())
        self._zoom_text()

    def fit_width(self):
        self.view.fit_width(self.page())
        self._zoom_text()

    def fit_page(self):
        self.view.fit_page(self.page())
        self._zoom_text()

    def _zoom_text(self):
        self.zoom_lbl.setText(f"{round(self.view.zoom * 100 / 1.3333)} %")    # 100 % = real size on screen

    # ---- page list
    def toggle_pages(self, show=None):
        show = not self.pages_box.isVisible() if show is None else show
        self.pages_box.setVisible(show)
        if show:
            self.thumbs.scrollToItem(self.thumbs.item(self.cur))
        return show

    def _reordered(self, order):
        self.apply_pending()
        self.push_undo()
        cur_old = self.cur
        self.doc.select(order)
        self.cur = order.index(cur_old)
        sel = [order.index(r) for r in (self._last_sel or [cur_old]) if r in order]
        self.after_structure(select=sel)

    _last_sel = None

    def _page_menu(self, pos):
        it = self.thumbs.itemAt(pos)
        if it is not None and not it.isSelected():
            self.thumbs.clearSelection()
            it.setSelected(True)
            self.goto(self.thumbs.row(it), from_list=True)
        n = len(self.selected_pages())
        m = QMenu(self)
        m.addAction(f"Turn right ({n})" if n > 1 else "Turn right", lambda: self.rotate_pages(90))
        m.addAction(f"Turn left ({n})" if n > 1 else "Turn left", lambda: self.rotate_pages(-90))
        m.addSeparator()
        m.addAction("Copy (duplicate)", self.duplicate_pages)
        m.addAction("Insert blank page after", self.insert_blank)
        m.addAction("Insert pages from another PDF …", self.insert_pdf)
        m.addSeparator()
        m.addAction("Move up", lambda: self.move_pages(-1))
        m.addAction("Move down", lambda: self.move_pages(1))
        m.addAction("Save as a new PDF …", self.extract_pages)
        m.addSeparator()
        a = m.addAction(f"Delete ({n})" if n > 1 else "Delete", self.delete_pages)
        a.setEnabled(n < self.count())
        m.exec(self.thumbs.mapToGlobal(pos))

    def duplicate_pages(self):
        rows = self.selected_pages()
        self.apply_pending()
        self.push_undo()
        order = list(range(self.count()))
        last = rows[-1]
        order[last + 1:last + 1] = rows
        self.doc.select(order)
        new = list(range(last + 1, last + 1 + len(rows)))
        self.cur = new[0]
        self.after_structure(select=new)

    def reverse_pages(self):
        self.apply_pending()
        self.push_undo()
        self.doc.select(list(range(self.count() - 1, -1, -1)))
        self.cur = self.count() - 1 - self.cur
        self.after_structure()

    # ---- before / after preview
    def show_preview(self, before, after, note):
        self.compare.set_images(before, after, note)
        self.centre_stack.setCurrentIndex(1)

    def close_preview(self):
        self.centre_stack.setCurrentIndex(0)

    def selected_pages(self):
        rows = sorted(self.thumbs.row(i) for i in self.thumbs.selectedItems())
        return rows or [self.cur]

    # ---- tool strip (the "sub tab" under the right-hand panel)
    def _btn(self, text, slot, tip=""):
        b = QPushButton(text)
        if tip:
            b.setToolTip(tip)
        b.clicked.connect(slot)
        self.strip_lay.addWidget(b)
        return b

    def _label(self, text, grey=False):
        lb = QLabel(text)
        lb.setWordWrap(True)
        if grey:
            lb.setStyleSheet("color:#6B5F86; font-size:12px;")
        self.strip_lay.addWidget(lb)
        return lb

    @staticmethod
    def _clear_layout(lay):
        while lay is not None and lay.count():
            item = lay.takeAt(0)
            if item.widget() is not None:
                item.widget().hide()
                item.widget().deleteLater()
            elif item.layout() is not None:
                DocTab._clear_layout(item.layout())

    def set_tool(self, name, layout=None):
        """name: pages | sign | edit | redact | None.  The controls go into `layout` (the tools panel)."""
        prev, self.tool = self.tool, name          # set first: saving below must not open the tool a second time
        self.apply_pending()
        if layout is not None:
            self.strip_lay = layout
        if name is not None and getattr(self, "strip_lay", None) is not None:
            self._clear_layout(self.strip_lay)      # never two sets of buttons in the panel
        self.view.mode = None
        self.view.unsetCursor()
        self.clear_hl()
        self.close_preview()
        if name is None:
            return
        getattr(self, "_tool_" + name)()

    def _tool_pages(self):
        self.toggle_pages(True)
        self._label("Drag pages in the list on the left to move them. Shift or Ctrl selects more pages.", grey=True)
        self._btn("Move up", lambda: self.move_pages(-1), "Move the selected page(s) one place up")
        self._btn("Move down", lambda: self.move_pages(1), "Move the selected page(s) one place down")
        self._btn("Copy (duplicate)", self.duplicate_pages, "Copy the selected page(s)")
        self._btn("Insert blank page", self.insert_blank, "An empty page after the selected page")
        self._btn("Insert pages from PDF …", self.insert_pdf, "Insert pages of another PDF after the selected page")
        self._btn("Reverse the order", self.reverse_pages, "Last page first")
        self._btn("Delete selected", self.delete_pages)

    def _tool_sign(self):
        self._btn("Draw new...", self.sig_draw)
        self._btn("Add image...", self.sig_image)
        self.sig_combo = QComboBox()
        self.sig_combo.setIconSize(QSize(110, 36))
        self.sig_combo.setMinimumWidth(150)
        self.strip_lay.addWidget(self.sig_combo)
        self._sig_load_library()
        self._btn("Place on page", self.sig_place, "Put the chosen signature on this page, then drag / resize it")
        self._btn("Remove", self.remove_selected_sig, "Remove the selected signature from the page")
        self._btn("Forget", self.sig_forget, "Delete the chosen signature from your list")
        self._btn("Apply", self.apply_pending, "Write the placed signatures into the PDF")
        self._label("Drag to move, blue corner to resize. Apply writes it into the PDF.", grey=True)

    def _tool_edit(self):
        self.mode_group = QButtonGroup(self)
        for i, (key, label, tip) in enumerate(self.EDIT_MODES):
            rb = QRadioButton(label)
            rb.setToolTip(tip)
            rb.setChecked(key == self.edit_mode)
            self.mode_group.addButton(rb, i)
            self.strip_lay.addWidget(rb)
        self.mode_group.idClicked.connect(self._mode_changed)
        self.hint = self._label("", grey=True)
        self._mode_changed(next(i for i, m in enumerate(self.EDIT_MODES) if m[0] == self.edit_mode))
        self.load_lines()

    def _tool_redact(self):
        self._label("Draw a box over anything that must disappear. The text and pictures under it are really "
                    "removed - not only covered.", grey=True)
        self.view.mode = "rect"
        self.view.setCursor(Qt.CrossCursor)

    def _tool_shrink(self):
        self._label("<b>Shrink</b>")
        self.shrink_combo = QComboBox()
        for p in PRESETS:
            self.shrink_combo.addItem(p[0])
        self.shrink_combo.setCurrentIndex(1)
        self.strip_lay.addWidget(self.shrink_combo)
        self.shrink_btn = self._btn("Shrink to a new file", self.shrink_now,
                                    "Makes a smaller copy and opens it in a new tab. This file stays as it is.")
        self.shrink_lbl = self._label("", grey=True)

    def _tool_add(self):
        self._label("<b>Add files</b>")
        self._btn("Add pictures, Word, Excel, text as pages...", self.add_files,
                  "Converts the files and inserts them after the current page")

    def _tool_more(self):
        self._label("<b>More</b>")
        self._btn("Page numbers...", self.page_numbers)
        self._btn("Watermark...", self.watermark)
        self._btn("Password...", self.set_password, "The password is applied when you save")
        self._btn("Document info", self.doc_info)

    # ---- page tools
    def after_structure(self, select=None):
        self.set_modified(True)
        self.rebuild(keep=True, select=select)

    def insert_pdf(self):
        path, _ = QFileDialog.getOpenFileName(self, "Insert pages from PDF", last_dir(), PDF_FILTER)
        if not path:
            return
        src = open_pdf(self, path)
        if not src:
            return
        remember_dir(path)
        pages = list(range(src.page_count))
        if src.page_count > 1:
            while True:
                txt, ok = QInputDialog.getText(
                    self, "Which pages?", f"{os.path.basename(path)} has {src.page_count} pages.\n"
                    "Which pages to insert?  Examples:  all   |   3   |   1-4, 7, 9-", text="all")
                if not ok:
                    return
                try:
                    pages = parse_ranges(txt, src.page_count)
                    break
                except ValueError as e:
                    QMessageBox.warning(self, APP_NAME, str(e))
        self.apply_pending()
        self.push_undo()
        at = max(self.selected_pages()) + 1
        for k, p in enumerate(pages):
            self.doc.insert_pdf(src, from_page=p, to_page=p, start_at=at + k)
        self.cur = at
        self.after_structure(select=list(range(at, at + len(pages))))

    def insert_blank(self):
        self.apply_pending()
        self.push_undo()
        r = self.page().rect
        at = max(self.selected_pages()) + 1
        self.doc.new_page(pno=at, width=r.width, height=r.height)
        self.cur = at
        self.after_structure(select=[at])

    def move_pages(self, step):
        rows = self.selected_pages()
        n = self.count()
        if (step < 0 and rows[0] == 0) or (step > 0 and rows[-1] == n - 1):
            return
        self.apply_pending()
        self.push_undo()
        for r in (rows if step < 0 else list(reversed(rows))):
            if step < 0:
                self.doc.move_page(r, r - 1)
            else:
                self.doc.move_page(r, r + 2 if r + 2 < n else -1)
        self.cur = min(max(self.cur + step, 0), n - 1)
        self.after_structure(select=[r + step for r in rows])

    def rotate_pages(self, deg):
        self.apply_pending()
        rows = self.selected_pages()
        self.push_undo()
        for r in rows:
            p = self.doc[r]
            p.set_rotation((p.rotation + deg) % 360)
        self.after_structure(select=rows)

    def delete_pages(self):
        rows = self.selected_pages()
        if len(rows) >= self.count():
            QMessageBox.information(self, APP_NAME, "A PDF needs at least one page.")
            return
        self.apply_pending()
        self.push_undo()
        self.doc.delete_pages(rows)
        self.cur = min(rows[0], self.count() - 1)
        self.after_structure(select=[self.cur])

    def extract_pages(self):
        rows = self.selected_pages()
        out = fitz.open()
        for r in rows:
            out.insert_pdf(self.doc, from_page=r, to_page=r)
        path, _ = QFileDialog.getSaveFileName(self, "Save the extracted pages as",
                                              default_name(self.path, "pages"), PDF_FILTER)
        if path and save_doc(self, out, path):
            self.status.setText(f"Saved {len(rows)} page(s)")

    # ---- add files
    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Choose files to add", last_dir(), CREATE_FILTER)
        if not files:
            return
        remember_dir(files[0])
        self.apply_pending()
        QApplication.setOverrideCursor(Qt.WaitCursor)
        added, errors = 0, []
        try:
            self.push_undo()
            at = max(self.selected_pages()) + 1
            for f in files:
                if not conv.kind_of(f):
                    errors.append(f"{os.path.basename(f)}: not supported")
                    continue
                try:
                    d = conv.convert_file(f, "A4", 7 * 72 / 25.4, False, True)
                    self.doc.insert_pdf(d, start_at=at + added)
                    added += d.page_count
                except Exception as e:
                    errors.append(f"{os.path.basename(f)}: {e}")
        finally:
            QApplication.restoreOverrideCursor()
        if added:
            self.cur = at
            self.after_structure(select=list(range(at, at + added)))
        else:
            self.undo_stack.pop()
        if errors:
            QMessageBox.warning(self, APP_NAME, "Problems:\n" + "\n".join(errors[:10]))

    # ---- more
    def page_numbers(self):
        pos, ok = QInputDialog.getItem(self, "Page numbers", "Where?",
                                       ["Bottom centre", "Bottom right", "Bottom left", "Top centre", "Top right"], 0, False)
        if not ok:
            return
        fmt, ok = QInputDialog.getItem(self, "Page numbers", "Look:", ["1", "Page 1", "1 / N", "Page 1 of N"], 2, False)
        if not ok:
            return
        self.apply_pending()
        self.push_undo()
        n = self.count()
        for i in range(n):
            p = self.doc[i]
            txt = {"1": f"{i + 1}", "Page 1": f"Page {i + 1}", "1 / N": f"{i + 1} / {n}",
                   "Page 1 of N": f"Page {i + 1} of {n}"}[fmt]
            r = p.rect
            w = pdf_safe.text_length(txt, "helv", 10)
            x = {"centre": (r.width - w) / 2, "right": r.width - w - 36, "left": 36}[pos.split()[1]]
            y = r.height - 24 if pos.startswith("Bottom") else 36
            insert_text_upright(p, QPointF(x, y), dict(text=txt, size=10.0, fontname="helv", color=(0.2, 0.2, 0.2)))
        self.set_modified(True)
        self.show_current()

    def watermark(self):
        txt, ok = QInputDialog.getText(self, "Watermark", "Text across every page:", text="DRAFT")
        if not ok or not txt.strip():
            return
        self.apply_pending()
        self.push_undo()
        for p in self.doc:
            r = p.rect
            size = min(r.width, r.height) / max(len(txt), 3) * 1.5
            w = pdf_safe.text_length(txt, "hebo", size)
            c = fitz.Point(r.width / 2, r.height / 2)
            start = fitz.Point(c.x - w / 2, c.y + size / 3)
            p.insert_text(start * p.derotation_matrix, txt, fontsize=size, **pdf_safe.font_kw("hebo"), color=(0.6, 0.6, 0.6),
                          fill_opacity=0.35, morph=(c * p.derotation_matrix, fitz.Matrix(45)), overlay=True)
        self.set_modified(True)
        self.show_current()

    def set_password(self):
        pw, ok = QInputDialog.getText(self, "Password", "Password needed to open this PDF\n(empty = remove the password):",
                                      QLineEdit.Password)
        if not ok:
            return
        self.password = pw or None
        self.set_modified(True)
        self.status.setText("Password will be set when you save" if pw else "Password removed on next save")

    def doc_info(self):
        m = self.doc.metadata or {}
        size = human(os.path.getsize(self.path)) if self.path and os.path.exists(self.path) else "not saved yet"
        QMessageBox.information(
            self, "Document info",
            f"File: {self.path or '(not saved yet)'}\nSize: {size}\nPages: {self.count()}\n"
            f"Title: {m.get('title') or '-'}\nAuthor: {m.get('author') or '-'}\nCreated with: {m.get('producer') or '-'}")

    # ---- shrink
    def shrink_now(self):
        if self.job and self.job.isRunning():
            return
        self.apply_pending()
        preset = PRESETS[self.shrink_combo.currentIndex()]
        self.shrink_btn.setEnabled(False)
        self.shrink_lbl.setText("Working...")
        self.job = ShrinkJob(self.doc.tobytes(), preset, find_ghostscript())
        self.job.finished_job.connect(self._shrink_done)
        self.job.start()

    def _shrink_done(self, data, before, after, msg):
        try:
            self.shrink_btn.setEnabled(True)
            self.shrink_lbl.setText("")
        except RuntimeError:
            pass                              # strip was closed meanwhile
        if data is None:
            QMessageBox.warning(self, APP_NAME, "Could not shrink this file.\n" + msg)
            return
        if after >= before:
            QMessageBox.information(self, APP_NAME, "This PDF is already small - nothing to gain.")
            return
        base = os.path.splitext(self.name)[0]
        self.status.setText(f"{human(before)} -> {human(after)}  ({100 - after * 100 // before}% smaller)")
        self.openBytes.emit(data, base + "_small.pdf")

    # ---- signatures
    def _sig_load_library(self):
        for f in sorted(glob.glob(os.path.join(signatures_dir(), "*.png"))):
            img = QImage(f)
            if not img.isNull():
                with open(f, "rb") as fh:
                    self._sig_add(img, fh.read(), f, select=False)

    def _sig_add(self, img, png, path=None, select=True):
        thumb = QPixmap.fromImage(img).scaled(110, 36, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        canvas = QPixmap(110, 36)
        canvas.fill(Qt.white)
        p = QPainter(canvas)
        p.drawPixmap((110 - thumb.width()) // 2, (36 - thumb.height()) // 2, thumb)
        p.end()
        self.sig_combo.addItem(QIcon(canvas), f"Signature {self.sig_combo.count() + 1}", (img, png, path))
        if select:
            self.sig_combo.setCurrentIndex(self.sig_combo.count() - 1)

    def _sig_save(self, png):
        path = os.path.join(signatures_dir(), f"signature_{int(time.time() * 1000)}.png")
        with open(path, "wb") as f:
            f.write(png)
        return path

    def sig_draw(self):
        dlg = DrawSignatureDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        img = dlg.pad.to_image()
        if img is None:
            return
        png = qimage_png_bytes(img)
        path = self._sig_save(png) if dlg.keep.isChecked() else None
        self._sig_add(img, png, path)
        self.sig_place()

    def sig_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose signature image", last_dir(), IMG_FILTER)
        if not path:
            return
        remember_dir(path)
        img = QImage(path)
        if img.isNull():
            QMessageBox.warning(self, APP_NAME, "Could not read this image.")
            return
        img = img.convertToFormat(QImage.Format_ARGB32)
        if QMessageBox.question(self, APP_NAME, "Make the white background transparent?\n"
                                "(Choose Yes for a scanned or photographed signature on white paper.)") == QMessageBox.Yes:
            img = self._whites_to_transparent(img)
        png = qimage_png_bytes(img)
        self._sig_add(img, png, self._sig_save(png))
        self.sig_place()

    @staticmethod
    def _whites_to_transparent(img):
        img = img.copy()
        for y in range(img.height()):
            for x in range(img.width()):
                c = img.pixelColor(x, y)
                light = min(c.red(), c.green(), c.blue())
                if light > 200:
                    c.setAlpha(max(0, int(c.alpha() * (255 - light) / 55)))
                    img.setPixelColor(x, y, c)
        return img

    def sig_forget(self):
        i = self.sig_combo.currentIndex()
        if i < 0:
            return
        _, _, path = self.sig_combo.itemData(i)
        if path and os.path.exists(path):
            os.remove(path)
        self.sig_combo.removeItem(i)

    def sig_place(self):
        i = self.sig_combo.currentIndex()
        if i < 0:
            QMessageBox.information(self, APP_NAME, "Draw a signature or add an image first.")
            return
        img, png, _ = self.sig_combo.itemData(i)
        page = self.page()
        sig = SignatureItem(img, png, min(160.0, page.rect.width * 0.35))
        c = self.view.mapToScene(self.view.viewport().rect().center())
        sig.setPos(min(max(0, c.x() - sig.w / 2), page.rect.width - sig.w),
                   min(max(0, c.y() - sig.h / 2), page.rect.height - sig.h))
        n = len(self.pending.get(self.cur, []))
        sig.moveBy(12 * (n % 5), 12 * (n % 5))      # each new one a little lower right, so it is not hidden under the last
        sc = self.view.scene()
        sc.addItem(sig)
        sc.clearSelection()
        sig.setSelected(True)
        self.view.ensureVisible(sig)
        self.pending.setdefault(self.cur, []).append(sig)
        self.set_modified(True)

    def remove_selected_sig(self):
        sc = self.view.scene()
        for it in list(sc.selectedItems()):
            if isinstance(it, SignatureItem):
                sc.removeItem(it)
                for lst in self.pending.values():
                    if it in lst:
                        lst.remove(it)

    def apply_pending(self):
        """write placed signatures into the PDF"""
        items = [(p, it) for p, lst in self.pending.items() for it in lst]
        if not items:
            self.pending = {}
            return
        sc = self.view.scene()
        for _, it in items:
            if it.scene() is sc:
                sc.removeItem(it)
        self.pending = {}
        snapshot = self.doc.tobytes()
        self.undo_stack.append(snapshot)
        for p, it in items:
            place_image_on_page(self.doc[p], it.rect_pt(), it.png)
        self.set_modified(True)
        self.show_current()

    # ---- text editing
    def _mode_changed(self, i):
        self.edit_mode = self.EDIT_MODES[i][0]
        self.hint.setText(self.EDIT_MODES[i][2])
        self.view.mode = "rect" if self.edit_mode in ("delete", "whiteout") else "click"
        self.view.setCursor(Qt.IBeamCursor if self.edit_mode in ("edit", "add") else Qt.CrossCursor)
        self.clear_hl()

    def load_lines(self):
        self.clear_hl()
        self.lines = []
        page = self.page()
        rm = page.rotation_matrix
        d = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
        for b in d.get("blocks", []):
            if b.get("type") != 0:
                continue
            for ln in b.get("lines", []):
                if not "".join(s["text"] for s in ln["spans"]).strip():
                    continue
                unrot = fitz.Rect(ln["bbox"])
                disp = unrot * rm
                self.lines.append((QRectF(disp.x0, disp.y0, disp.width, disp.height), unrot, ln))

    def clear_hl(self):
        if self.hl is not None:
            if self.hl.scene():
                self.hl.scene().removeItem(self.hl)
            self.hl = None

    def line_at(self, p):
        best = None
        for disp, unrot, ln in self.lines:
            if disp.adjusted(-1, -1, 1, 1).contains(p):
                if best is None or disp.width() * disp.height() < best[0].width() * best[0].height():
                    best = (disp, unrot, ln)
        return best

    def on_hover(self, p):
        if self.tool != "edit" or self.edit_mode not in ("edit", "add"):
            return
        hit = self.line_at(p)
        self.clear_hl()
        if hit:
            self.hl = QGraphicsRectItem(hit[0].adjusted(-1, -1, 1, 1))
            self.hl.setPen(QPen(QColor("#C2395C"), 0))
            self.hl.setBrush(QColor(194, 57, 92, 40))
            self.hl.setZValue(3)
            self.view.scene().addItem(self.hl)

    def on_click(self, p):
        if self.tool != "edit":
            return
        if self.edit_mode == "add":
            # wait a moment: a double-click on existing text means "edit it", not "add new text here"
            t = getattr(self, "_add_timer", None)
            if t is None:
                t = self._add_timer = QTimer(self)
                t.setSingleShot(True)
                t.timeout.connect(lambda: self._add_text(self._add_pos))
            self._add_pos = QPointF(p)
            t.start(QApplication.doubleClickInterval() + 40)
        elif self.edit_mode == "edit":
            hit = self.line_at(p)
            if hit:
                self._edit_line(hit)

    def on_double_click(self, p):
        if self.tool != "edit" or self.edit_mode not in ("add", "edit"):
            return
        t = getattr(self, "_add_timer", None)
        if t is not None:
            t.stop()
        if self.edit_mode == "edit":
            return                      # the first click already opened the editor
        hit = self.line_at(p)
        if hit:
            self._edit_line(hit)
        else:
            self._add_text(p)

    def _add_text(self, p):
        page = self.page()
        dlg = TextDialog(self, "Add text", size=getattr(self, "_last_size", 11.0))
        if dlg.exec() != QDialog.Accepted:
            return
        v = dlg.values()
        if not v["text"].strip():
            return
        self._last_size = v["size"]
        self.push_undo()
        insert_text_upright(page, QPointF(p.x(), p.y() + v["size"] * 0.8), v)
        self._edited()                  # the new text is now a line that can be double-clicked

    def _edit_line(self, hit):
        page = self.page()
        disp, unrot, ln = hit
        spans = ln["spans"]
        s0 = spans[0]
        text = "".join(s["text"] for s in spans)
        dlg = TextDialog(self, "Change text", text, s0["size"], guess_font_index(s0["font"]),
                         bool(s0["flags"] & 16) or "bold" in s0["font"].lower(), int_to_qcolor(s0["color"]))
        dlg.text.selectAll()
        if dlg.exec() != QDialog.Accepted:
            return
        v = dlg.values()
        self.push_undo()
        dx, dy = ln.get("dir", (1, 0))
        if abs(dx) >= abs(dy):
            h = unrot.height
            r = fitz.Rect(unrot.x0 - 0.5, unrot.y0 + h * 0.15, unrot.x1 + 0.5, unrot.y1 - h * 0.15)
        else:
            wd = unrot.width
            r = fitz.Rect(unrot.x0 + wd * 0.15, unrot.y0 - 0.5, unrot.x1 - wd * 0.15, unrot.y1 + 0.5)
        redact_text(page, r)
        if v["text"].strip():
            angle = int(round(math.degrees(math.atan2(-dy, dx)) / 90.0)) * 90 % 360
            page.insert_text(fitz.Point(s0["origin"]), v["text"], fontsize=v["size"],
                             **pdf_safe.font_kw(v["fontname"]), color=v["color"], rotate=angle)
        self._edited()

    def on_rect(self, r):
        if self.tool == "redact":
            page = self.page()
            fr = fitz.Rect(r.x(), r.y(), r.x() + r.width(), r.y() + r.height()) * page.derotation_matrix
            self.push_undo()
            page.add_redact_annot(fr, fill=(0, 0, 0))
            page.apply_redactions()
            self._edited()
            return
        if self.tool != "edit":
            return
        page = self.page()
        fr = fitz.Rect(r.x(), r.y(), r.x() + r.width(), r.y() + r.height()) * page.derotation_matrix
        self.push_undo()
        redact_text(page, fr, fill=(1, 1, 1) if self.edit_mode == "whiteout" else False)
        self._edited()

    def _edited(self):
        self.view.show_page(self.page())
        self.load_lines()
        self._thumb_next = min(self._thumb_next, self.cur)
        try:
            page = self.page()
            self.thumbs.item(self.cur).setIcon(thumb_icon(render_page(page, min(110.0 / page.rect.width, 150.0 / page.rect.height))))
        except Exception:
            pass

    # ---- saving / closing / export
    def save(self, ask_path=False):
        self.apply_pending()
        path = self.path
        if ask_path or not path:
            start = path or os.path.join(last_dir(), self.name)
            p, _ = QFileDialog.getSaveFileName(self, "Save PDF as", start, PDF_FILTER)
            if not p:
                return False
            path = ensure_pdf_ext(p)
        extra = {}
        if self.password:
            extra = dict(encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=self.password, owner_pw=self.password)
        if not save_doc(self, self.doc, path, **extra):
            return False
        self.path = path
        self.name = os.path.basename(path)
        add_recent_pdf(path)
        self.set_modified(False)
        self._font_hint(path)
        return True

    def _font_hint(self, path):
        """upload portals often refuse PDFs whose letters are not in the file: say so, once per file"""
        try:
            bad = pdf_safe.unembedded_fonts(self.doc)
        except Exception:
            return
        if not bad or office_mode() or getattr(self, "_hinted", None) == path:
            return
        self._hinted = path
        QMessageBox.information(self, APP_NAME,
            "Saved.\n\nThis PDF uses fonts that are not inside the file (" + ", ".join(bad[:4]) + ("…" if len(bad) > 4 else "") + ").\n"
            "Official upload websites often refuse such files.\n\n"
            "Before you upload it, use File → Save for upload (Ctrl+Alt+S).")

    def save_for_upload(self):
        """a copy that government upload portals accept (they turn every upload into PDF/A)"""
        self.apply_pending()
        box = QMessageBox(QMessageBox.Question, APP_NAME,
                          "Save a copy that official upload websites accept (government, tax, job or health portals)?\n\n"
                          "Form fields and notes become part of the page, scripts, attachments and bookmarks are "
                          "removed, and any page that could still be refused is turned into a clean picture.", parent=self)
        normal = box.addButton("Save copy", QMessageBox.AcceptRole)
        pic = box.addButton("All pages as pictures (always accepted)", QMessageBox.ActionRole)
        box.addButton("Cancel", QMessageBox.RejectRole)
        box.setDefaultButton(normal)
        box.exec()
        if box.clickedButton() not in (normal, pic):
            return False
        base = os.path.splitext(self.name)[0]
        start = os.path.join(os.path.dirname(self.path) if self.path else last_dir(), base + " (upload).pdf")
        p, _ = QFileDialog.getSaveFileName(self, "Save copy for upload", start, PDF_FILTER)
        if not p:
            return False
        p = ensure_pdf_ext(p)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            data, pics = pdf_safe.upload_copy(self.doc.tobytes(), as_picture=box.clickedButton() is pic, password=self.password)
            with open(p, "wb") as fh:
                fh.write(data)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, APP_NAME, f"Could not make the upload copy:\n{e}")
            return False
        QApplication.restoreOverrideCursor()
        remember_dir(p)
        add_recent_pdf(p)
        size = os.path.getsize(p)
        note = ""
        if pics and box.clickedButton() is not pic:
            note = f"\n\nPage(s) {', '.join(map(str, pics))} had something upload websites refuse, so they are now pictures."
        if size > 9 * 1024 * 1024:
            note += "\n\nThe file is larger than 9 MB, which many upload websites refuse: use Compress first."
        QMessageBox.information(self, APP_NAME, f"Saved for upload:\n{p}\n({human(size)}){note}")
        return True

    def maybe_close(self):
        """True when it is fine to close this tab"""
        if not self.modified and not any(self.pending.values()):
            return True
        box = QMessageBox(QMessageBox.Question, APP_NAME, f"Save the changes to {self.name}?", parent=self)
        save = box.addButton("Save", QMessageBox.AcceptRole)
        box.addButton("Don't save", QMessageBox.DestructiveRole)
        cancel = box.addButton("Cancel", QMessageBox.RejectRole)
        box.setDefaultButton(save)
        box.exec()
        b = box.clickedButton()
        if b is cancel:
            return False
        if b is save:
            return self.save()
        return True

    def export_images(self):
        rows = self.selected_pages() if len(self.selected_pages()) > 1 else list(range(self.count()))
        folder = QFileDialog.getExistingDirectory(self, "Choose a folder for the pictures", last_dir())
        if not folder:
            return
        remember_dir(folder)
        self.apply_pending()
        base = os.path.splitext(self.name)[0]
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            for r in rows:
                self.doc[r].get_pixmap(dpi=200).save(unique_path(os.path.join(folder, f"{base}_page_{r + 1}.png")))
        finally:
            QApplication.restoreOverrideCursor()
        QMessageBox.information(self, APP_NAME, f"Saved {len(rows)} picture(s) in:\n{folder}")


# --------------------------------------------------------------------------
# Main window: icon bar | start screen / open PDFs / batch list | tools
# --------------------------------------------------------------------------
UI_TOOL = {"organise": "pages", "text": "edit", "sign": "sign", "redact": "redact"}


class MainWindow(QMainWindow):
    def __init__(self, files=()):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(make_app_icon())
        self.setAcceptDrops(True)
        self.worker = None
        self.job_ctx = None

        self.home = ui.StartScreen(
            "Open, edit and fix your PDFs - all on this computer",
            [("Add files", "new", self.open_dialog, True), ("Add folder", "open", self.add_folder, False)],
            "or drop PDF files here", recent_pdfs, self.open_path)

        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setTabsClosable(True)
        self.tabs.setMovable(True)
        self.tabs.setStyleSheet("QTabWidget::pane{border:none;} QTabBar{background:#F7F4FB;}"
                                "QTabBar::tab{padding:7px 12px; margin:6px 0 0 6px; border:1px solid transparent;"
                                " border-bottom:none; border-top-left-radius:8px; border-top-right-radius:8px;"
                                " color:#6B5F86; background:transparent;}"
                                "QTabBar::tab:selected{background:white; border-color:#E7E1F0; color:#2A1241;}")
        self.tabs.tabCloseRequested.connect(self.close_tab)
        self.tabs.currentChanged.connect(self._tab_changed)

        self.batch = panel.BatchView()
        self.batch.saveAll.connect(self.batch_save_all)
        self.batch.stop.connect(self.stop_job)
        self.batch.clear.connect(self.close_batch)
        self.batch.addMore.connect(self.batch_add_more)

        self.centre = QStackedWidget()
        self.centre.addWidget(self.home)       # 0
        self.centre.addWidget(self.tabs)       # 1
        self.centre.addWidget(self.batch)      # 2

        self.rail = ui.Rail(
            [("new", "new", "New PDF from pictures, Word, Excel or text (Ctrl+N)", lambda: self.create_dialog()),
             ("open", "open", "Open PDF (Ctrl+O)", self.open_dialog),
             ("save", "save", "Save (Ctrl+S)", lambda: self._cur() and self._cur().save()),
             ("export", "export", "Save as / export", self.export_menu_popup),
             ("print", "print", "Print (Ctrl+P)", self.print_current),
             None,
             ("pen", "pen", "Add or edit text", lambda: self.tools.open_tool("text")),
             ("pages", "pages", "Show or hide the page list", self.toggle_pages)],
            [("settings", "settings", "Settings and help", self.settings_menu)])
        self.rail.button("pages").setCheckable(True)

        self.tools = panel.ToolsPanel()
        self.tools.toolOpened.connect(self.tool_opened)
        self.tools.toolClosed.connect(self.tool_closed)
        self.tools.runClicked.connect(self.run_tool)
        self.tools.stopClicked.connect(self.stop_job)
        self.tools.previewClicked.connect(self.preview_tool)

        central = QWidget()
        cl = QVBoxLayout(central)
        cl.setContentsMargins(0, 0, 0, 0)
        cl.setSpacing(0)
        cl.addWidget(ui.make_header(APP_NAME, "Create, organise, scan, convert, sign and protect PDFs",
                                    make_app_icon(), self.about))
        body = QHBoxLayout()
        body.setSpacing(0)
        body.addWidget(self.rail)
        body.addWidget(self.centre, 1)
        body.addWidget(self.tools)
        cl.addLayout(body, 1)
        self.setCentralWidget(central)

        self._menu()
        self.resize(1280, 860)
        g = settings().value("geometry")
        if g:
            self.restoreGeometry(g)
        QShortcut(QKeySequence(Qt.Key_PageDown), self, activated=lambda: self._page_step(1))
        QShortcut(QKeySequence(Qt.Key_PageUp), self, activated=lambda: self._page_step(-1))
        self._tab_changed()

        files = [f for f in files if os.path.isfile(f) or os.path.isdir(f)]
        for f in files:
            if os.path.isdir(f):
                self.add_folder(f)
            elif f.lower().endswith(".pdf"):
                self.open_path(f)
        others = [f for f in files if os.path.isfile(f) and not f.lower().endswith(".pdf") and conv.kind_of(f)]
        if others:
            QTimer.singleShot(0, lambda: self.create_dialog(others))

    # ---- menu (also gives the keyboard shortcuts)
    def _menu(self):
        mb = self.menuBar()
        fm = mb.addMenu("File")

        def act(menu, text, key, fn):
            a = QAction(text, self)
            if key:
                a.setShortcut(key)
            a.triggered.connect(fn)
            menu.addAction(a)
            return a

        act(fm, "Open PDF...", "Ctrl+O", self.open_dialog)
        act(fm, "Open a folder (many PDFs at once)...", "Ctrl+Shift+O", self.add_folder)
        self.recent_menu = fm.addMenu("Open recent")
        self.recent_menu.aboutToShow.connect(self._fill_recent)
        act(fm, "Create PDF...", "Ctrl+N", lambda: self.create_dialog())
        fm.addSeparator()
        self.act_save = act(fm, "Save", "Ctrl+S", lambda: self._cur() and self._cur().save())
        self.act_saveas = act(fm, "Save as...", "Ctrl+Shift+S", lambda: self._cur() and self._cur().save(True))
        self.act_upload = act(fm, "Save for upload…", "Ctrl+Alt+S",
                              lambda: self._cur() and self._cur().save_for_upload())
        self.act_print = act(fm, "Print...", "Ctrl+P", self.print_current)
        act(fm, "Two sides on one page (ID card)...", None, self.two_sides)
        fm.addSeparator()
        self.act_close = act(fm, "Close", "Ctrl+W", lambda: self.close_tab(self.tabs.currentIndex()))
        fm.addSeparator()
        act(fm, "Exit", "Ctrl+Q", self.close)
        act(fm, "Quit (force)", "Ctrl+Shift+Q", lambda: os._exit(0))
        em = mb.addMenu("Edit")
        self.act_undo = act(em, "Undo", "Ctrl+Z", lambda: self._cur() and self._cur().undo())
        act(em, "Find a tool", "Ctrl+F", self._find_tool)
        vm = mb.addMenu("View")
        act(vm, "Page list", "Ctrl+L", self.toggle_pages)
        act(vm, "Zoom in", "Ctrl++", lambda: self._cur() and self._cur().zoom_by(1.2))
        act(vm, "Zoom out", "Ctrl+-", lambda: self._cur() and self._cur().zoom_by(1 / 1.2))
        act(vm, "Fit width", "Ctrl+0", lambda: self._cur() and self._cur().fit_width())
        act(vm, "Fit page", "Ctrl+9", lambda: self._cur() and self._cur().fit_page())

    def _find_tool(self):
        self.tools.back()
        self.tools.search.setFocus()
        self.tools.search.selectAll()

    def _fill_recent(self):
        self.recent_menu.clear()
        items = recent_pdfs()
        for p in items[:12]:
            a = QAction(os.path.basename(p), self)
            a.setToolTip(p)
            a.triggered.connect(lambda checked=False, q=p: self.open_path(q))
            self.recent_menu.addAction(a)
        if not items:
            self.recent_menu.addAction("(none)").setEnabled(False)

    def export_menu_popup(self):
        d = self._cur()
        m = QMenu(self)
        m.addAction("Save as …", lambda: d and d.save(True)).setEnabled(d is not None)
        m.addAction("Save for upload …", lambda: d and d.save_for_upload()).setEnabled(d is not None)
        m.addSeparator()
        for key in ("pdf2img", "pdf2word", "pdf2text", "extract"):
            t = ops.TOOL[key]
            a = m.addAction(ui.line_icon(t.symbol, 18, ui.ICON_INK), t.name + " …",
                            lambda k=key: self.tools.open_tool(k))
            a.setEnabled(d is not None or self.centre.currentIndex() == 2)
        b = self.rail.button("export")
        m.exec(b.mapToGlobal(b.rect().topRight()))

    def settings_menu(self):
        m = QMenu(self)
        m.addAction("About Digonto PDF", self.about)
        m.addAction("Keyboard shortcuts", lambda: ui.show_all_shortcuts(self, APP_NAME))
        m.addSeparator()
        a = m.addAction("Upload-ready PDFs when saving")
        a.setCheckable(True)
        a.setChecked(office_mode())
        a.setToolTip("Every save is made like Save for upload: fonts inside the file, forms fixed, and pages that upload websites would refuse become clean pictures with searchable text")
        a.toggled.connect(lambda on: settings().setValue("office_ready", on))
        m.addSeparator()
        m.addAction("Clear the recent files list", self._clear_recent)
        m.addAction("Clear recent tools and favourites", self._clear_tools)
        b = self.rail.button("settings")
        m.exec(b.mapToGlobal(b.rect().topRight()))

    def _clear_recent(self):
        settings().setValue("recent_pdfs", [])
        self.home.refresh()

    def _clear_tools(self):
        panel.settings().setValue("tools/recent", [])
        panel.settings().setValue("tools/fav", [])
        self.tools.fill()

    # ---- what is shown
    def _cur(self):
        if self.centre.currentIndex() != 1:
            return None
        w = self.tabs.currentWidget()
        return w if isinstance(w, DocTab) else None

    def _show_centre(self):
        if self.centre.currentIndex() == 2 and self.batch.files:
            pass
        elif self.tabs.count():
            self.centre.setCurrentIndex(1)
        else:
            self.centre.setCurrentIndex(0)
            self.home.refresh()
        self._tab_changed()

    def _tab_changed(self, *_):
        d = self._cur()
        batch = len(self.batch.files) if self.centre.currentIndex() == 2 else 0
        for a in (self.act_save, self.act_saveas, self.act_upload, self.act_close, self.act_undo, self.act_print):
            a.setEnabled(d is not None)
        self.rail.set_enabled(["save", "print", "pen", "pages"], d is not None)
        self.rail.set_enabled(["export"], d is not None or batch > 0)
        self.rail.button("pages").setChecked(bool(d and d.pages_box.isVisible()))
        self.tools.set_context(d is not None, batch, len(d.selected_pages()) if d else 0)
        if d is not None and self.tools.current in UI_TOOL and d.tool != UI_TOOL[self.tools.current]:
            self.tool_opened(self.tools.current)
        self.setWindowTitle(f"{d.name} - {APP_NAME}" if d else APP_NAME)

    def _title_changed(self):
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if isinstance(w, DocTab):
                self.tabs.setTabText(i, w.title())
                self.tabs.setTabToolTip(i, w.path or "not saved yet")
        d = self._cur()
        if d:
            self.setWindowTitle(f"{d.name} - {APP_NAME}")
            self.tools.context["selected"] = len(d.selected_pages())

    def toggle_pages(self):
        d = self._cur()
        if d:
            self.rail.button("pages").setChecked(d.toggle_pages())

    def _page_step(self, s):
        d = self._cur()
        if d:
            d.goto(d.cur + s)

    # ---- tools
    def tool_opened(self, key):
        t = ops.TOOL[key]
        d = self._cur()
        if t.kind == "ui" and d is not None:
            d.set_tool(UI_TOOL[key], self.tools.ui_box)
            self.rail.button("pages").setChecked(d.pages_box.isVisible())
        elif key == "info" and d is not None:
            self._info_fields(d)
        elif key == "unlock" and d is not None:
            has = bool(d.password)
            lab = QLabel("This PDF has a password. Press the button, then save." if has
                         else "This PDF has no password.")
            lab.setWordWrap(True)
            self.tools.ui_box.addWidget(lab)
            self.tools.run_btn.setEnabled(has)

    def tool_closed(self, key):
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if isinstance(w, DocTab) and w.tool is not None:
                w.set_tool(None)
            if isinstance(w, DocTab):
                w.close_preview()

    def _info_fields(self, d):
        m = d.doc.metadata or {}
        size = human(os.path.getsize(d.path)) if d.path and os.path.exists(d.path) else "not saved yet"
        facts = QLabel(f"<b>Pages:</b> {d.count()}<br><b>Size:</b> {size}<br><b>Made with:</b> "
                       f"{(m.get('producer') or '-')[:40]}<br><b>Password:</b> {'yes' if d.password else 'no'}")
        facts.setWordWrap(True)
        self.tools.ui_box.addWidget(facts)
        self._info_edits = {}
        for key, label in (("title", "Title"), ("author", "Author"), ("subject", "Subject"), ("keywords", "Keywords")):
            lab = QLabel(label)
            lab.setObjectName("dgField")
            e = QLineEdit(m.get(key) or "")
            self.tools.ui_box.addWidget(lab)
            self.tools.ui_box.addWidget(e)
            self._info_edits[key] = e

    def preview_tool(self, key, opts):
        d = self._cur()
        if d is None:
            if self.centre.currentIndex() == 2:
                self.tools.show_result("Open one of the files to see a preview.", True)
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            d.apply_pending()
            before, after, note = panel.preview_images(key, d.doc, d.cur, opts)
            d.show_preview(before, after, note)
        except Exception as e:                                  # noqa: BLE001
            self.tools.show_result(str(e), True)
        finally:
            QApplication.restoreOverrideCursor()

    def run_tool(self, key, opts):
        t = ops.TOOL[key]
        if self.worker is not None:
            return
        if t.kind in ("action", "ui"):
            return self.run_action(key, opts)
        batch = self.centre.currentIndex() == 2 and self.batch.files
        if batch:
            jobs = []
            for f in self.batch.files:
                try:
                    with open(f["path"], "rb") as fh:
                        data = fh.read()
                except OSError as e:
                    f["err"] = str(e)
                    continue
                jobs.append(dict(data=data, base=os.path.splitext(f["name"])[0], password=None))
            self.batch.reset_results()
            self.job_ctx = ("batch", t, opts)
        else:
            d = self._cur()
            if d is None:
                return
            d.apply_pending()
            opts = dict(opts, _selected=d.selected_pages())
            jobs = [dict(data=d.doc.tobytes(), base=os.path.splitext(d.name)[0])]
            self.job_ctx = ("doc", t, opts, d)
        self.worker = panel.JobWorker(t, opts, jobs)
        self.worker.progress.connect(self._job_progress)
        self.worker.jobDone.connect(self._job_done)
        self.worker.allDone.connect(self._all_done)
        self.tools.busy(True)
        if batch:
            self.batch.busy(True)
        self.worker.start()

    def _job_progress(self, k, frac, text):
        if self.job_ctx[0] == "batch":
            n = len(self.batch.files)
            self.batch.progress(k, frac)
            self.tools.show_progress((k + frac) / max(n, 1), f"File {k + 1} of {n}  {text}")
        else:
            self.tools.show_progress(frac, text)

    def _job_done(self, k, result, msg, err):
        if self.job_ctx[0] == "batch":
            self.batch.job_done(k, result, msg, err)
            return
        _, t, opts, d = self.job_ctx
        if err:
            self.tools.show_result(err, True)
            return
        if t.kind == "files":
            self._save_files(result, d)
            return
        if result is None:
            self.tools.show_result(msg)
            return
        d.push_undo()
        d.doc = fitz.open(stream=result, filetype="pdf")
        d.rebuild(keep=True)
        d.close_preview()
        extra = ""
        if t.key == "compress" and d.path and os.path.exists(d.path):
            extra = "  - press Save to keep it"
        self.tools.show_result(msg + extra)

    def _all_done(self, stopped):
        kind = self.job_ctx[0] if self.job_ctx else None
        self.worker.wait()
        self.worker = None
        self.tools.busy(False)
        if kind == "batch":
            self.batch.busy(False)
            ok = sum(1 for f in self.batch.files if f["result"])
            bad = sum(1 for f in self.batch.files if f["err"])
            text = f"Done: {ok} file(s) ready" + (f", {bad} with a problem" if bad else "") + " - press Save all."
            self.tools.show_result("Stopped. " + text if stopped else text, bad > 0 and ok == 0)
        elif stopped:
            self.tools.show_result("Stopped - nothing was changed.")
        self.job_ctx = None

    def stop_job(self):
        if self.worker is not None:
            self.worker.stop()

    def _save_files(self, files, d):
        if not files:
            self.tools.show_result("Nothing to save.")
            return
        if len(files) == 1:
            name, data = files[0]
            ext = os.path.splitext(name)[1]
            p, _ = QFileDialog.getSaveFileName(self, "Save as", os.path.join(last_dir(), name), f"*{ext}")
            if not p:
                return
            if not p.lower().endswith(ext.lower()):
                p += ext
            with open(p, "wb") as f:
                f.write(office_bytes(data))
            remember_dir(p)
            self.tools.show_result(f"Saved: {os.path.basename(p)}")
            self._offer_open(p)
            return
        folder = QFileDialog.getExistingDirectory(self, f"Choose a folder for the {len(files)} files", last_dir())
        if not folder:
            return
        remember_dir(folder)
        for name, data in files:
            with open(unique_path(os.path.join(folder, name)), "wb") as f:
                f.write(data)
        self.tools.show_result(f"Saved {len(files)} files in {os.path.basename(folder) or folder}")
        ui.open_external(folder)

    def _offer_open(self, path):
        if path.lower().endswith(".pdf"):
            self.open_path(path)
        else:
            ui.open_external(os.path.dirname(path))

    # ---- tools that need a dialog
    def run_action(self, key, opts):
        d = self._cur()
        if key == "merge":
            return self.merge(d)
        if key == "img2pdf":
            return self.create_dialog()
        if key == "redact":                                      # the "black out words" button
            batch = self.centre.currentIndex() == 2
            if batch or d is not None:
                return self.run_tool_doc("redact", opts)
            return
        if key == "compare" and d is not None:
            return self.compare(d)
        if key == "info" and d is not None:
            d.push_undo()
            ops.op_info(d.doc, {k: e.text() for k, e in self._info_edits.items()})
            return self.tools.show_result("Saved in the document - press Save to keep it.")
        if key == "lock":
            pw, pw2 = opts.get("pw", ""), opts.get("pw2", "")
            if not pw:
                return self.tools.show_result("Type a password.", True)
            if pw != pw2:
                return self.tools.show_result("The two passwords are not the same.", True)
            if self.centre.currentIndex() == 2:
                return self._batch_lock(pw)
            if d is not None:
                d.password = pw
                d.set_modified(True)
                return self.tools.show_result("The password is set when you save. Don't forget it - "
                                              "nobody can open the PDF without it.")
        if key == "unlock":
            if self.centre.currentIndex() == 2:
                return self._batch_lock(None)
            if d is not None and d.password:
                d.password = None
                d.set_modified(True)
                return self.tools.show_result("Password removed - press Save to keep it.")

    def run_tool_doc(self, key, opts):
        """run a 'ui' tool's button through the normal background worker (redact words)"""
        t = ops.TOOL[key]
        fake = ops.Tool(t.key, t.name, t.group, t.symbol, t.tip, "doc", t.fields, t.run, t.button)
        saved = ops.TOOL[key]
        ops.TOOL[key] = fake
        try:
            self.run_tool(key, opts)
        finally:
            ops.TOOL[key] = saved

    def _batch_lock(self, pw):
        for f in self.batch.files:
            try:
                doc = fitz.open(f["path"])
                if doc.needs_pass:
                    f.update(result=None, err="has a password - open it on its own")
                    continue
                kw = dict(encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=pw, owner_pw=pw) if pw else {}
                f.update(result=doc.tobytes(garbage=3, deflate=True, **kw), msg="", err="")
            except Exception as e:                                # noqa: BLE001
                f.update(result=None, err=str(e))
        self.batch._rebuild()
        self.tools.show_result("Ready - press Save all.")

    def merge(self, d):
        if self.centre.currentIndex() == 2 and self.batch.files:
            paths = [f["path"] for f in self.batch.files]
        else:
            paths, _ = QFileDialog.getOpenFileNames(self, "Choose the PDFs to put together (in this order)",
                                                    last_dir(), PDF_FILTER)
            if not paths:
                return
        out = fitz.open()
        if d is not None and self.centre.currentIndex() == 1:
            d.apply_pending()
            out.insert_pdf(d.doc)
        bad = []
        for p in paths:
            src = open_pdf(self, p)
            if src is None:
                bad.append(os.path.basename(p))
                continue
            out.insert_pdf(src)
        if out.page_count == 0:
            return
        self.add_doc(out, None, "Merged.pdf", True)
        self.tools.show_result(f"Merged into a new PDF with {out.page_count} pages - press Save to keep it."
                               + (f"  Skipped: {', '.join(bad)}" if bad else ""))

    def compare(self, d):
        p, _ = QFileDialog.getOpenFileName(self, "Compare with which PDF?", last_dir(), PDF_FILTER)
        if not p:
            return
        other = open_pdf(self, p)
        if other is None:
            return
        n, html = ops.compare_html(d.doc, other, d.name, os.path.basename(p))
        dlg = QDialog(self)
        dlg.setWindowTitle(f"Compare - {d.name} / {os.path.basename(p)}")
        dlg.resize(980, 700)
        v = QVBoxLayout(dlg)
        v.addWidget(QLabel(f"<b>{n} page(s) with changes.</b>  Red: only in {d.name}   Green: only in "
                           f"{os.path.basename(p)}" if n else ""))
        from PySide6.QtWidgets import QTextBrowser
        tb = QTextBrowser()
        tb.setHtml(html)
        v.addWidget(tb, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Close)
        bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        dlg.exec()

    # ---- batch (many files)
    def add_folder(self, folder=None):
        if not folder:
            folder = QFileDialog.getExistingDirectory(self, "Choose a folder with PDF files", last_dir())
            if not folder:
                return
        remember_dir(os.path.join(folder, "x"))
        paths = sorted(glob.glob(os.path.join(folder, "*.pdf")) + glob.glob(os.path.join(folder, "*.PDF")))
        paths = sorted(set(paths), key=lambda p: os.path.basename(p).lower())
        if not paths:
            QMessageBox.information(self, APP_NAME, "There are no PDF files in this folder.")
            return
        self.open_batch(paths, os.path.basename(folder.rstrip("/\\")) or folder)

    def open_batch(self, paths, name=""):
        self.batch.set_files(paths, name)
        self.centre.setCurrentIndex(2)
        self._tab_changed()

    def batch_add_more(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Add PDFs to the list", last_dir(), PDF_FILTER)
        if files:
            self.batch.set_files(files)
            self._tab_changed()

    def close_batch(self):
        if self.worker is not None:
            return
        if any(f["result"] for f in self.batch.files):
            r = QMessageBox.question(self, APP_NAME, "Close the list without saving the results?")
            if r != QMessageBox.Yes:
                return
        self.batch.files = []
        self.batch._rebuild()
        self.centre.setCurrentIndex(1 if self.tabs.count() else 0)
        self._show_centre()

    def batch_save_all(self):
        ready = [f for f in self.batch.files if f["result"]]
        if not ready:
            return
        folder = QFileDialog.getExistingDirectory(self, "Save the results in which folder?", last_dir())
        if not folder:
            return
        n = 0
        for f in ready:
            if isinstance(f["result"], list):
                for name, data in f["result"]:
                    with open(unique_path(os.path.join(folder, name)), "wb") as fh:
                        fh.write(office_bytes(data))
                    n += 1
            else:
                target = os.path.join(folder, f["name"])
                if os.path.abspath(target) == os.path.abspath(f["path"]):
                    target = unique_path(os.path.join(folder, os.path.splitext(f["name"])[0] + " - Digonto.pdf"))
                else:
                    target = unique_path(target)
                with open(target, "wb") as fh:
                    fh.write(office_bytes(f["result"]))
                n += 1
        remember_dir(os.path.join(folder, "x"))
        self.tools.show_result(f"Saved {n} file(s) in {os.path.basename(folder) or folder}.")
        ui.open_external(folder)

    # ---- documents
    def add_doc(self, doc, path=None, name=None, modified=False):
        d = DocTab(doc, path, name, modified)
        d.password = getattr(doc, "_dg_pw", None)
        d.changed.connect(self._title_changed)
        d.openBytes.connect(lambda data, nm: self.add_doc(fitz.open(stream=data, filetype="pdf"), None, nm, True))
        d.thumbs.itemSelectionChanged.connect(self._title_changed)
        i = self.tabs.addTab(d, d.title())
        self.tabs.setTabToolTip(i, path or "not saved yet")
        self.centre.setCurrentIndex(1)
        self.tabs.setCurrentIndex(i)
        self._tab_changed()
        return d

    def open_path(self, path):
        path = os.path.abspath(path)
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if isinstance(w, DocTab) and w.path and os.path.abspath(w.path) == path:
                self.centre.setCurrentIndex(1)
                self.tabs.setCurrentIndex(i)
                return
        doc = open_pdf(self, path)
        if not doc:
            return
        remember_dir(path)
        add_recent_pdf(path)
        ui.add_recent("digonto-pdf", path)
        self.add_doc(doc, path)
        self.home.refresh()

    def open_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Open PDF", last_dir(), PDF_FILTER)
        for f in files:
            self.open_path(f)

    def create_dialog(self, files=()):
        dlg = CreateDialog(self, files)
        if dlg.exec() == QDialog.Accepted:
            for name, data in dlg.results:
                self.add_doc(fitz.open(stream=data, filetype="pdf"), None, name, True)

    def two_sides(self):
        import digonto_sides
        d = digonto_sides.SidesDialog(self, "pdf")
        if d.exec() != QDialog.Accepted or not d.items:
            return
        try:
            self.add_doc(fitz.open(stream=d.pdf_bytes(), filetype="pdf"), None, "Two sides.pdf", True)
        except Exception as e:  # noqa: BLE001
            QMessageBox.warning(self, "Two sides on one page", f"Could not make the page.\n{e}")

    def print_current(self):
        w = self._cur()
        if w is None:
            return
        w.apply_pending()
        doc = w.doc
        zoom = 300 / 72.0

        def page(i, p, rect):
            pix = doc[i].get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
            p.drawImage(ui.fit_rect(img.width(), img.height(), rect), img)
        ui.print_pages(self, doc.page_count, page, w.name)

    def close_tab(self, i):
        w = self.tabs.widget(i)
        if not isinstance(w, DocTab):
            return
        if self.worker is not None and self.job_ctx and self.job_ctx[0] == "doc" and self.job_ctx[3] is w:
            return
        if not w.maybe_close():
            return
        self.tabs.removeTab(i)
        w.deleteLater()
        self._show_centre()

    # ---- window
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            if self.centre.currentIndex() == 0:
                self.home.set_dropping(True)

    def dragLeaveEvent(self, e):
        self.home.set_dropping(False)

    def dropEvent(self, e):
        self.home.set_dropping(False)
        files = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        folders = [f for f in files if os.path.isdir(f)]
        pdfs = [f for f in files if f.lower().endswith(".pdf") and os.path.isfile(f)]
        others = [f for f in files if os.path.isfile(f) and not f.lower().endswith(".pdf") and conv.kind_of(f)]
        for f in folders:
            self.add_folder(f)
        if self.centre.currentIndex() == 2 and pdfs:
            self.batch.set_files(pdfs)
            self._tab_changed()
        else:
            for f in pdfs:
                self.open_path(f)
        if others:
            self.create_dialog(others)

    def about(self):
        ui.about_dialog(
            self, APP_NAME,
            f"<p><b>{APP_NAME} {APP_VERSION}</b></p>"
            "<p>Open, organise, clean up, OCR, convert, sign and protect PDF files on Ubuntu, macOS and Windows.</p>"
            f"<p>Made with Qt for Python (PySide6) and PyMuPDF {fitz.VersionBind}. Text recognition (OCR) uses the "
            "Tesseract engine inside PyMuPDF with English, German, Italian and Bengali language files.</p>"
            "<p>Digonto PDF never connects to the internet. Your files stay on your computer.</p>",
            make_app_icon())

    def closeEvent(self, e):
        if self.worker is not None:
            self.worker.stop()
            self.worker.wait(3000)
        for i in range(self.tabs.count()):
            w = self.tabs.widget(i)
            if isinstance(w, DocTab):
                self.centre.setCurrentIndex(1)
                self.tabs.setCurrentIndex(i)
                if not w.maybe_close():
                    e.ignore()
                    return
        settings().setValue("geometry", self.saveGeometry())
        super().closeEvent(e)
        ui.force_exit_soon()


def main():
    if sys.platform == "win32":
        try:   # own taskbar icon instead of the Python icon
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Digonto.PDF")
        except Exception:
            pass
    if hasattr(Qt, "HighDpiScaleFactorRoundingPolicy"):
        QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    ui.mac_app_name(APP_NAME)
    app = QApplication(sys.argv)
    if ui.forward_to_running("digonto-pdf", sys.argv[1:]):
        os._exit(0)
    app.setApplicationName("Digonto PDF")
    app.setApplicationDisplayName("Digonto PDF")
    app.setOrganizationName("DigontoPDF")
    app.setDesktopFileName("digonto-pdf")
    ui.migrate_legacy()
    ui.apply_theme(app)
    ui.install_dialogs()
    ui.install_about()
    app.setWindowIcon(make_app_icon())
    def _bye():                       # closing the last window ends the program for good
        import threading
        t = threading.Timer(1.2, lambda: os._exit(0)); t.daemon = True; t.start()
    app.lastWindowClosed.connect(_bye)
    app.aboutToQuit.connect(_bye)
    from PySide6.QtCore import QEvent, QObject

    class _Opener(QObject):          # macOS hands files to an app through an event (double-click, "Open with", dock drop)
        win = None
        pending = []

        def eventFilter(self, o, e):
            if e.type() == QEvent.FileOpen:
                if self.win is None:
                    self.pending.append(e.file())
                else:
                    self.win.open_path(e.file())
                    self.win.raise_()
                    self.win.activateWindow()
                return True
            return False
    opener = _Opener()
    app.installEventFilter(opener)
    w = MainWindow(sys.argv[1:])
    opener.win = w
    w.show()
    ui.serve_single("digonto-pdf", w, w.open_path)
    for pth in opener.pending:
        w.open_path(pth)
    os._exit(app.exec())


if __name__ == "__main__":
    main()

