"""Two sides on one page (ID card, driving licence, bank card ...).

Open the scan of the FRONT and of the BACK (picture or PDF), the card is found, straightened and cut out,
and both sides are put on one page - at the size of the scan, or at the standard card size.
Shared by Digonto PDF and Digonto Images. Uses only Qt (and PyMuPDF, only to read a PDF scan).
"""
import math
import os

from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QPointF, QRectF, QSize, Qt
from PySide6.QtGui import QColor, QImage, QPainter, QPixmap, QTransform
from PySide6.QtWidgets import (QCheckBox, QComboBox, QDialog, QDialogButtonBox, QDoubleSpinBox, QFileDialog, QGridLayout,
                               QHBoxLayout, QLabel, QMessageBox, QPushButton, QVBoxLayout, QWidget)

PAGES = {"A4 (210 x 297 mm)": (210.0, 297.0), "Letter (216 x 279 mm)": (215.9, 279.4), "A5 (148 x 210 mm)": (148.0, 210.0)}
CARD_MM = (85.6, 54.0)                       # ID-1: ID card, driving licence, bank card


# --------------------------------------------------------------------------- reading a scan
def load_side(path):
    """-> (QImage, dpi). A PDF is read from its first page at 300 dpi."""
    if path.lower().endswith(".pdf"):
        import pymupdf
        doc = pymupdf.open(path)
        try:
            page = doc[0]
            imgs = page.get_images(full=True)
            if len(imgs) == 1:                      # a scanner PDF: one picture on the page -> take the picture itself, same pixels
                try:
                    info = doc.extract_image(imgs[0][0])
                    qi = QImage.fromData(info["image"])
                    if not qi.isNull():
                        qi = qi.convertToFormat(QImage.Format_RGB32)
                        inch_w = page.rect.width / 72.0
                        dpi = qi.width() / inch_w if inch_w else 300.0
                        if not 100 <= dpi <= 1200 or abs(qi.width() / 8.27 - 300) < 60:   # page size in the PDF is not the paper size: assume an A4 scan
                            dpi = qi.width() / 8.27 if abs(qi.width() / 8.27 - 300) < 60 else 300.0
                        return qi, float(dpi)
                except Exception:
                    pass
            pix = page.get_pixmap(matrix=pymupdf.Matrix(300 / 72.0, 300 / 72.0), alpha=False)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
        finally:
            doc.close()
        return img, 300.0
    img = QImage(path)
    if img.isNull():
        raise OSError("this file is not a picture or a PDF")
    dpm = img.dotsPerMeterX()
    dpi = dpm * 0.0254 if dpm > 3800 else 300.0       # no dpi stored in the file (96 dpi is only a default) -> a scanner gives 300
    return img.convertToFormat(QImage.Format_RGB32), float(dpi)


# --------------------------------------------------------------------------- finding the card
def _percentile(sorted_vals, q):
    return sorted_vals[min(len(sorted_vals) - 1, max(0, int(q * (len(sorted_vals) - 1))))]


def find_card(img):
    """-> (angle_degrees, centre_x, centre_y, width, height) in pixels of img, or None when no clear card is seen."""
    sm = img.scaledToWidth(900, Qt.SmoothTransformation) if img.width() > 900 else img.copy()
    sm = sm.convertToFormat(QImage.Format_RGB32)
    w, h = sm.width(), sm.height()
    k = img.width() / float(w)
    data = bytes(sm.constBits())[: sm.sizeInBytes()]
    bpl = sm.bytesPerLine()

    def px(x, y):
        o = y * bpl + x * 4
        return data[o + 2], data[o + 1], data[o]          # r, g, b

    ring = []
    for x in range(0, w, 3):
        ring += [px(x, 1), px(x, h - 2)]
    for y in range(0, h, 3):
        ring += [px(1, y), px(w - 2, y)]
    bg = tuple(sorted(c[i] for c in ring)[len(ring) // 2] for i in range(3))
    devs = sorted(max(abs(c[0] - bg[0]), abs(c[1] - bg[1]), abs(c[2] - bg[2])) for c in ring)
    thr = max(7, devs[int(0.99 * (len(devs) - 1))] + 5)           # a light card on white paper still shows; scanner noise does not
    on = set()
    for y in range(0, h, 2):
        base = y * bpl
        for x in range(0, w, 2):
            o = base + x * 4
            if max(abs(data[o + 2] - bg[0]), abs(data[o + 1] - bg[1]), abs(data[o] - bg[2])) > thr:
                on.add((x, y))
    # a speck of dust is alone, the card is solid: keep only points whose neighbours are on too
    pts = [(x, y) for (x, y) in on
           if sum(((x + dx, y + dy) in on) for dx in (-4, -2, 0, 2, 4) for dy in (-4, -2, 0, 2, 4)) >= 15]
    if len(pts) < 0.02 * (w * h / 4):
        return None
    if len(pts) > 5000:
        pts = pts[:: max(1, len(pts) // 5000)]
    best = None
    a = -10.0
    while a <= 10.01:
        c, s = math.cos(math.radians(a)), math.sin(math.radians(a))
        xs = sorted(x * c - y * s for x, y in pts)
        ys = sorted(x * s + y * c for x, y in pts)
        x0, x1, y0, y1 = _percentile(xs, 0.002), _percentile(xs, 0.998), _percentile(ys, 0.002), _percentile(ys, 0.998)
        area = (x1 - x0) * (y1 - y0)
        if best is None or area < best[0]:
            best = (area, a, x0, x1, y0, y1)
        a += 0.5
    _, a, x0, x1, y0, y1 = best
    bw, bh = x1 - x0, y1 - y0
    if bw > 0.97 * w and bh > 0.97 * h:                 # the "card" is the whole scan: nothing to cut
        return None
    if bw < 0.12 * w or bh < 0.12 * h:
        return None
    c, s = math.cos(math.radians(a)), math.sin(math.radians(a))
    rx, ry = (x0 + x1) / 2, (y0 + y1) / 2               # centre in the rotated frame
    cx, cy = rx * c + ry * s, -rx * s + ry * c          # back to the picture
    pad = 0.006 * bw * k                                  # a hair of paper around the card, never cut into it
    return a, cx * k, cy * k, bw * k + 2 * pad, bh * k + 2 * pad


def cut_card(img, found, margin_px=0):
    """straighten the picture and cut the card out of it"""
    a, cx, cy, bw, bh = found
    W, H = max(8, int(round(bw)) + 2 * margin_px), max(8, int(round(bh)) + 2 * margin_px)
    out = QImage(W, H, QImage.Format_RGB32)
    out.fill(QColor("white"))
    p = QPainter(out)
    p.setRenderHints(QPainter.SmoothPixmapTransform | QPainter.Antialiasing)
    tr = QTransform()
    tr.translate(W / 2.0, H / 2.0)
    tr.rotate(a)
    tr.translate(-cx, -cy)
    p.setTransform(tr)
    p.drawImage(0, 0, img)
    p.end()
    return out


def rotate90(img, turns=1):
    return img.transformed(QTransform().rotate(90 * (turns % 4)), Qt.SmoothTransformation)


# --------------------------------------------------------------------------- layout and output
def layout(sides, page_mm, size_mode, custom_w_mm, gap_mm, margin_mm=12.0, valign="top", halign="center"):
    """sides = [(QImage, dpi), (QImage, dpi)] -> (items, page_mm); item = (QImage, x_mm, y_mm, w_mm, h_mm)"""
    pw, ph = page_mm
    sized = []
    for img, dpi in sides:
        if size_mode == "card":
            w_mm = CARD_MM[0] if img.width() >= img.height() else CARD_MM[1]
        elif size_mode == "custom":
            w_mm = custom_w_mm
        else:                                            # keep the scan's own size
            w_mm = img.width() / dpi * 25.4
        sized.append([img, w_mm, w_mm * img.height() / img.width()])
    usable_w, usable_h = pw - 2 * margin_mm, ph - 2 * margin_mm
    total_h = sum(s[2] for s in sized) + gap_mm * (len(sized) - 1)
    shrink = min(1.0, usable_w / max(s[1] for s in sized), usable_h / total_h)      # never bigger than the page
    th = total_h * shrink
    y = {"top": margin_mm, "bottom": ph - margin_mm - th}.get(valign, (ph - th) / 2.0)
    items = []
    for img, w_mm, h_mm in sized:
        w2, h2 = w_mm * shrink, h_mm * shrink
        x = {"left": margin_mm, "right": pw - margin_mm - w2}.get(halign, (pw - w2) / 2.0)
        items.append((img, x, y, w2, h2))
        y += h2 + gap_mm * shrink
    return items, page_mm


def _jpeg_bytes(img, quality=95):
    ba = QByteArray()
    buf = QBuffer(ba)
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "JPG", quality)
    return bytes(ba)


def build_pdf_bytes(items, page_mm):
    """one PDF page with the pictures at their size"""
    import pymupdf
    mm = 72.0 / 25.4
    doc = pymupdf.open()
    page = doc.new_page(width=page_mm[0] * mm, height=page_mm[1] * mm)
    for img, x, y, w, h in items:
        page.insert_image(pymupdf.Rect(x * mm, y * mm, (x + w) * mm, (y + h) * mm), stream=_jpeg_bytes(img))
    data = doc.tobytes(garbage=3, deflate=True)
    doc.close()
    return data


def render_page(items, page_mm, width_px):
    """a picture of the page (preview, or a raster page)"""
    scale = width_px / page_mm[0]
    out = QImage(int(round(page_mm[0] * scale)), int(round(page_mm[1] * scale)), QImage.Format_RGB32)
    out.fill(QColor("white"))
    p = QPainter(out)
    p.setRenderHints(QPainter.SmoothPixmapTransform | QPainter.Antialiasing)
    for img, x, y, w, h in items:
        p.drawImage(QRectF(x * scale, y * scale, w * scale, h * scale), img)
    p.end()
    return out


# --------------------------------------------------------------------------- the window
class SidesDialog(QDialog):
    """choose both sides, see the page, press OK"""

    def __init__(self, parent=None, want="pdf"):
        super().__init__(parent)
        self.setWindowTitle("Two sides on one page")
        self.setMinimumSize(820, 560)
        self.want = want
        self.src = [None, None]               # (QImage, dpi) as loaded
        self.names = ["", ""]
        self.turns = [0, 0]
        self.items, self.page_mm = [], PAGES["A4 (210 x 297 mm)"]

        top = QVBoxLayout(self)
        intro = QLabel("Open the <b>front</b> and the <b>back</b> (picture or PDF). The card is found, straightened and cut out, "
                       "then both sides are placed on one page.")
        intro.setWordWrap(True)
        top.addWidget(intro)
        row = QHBoxLayout()
        top.addLayout(row, 1)

        left = QVBoxLayout()
        row.addLayout(left, 0)
        self.lab = []
        for i, name in enumerate(("Front", "Back")):
            box = QGridLayout()
            t = QLabel(f"<b>{name}</b>")
            box.addWidget(t, 0, 0)
            b = QPushButton("Choose file ...")
            b.clicked.connect(lambda _=False, k=i: self.choose(k))
            box.addWidget(b, 0, 1)
            r = QPushButton("Turn 90°")
            r.clicked.connect(lambda _=False, k=i: self.turn(k))
            box.addWidget(r, 0, 2)
            n = QLabel("no file yet")
            n.setWordWrap(True)
            n.setMaximumWidth(300)
            box.addWidget(n, 1, 0, 1, 3)
            self.lab.append(n)
            left.addLayout(box)
        self.auto = QCheckBox("Find the card, straighten and cut it out")
        self.auto.setChecked(True)
        self.auto.toggled.connect(self.refresh)
        left.addWidget(self.auto)
        self.land = QCheckBox("Turn cards to lie flat (landscape)")
        self.land.setChecked(True)
        self.land.toggled.connect(self.refresh)
        left.addWidget(self.land)

        form = QGridLayout()
        left.addLayout(form)
        form.addWidget(QLabel("Page"), 0, 0)
        self.page = QComboBox()
        self.page.addItems(list(PAGES))
        self.page.currentIndexChanged.connect(self.refresh)
        form.addWidget(self.page, 0, 1)
        form.addWidget(QLabel("Size"), 1, 0)
        self.size = QComboBox()
        self.size.addItem("Keep the size of the scan", "scan")
        self.size.addItem("Standard card (85.6 x 54 mm)", "card")
        self.size.addItem("My own width ...", "custom")
        self.size.currentIndexChanged.connect(self.refresh)
        form.addWidget(self.size, 1, 1)
        form.addWidget(QLabel("Width (mm)"), 2, 0)
        self.cw = QDoubleSpinBox()
        self.cw.setRange(20, 400)
        self.cw.setValue(85.6)
        self.cw.setDecimals(1)
        self.cw.valueChanged.connect(self.refresh)
        form.addWidget(self.cw, 2, 1)
        form.addWidget(QLabel("Space between (mm)"), 3, 0)
        self.gap = QDoubleSpinBox()
        self.gap.setRange(0, 100)
        self.gap.setValue(8)
        self.gap.valueChanged.connect(self.refresh)
        form.addWidget(self.gap, 3, 1)
        form.addWidget(QLabel("Up / down"), 4, 0)
        self.va = QComboBox()
        for t, v in (("Top", "top"), ("Middle", "middle"), ("Bottom", "bottom")):
            self.va.addItem(t, v)
        self.va.currentIndexChanged.connect(self.refresh)
        form.addWidget(self.va, 4, 1)
        form.addWidget(QLabel("Left / right"), 5, 0)
        self.ha = QComboBox()
        for t, v in (("Centre", "center"), ("Left", "left"), ("Right", "right")):
            self.ha.addItem(t, v)
        self.ha.currentIndexChanged.connect(self.refresh)
        form.addWidget(self.ha, 5, 1)
        form.addWidget(QLabel("Distance from edge (mm)"), 6, 0)
        self.margin = QDoubleSpinBox()
        self.margin.setRange(0, 100)
        self.margin.setValue(15)
        self.margin.valueChanged.connect(self.refresh)
        form.addWidget(self.margin, 6, 1)
        self.info = QLabel("")
        self.info.setWordWrap(True)
        self.info.setMaximumWidth(300)
        left.addWidget(self.info)
        left.addStretch(1)

        self.view = QLabel("Choose both sides to see the page")
        self.view.setAlignment(Qt.AlignCenter)
        self.view.setMinimumSize(380, 480)
        self.view.setStyleSheet("background:#8c8697; border-radius:6px; color:white;")
        row.addWidget(self.view, 1)

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText("Make the page")
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        top.addWidget(bb)
        self.ok = bb.button(QDialogButtonBox.Ok)
        self.ok.setEnabled(False)

    # ---- actions
    def choose(self, k):
        p, _ = QFileDialog.getOpenFileName(self, ("Front", "Back")[k] + " - choose a picture or PDF", "",
                                           "Pictures and PDF (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.pdf)")
        if not p:
            return
        try:
            self.src[k] = load_side(p)
        except Exception as e:                        # noqa: BLE001
            QMessageBox.warning(self, "Two sides on one page", f"Could not open this file.\n{e}")
            return
        self.names[k] = os.path.basename(p)
        self.turns[k] = 0
        self.refresh()

    def turn(self, k):
        self.turns[k] = (self.turns[k] + 1) % 4
        self.refresh()

    def prepared(self, k):
        img, dpi = self.src[k]
        notes = []
        if self.auto.isChecked():
            found = find_card(img)
            if found:
                img = cut_card(img, found)
                notes.append(f"card found, turned {found[0]:+.1f}°")
            else:
                notes.append("no clear card edge - whole scan used")
            if self.land.isChecked() and img.height() > img.width():
                img = rotate90(img, 1)
        if self.turns[k]:
            img = rotate90(img, self.turns[k])
        return img, dpi, notes

    def refresh(self):
        self.cw.setEnabled(self.size.currentData() == "custom")
        for k in (0, 1):
            self.lab[k].setText(self.names[k] or "no file yet")
        if not (self.src[0] and self.src[1]):
            self.ok.setEnabled(False)
            return
        sides, notes = [], []
        for k in (0, 1):
            img, dpi, n = self.prepared(k)
            sides.append((img, dpi))
            notes.append(("Front: ", "Back: ")[k] + "; ".join(n))
        self.page_mm = PAGES[self.page.currentText()]
        self.items, self.page_mm = layout(sides, self.page_mm, self.size.currentData(), self.cw.value(), self.gap.value(),
                                          self.margin.value(), self.va.currentData(), self.ha.currentData())
        shown = ", ".join(f"{it[3]:.0f} x {it[4]:.0f} mm" for it in self.items)
        self.info.setText("\n".join(notes) + f"\nOn the page: {shown}")
        prev = render_page(self.items, self.page_mm, 380)
        pm = QPixmap.fromImage(prev)
        self.view.setPixmap(pm.scaled(self.view.size() - QSize(20, 20), Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.ok.setEnabled(True)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if self.items:
            self.refresh()

    # ---- result
    def pdf_bytes(self):
        return build_pdf_bytes(self.items, self.page_mm)

    def page_image(self, dpi=300):
        return render_page(self.items, self.page_mm, int(round(self.page_mm[0] / 25.4 * dpi)))
