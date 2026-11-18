"""Digonto PDF window parts: the tools panel on the right, the batch list, the before/after slider and the
background worker that runs the tools (so the window never freezes)."""
import os
import time

try:
    import pymupdf as fitz
except ImportError:
    import fitz

from PySide6.QtCore import QPointF, QRectF, QSettings, QSize, Qt, QThread, Signal
from PySide6.QtGui import QColor, QFont, QImage, QPainter, QPen
from PySide6.QtWidgets import (QCheckBox, QComboBox, QFrame, QHBoxLayout, QLabel, QLineEdit, QListWidget,
                               QListWidgetItem, QMenu, QProgressBar, QPushButton, QScrollArea, QSlider, QSpinBox,
                               QStackedWidget, QToolButton, QVBoxLayout, QWidget)

import digonto_ui as ui
import pdf_ops as ops
import pdf_scan as scan

INK, SOFT, LINE, PINK = ui.INK, ui.SOFT, ui.LINE, ui.PINK

PANEL_QSS = f"""
QFrame#dgTools {{ background: white; border-left: 1px solid {LINE}; }}
QFrame#dgTools QLineEdit {{ background: #F7F4FB; border: 1px solid {LINE}; border-radius: 9px;
                            padding: 7px 10px 7px 30px; font-size: 13px; }}
QFrame#dgTools QListWidget {{ border: none; background: transparent; font-size: 13px; outline: none; }}
QFrame#dgTools QListWidget::item {{ padding: 6px 6px; border-radius: 8px; color: {INK}; }}
QFrame#dgTools QListWidget::item:selected {{ background: #FCE4EA; color: {INK}; }}
QFrame#dgTools QListWidget::item:hover {{ background: #F6F0FB; }}
QFrame#dgTools QLabel#dgTitle {{ font-size: 15px; font-weight: 700; color: {INK}; }}
QFrame#dgTools QLabel#dgTip {{ color: {SOFT}; font-size: 12px; }}
QFrame#dgTools QLabel#dgField {{ color: {SOFT}; font-size: 11.5px; font-weight: 600; margin-top: 4px; }}
QFrame#dgTools QLabel#dgResult {{ color: #2E7D52; font-size: 12.5px; }}
QFrame#dgTools QComboBox, QFrame#dgTools QSpinBox {{ background: white; border: 1px solid {LINE};
                                                    border-radius: 8px; padding: 5px 8px; }}
QFrame#dgTools QToolButton#dgBack {{ border: none; border-radius: 8px; padding: 4px; }}
QFrame#dgTools QToolButton#dgBack:hover {{ background: #F1EAF8; }}
QFrame#dgTools QProgressBar {{ border: none; background: #EEE8F5; border-radius: 4px; max-height: 8px; }}
QFrame#dgTools QProgressBar::chunk {{ border-radius: 4px; background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                                      stop:0 {PINK}, stop:1 {ui.AMBER}); }}
QFrame#dgTools QPushButton {{ background: white; border: 1px solid #D9D0E6; border-radius: 9px; padding: 7px 12px;
                              color: {INK}; font-weight: 600; }}
QFrame#dgTools QPushButton:hover {{ background: #F7F2FB; }}
QFrame#dgTools QPushButton:disabled {{ color: #A79CB8; }}
QFrame#dgTools QPushButton#dgPrimary {{ border: none; padding: 9px 12px; font-size: 13.5px;
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 {PINK}, stop:1 {ui.AMBER}); }}
QFrame#dgTools QPushButton#dgPrimary:hover {{ background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
    stop:0 #FF5F82, stop:1 #FFA444); }}
QFrame#dgTools QPushButton#dgPrimary:disabled {{ background: #E9E3F0; color: #A79CB8; }}
QFrame#dgTools QWidget#dgWhite, QFrame#dgTools QScrollArea {{ background: white; }}
""" + ui.SHELL_QSS

WHICH_LABELS = {"all": "All pages", "selected": "Selected pages", "ranges": "Page numbers …"}


def settings():
    return QSettings("DigontoPDF", "DigontoPDF")


# ------------------------------------------------------------------------------------------------- worker --
class JobWorker(QThread):
    """runs one tool on one or many PDFs.  jobs: list of dicts with data, base, password"""
    progress = Signal(int, float, str)               # job, 0..1, text
    jobDone = Signal(int, object, str, str)          # job, result (bytes | [(name, bytes)] | None), message, error
    allDone = Signal(bool)                           # stopped?

    def __init__(self, tool, opts, jobs):
        super().__init__()
        self.tool, self.opts, self.jobs = tool, opts, jobs
        self._stop = False

    def stop(self):
        self._stop = True

    def run(self):
        for k, job in enumerate(self.jobs):
            if self._stop:
                break
            prog = (lambda f, t="", k=k: self.progress.emit(k, f, t))
            try:
                if self.tool.kind == "files":
                    res = ops.run_files_tool(self.tool, job["data"], self.opts, job["base"], prog,
                                             lambda: self._stop, job.get("password"))
                    msg = f"{len(res)} file(s)"
                else:
                    msg, res = ops.run_doc_tool(self.tool, job["data"], self.opts, prog, lambda: self._stop,
                                                job.get("password"))
                self.jobDone.emit(k, res, msg, "")
            except ops.Stopped:
                break
            except Exception as e:                                   # noqa: BLE001 - show it, go on with the next
                self.jobDone.emit(k, None, "", str(e) or type(e).__name__)
        self.allDone.emit(self._stop)


# ---------------------------------------------------------------------------------------------- options --
class OptionsForm(QWidget):
    """builds the input fields of a tool from its field list and reads them back"""
    changed = Signal()

    def __init__(self, fields, selected_count=0):
        super().__init__()
        self.inputs = {}
        self._labels = {}
        v = QVBoxLayout(self)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(4)
        store = settings()
        for f in fields:
            kind, key = f[0], f[1]
            saved = store.value(f"opt/{key}")
            if kind == "check":
                w = QCheckBox(f[2])
                w.setChecked(f[3] if saved is None else saved in (True, "true", "1"))
                w.toggled.connect(self.changed.emit)
            elif kind == "langs":
                v.addWidget(self._label(f[2]))
                box = QWidget()
                bl = QVBoxLayout(box)
                bl.setContentsMargins(0, 0, 0, 0)
                bl.setSpacing(2)
                chosen = saved.split(",") if isinstance(saved, str) and saved else list(f[4])
                checks = []
                for name, code in scan.LANGUAGES:
                    c = QCheckBox(name)
                    c.setChecked(code in chosen)
                    c.setProperty("code", code)
                    bl.addWidget(c)
                    checks.append(c)
                self.inputs[key] = ("langs", checks)
                v.addWidget(box)
                continue
            else:
                v.addWidget(self._label(f[2]))
                if kind == "choice":
                    w = QComboBox()
                    w.setSizeAdjustPolicy(QComboBox.AdjustToMinimumContentsLengthWithIcon)
                    w.setMinimumContentsLength(12)
                    for o in f[3]:
                        label = WHICH_LABELS.get(o, o) if key == "which" else o
                        if o == "selected" and key == "which" and selected_count:
                            label = f"Selected pages ({selected_count})"
                        w.addItem(label, o)
                    want = saved if saved in f[3] and key != "which" else f[4]
                    w.setCurrentIndex(max(0, f[3].index(want)))
                    w.currentIndexChanged.connect(self.changed.emit)
                elif kind in ("text", "password"):
                    w = QLineEdit()
                    if kind == "password":
                        w.setEchoMode(QLineEdit.Password)
                    else:
                        w.setText(saved if isinstance(saved, str) and key != "words" else f[3])
                    w.textChanged.connect(self.changed.emit)
                elif kind == "int":
                    w = QSpinBox()
                    w.setRange(f[3], f[4])
                    try:
                        w.setValue(int(saved) if saved is not None else f[5])
                    except (TypeError, ValueError):
                        w.setValue(f[5])
                    w.valueChanged.connect(self.changed.emit)
                elif kind == "slider":
                    w = QSlider(Qt.Horizontal)
                    w.setRange(f[3], f[4])
                    try:
                        w.setValue(int(saved) if saved is not None else f[5])
                    except (TypeError, ValueError):
                        w.setValue(f[5])
                    w.sliderReleased.connect(self.changed.emit)
                else:
                    continue
            self.inputs[key] = (kind, w)
            v.addWidget(w)
        _attach_labels(self)
        if "which" in self.inputs and "ranges" in self.inputs:
            combo = self.inputs["which"][1]
            rng = self.inputs["ranges"][1]
            lab = self._labels.get("ranges")

            def sync():
                on = combo.currentData() == "ranges"
                rng.setVisible(on)
                if lab:
                    lab.setVisible(on)
            combo.currentIndexChanged.connect(sync)
            sync()
        if "mode" in self.inputs:                                # split / crop: show only the field that matters
            combo = self.inputs["mode"][1]

            def sync2():
                m = combo.currentData()
                for key, show in (("every", m == "Every N pages"), ("margin", m == "Same margin all round")):
                    if key in self.inputs:
                        self.inputs[key][1].setVisible(show)
                        if key in self._labels:
                            self._labels[key].setVisible(show)
                if "ranges" in self.inputs and "which" not in self.inputs:
                    self.inputs["ranges"][1].setVisible(m == "By ranges")
                    if "ranges" in self._labels:
                        self._labels["ranges"].setVisible(m == "By ranges")
            combo.currentIndexChanged.connect(sync2)
            sync2()

    def _label(self, text):
        lab = QLabel(text)
        lab.setObjectName("dgField")
        lab.setWordWrap(True)
        return lab

    def values(self, remember=True):
        out = {}
        store = settings()
        for key, (kind, w) in self.inputs.items():
            if kind == "check":
                out[key] = w.isChecked()
            elif kind == "langs":
                out[key] = [c.property("code") for c in w if c.isChecked()] or ["eng"]
                if remember:
                    store.setValue(f"opt/{key}", ",".join(out[key]))
                continue
            elif kind == "choice":
                out[key] = w.currentData()
            elif kind in ("text", "password"):
                out[key] = w.text()
            elif kind in ("int", "slider"):
                out[key] = w.value()
            if remember and kind != "password":
                store.setValue(f"opt/{key}", out[key])
        return out


def _attach_labels(form):
    """map each field key to the label above it (labels are added right before their input)"""
    lay = form.layout()
    prev = None
    for i in range(lay.count()):
        w = lay.itemAt(i).widget()
        if isinstance(w, QLabel) and w.objectName() == "dgField":
            prev = w
            continue
        for key, (kind, inp) in form.inputs.items():
            if inp is w and prev is not None:
                form._labels[key] = prev
        prev = None


# --------------------------------------------------------------------------------------------- the panel --
class ToolsPanel(QFrame):
    """search, recent & favourites, the grouped tool list; choosing a tool shows its options and a big button"""
    toolOpened = Signal(str)          # the window prepares the tool (page view modes, ...)
    toolClosed = Signal(str)
    runClicked = Signal(str, dict)
    stopClicked = Signal()
    previewClicked = Signal(str, dict)

    def __init__(self):
        super().__init__()
        self.setObjectName("dgTools")
        self.setFixedWidth(250)
        self.setStyleSheet(PANEL_QSS)
        self.current = None
        self.form = None
        self.context = {"has_doc": False, "batch": 0, "selected": 0}
        self.stack = QStackedWidget(self)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.stack)

        # page 1: the list
        lp = QWidget()
        lp.setObjectName("dgWhite")
        lv = QVBoxLayout(lp)
        lv.setContentsMargins(12, 12, 12, 8)
        lv.setSpacing(4)
        self.search = QLineEdit()
        self.search.setPlaceholderText("Search tools")
        self.search.setClearButtonEnabled(True)
        self.search.textChanged.connect(self.fill)
        lv.addWidget(self.search)
        mag = QLabel(self.search)
        mag.setPixmap(ui.line_icon("search", 16, SOFT).pixmap(16, 16))
        mag.move(9, 9)
        mag.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.list = QListWidget()
        self.list.setIconSize(QSize(20, 20))
        self.list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list.itemClicked.connect(self._clicked)
        self.list.itemActivated.connect(self._clicked)
        self.list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list.customContextMenuRequested.connect(self._menu)
        lv.addWidget(self.list, 1)
        hint = QLabel("Right-click a tool to add it to your favourites")
        hint.setObjectName("dgTip")
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color:{SOFT}; font-size:10.5px;")
        lv.addWidget(hint)
        self.stack.addWidget(lp)

        # page 2: one tool
        self.tool_page = QWidget()
        self.stack.addWidget(self.tool_page)
        self.fill()

    # ---- list
    def favourites(self):
        v = settings().value("tools/fav", []) or []
        return [k for k in (v if isinstance(v, list) else [v]) if k in ops.TOOL]

    def recent(self):
        v = settings().value("tools/recent", []) or []
        return [k for k in (v if isinstance(v, list) else [v]) if k in ops.TOOL]

    def _remember(self, key):
        r = [key] + [k for k in self.recent() if k != key]
        settings().setValue("tools/recent", r[:4])

    def _section(self, text):
        it = QListWidgetItem(text.upper())
        it.setFlags(Qt.NoItemFlags)
        f = QFont()
        f.setPointSizeF(8.5)
        f.setBold(True)
        it.setFont(f)
        it.setForeground(QColor(SOFT))
        it.setSizeHint(QSize(200, 30))
        self.list.addItem(it)

    def _tool_item(self, t, star=False):
        it = QListWidgetItem(ui.line_icon(t.symbol, 20, ui.ICON_INK), "  " + t.name + ("  ★" if star else ""))
        it.setData(Qt.UserRole, t.key)
        it.setToolTip(t.tip)
        enabled = self._usable(t)
        if not enabled:
            it.setForeground(QColor("#B4AAC4"))
        self.list.addItem(it)

    def _usable(self, t):
        c = self.context
        if t.key in ("merge", "img2pdf"):
            return True
        if c["batch"]:
            return t.batch
        return c["has_doc"]

    def fill(self):
        self.list.clear()
        q = self.search.text().strip().lower()
        favs = self.favourites()
        if q:
            hits = [t for t in ops.TOOLS if q in t.name.lower() or q in t.tip.lower() or q in t.group.lower()]
            self._section(f"{len(hits)} found")
            for t in hits:
                self._tool_item(t, t.key in favs)
            return
        top = favs + [k for k in self.recent() if k not in favs]
        if top:
            self._section("★  Recent & favourites")
            for k in top[:6]:
                self._tool_item(ops.TOOL[k], k in favs)
        for g in ops.GROUPS:
            self._section(g)
            for t in ops.TOOLS:
                if t.group == g:
                    self._tool_item(t, t.key in favs)

    def _menu(self, pos):
        it = self.list.itemAt(pos)
        key = it.data(Qt.UserRole) if it else None
        if not key:
            return
        favs = self.favourites()
        m = QMenu(self)
        if key in favs:
            a = m.addAction("Remove from favourites")
        else:
            a = m.addAction("Add to favourites ★")
        if m.exec(self.list.mapToGlobal(pos)) is a:
            favs = [k for k in favs if k != key] if key in favs else favs + [key]
            settings().setValue("tools/fav", favs)
            self.fill()

    def _clicked(self, it):
        key = it.data(Qt.UserRole)
        if not key:
            return
        t = ops.TOOL[key]
        if not self._usable(t):
            self.list.clearSelection()
            return
        self.open_tool(key)

    def set_context(self, has_doc, batch=0, selected=0):
        self.context = {"has_doc": has_doc, "batch": batch, "selected": selected}
        if self.current is None:
            self.fill()
        elif self.current and not self._usable(ops.TOOL[self.current]):
            self.back()
        else:
            self._update_button()

    # ---- one tool
    def open_tool(self, key):
        if self.current:
            self.toolClosed.emit(self.current)
        t = ops.TOOL[key]
        self.current = key
        self._remember(key)
        old = self.tool_page
        self.tool_page = QWidget()
        self.stack.insertWidget(1, self.tool_page)
        self.stack.removeWidget(old)
        old.deleteLater()
        outer = QVBoxLayout(self.tool_page)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        inner = QWidget()
        inner.setObjectName("dgWhite")
        v = QVBoxLayout(inner)
        v.setContentsMargins(14, 12, 14, 10)
        v.setSpacing(8)
        head = QHBoxLayout()
        head.setSpacing(4)
        back = QToolButton()
        back.setObjectName("dgBack")
        back.setIcon(ui.line_icon("back", 18))
        back.setIconSize(QSize(18, 18))
        back.setToolTip("Back to all tools")
        back.clicked.connect(self.back)
        head.addWidget(back)
        title = QLabel(t.name)
        title.setObjectName("dgTitle")
        head.addWidget(title, 1)
        v.addLayout(head)
        tip = QLabel(t.tip + ("." if not t.tip.endswith(".") else ""))
        tip.setObjectName("dgTip")
        tip.setWordWrap(True)
        v.addWidget(tip)
        self.form = OptionsForm(t.fields, self.context["selected"]) if t.fields else None
        if self.form:
            v.addWidget(self.form)
            self.form.changed.connect(self._update_button)
        self.ui_box = QVBoxLayout()                    # the window puts its own controls here (sign, edit text ...)
        self.ui_box.setSpacing(6)
        v.addLayout(self.ui_box)
        if t.key == "ocr" and not scan.tessdata_ok():
            warn = QLabel("The OCR language files are missing - please reinstall Digonto PDF.")
            warn.setWordWrap(True)
            warn.setStyleSheet("color:#B3261E;")
            v.addWidget(warn)
        v.addStretch(1)
        scroll.setWidget(inner)
        outer.addWidget(scroll, 1)

        foot = QWidget()
        foot.setObjectName("dgWhite")
        fv = QVBoxLayout(foot)
        fv.setContentsMargins(14, 6, 14, 14)
        fv.setSpacing(6)
        self.result = QLabel("")
        self.result.setObjectName("dgResult")
        self.result.setWordWrap(True)
        self.result.setTextInteractionFlags(Qt.TextSelectableByMouse)
        fv.addWidget(self.result)
        self.prog_text = QLabel("")
        self.prog_text.setObjectName("dgTip")
        fv.addWidget(self.prog_text)
        self.prog = QProgressBar()
        self.prog.setRange(0, 1000)
        self.prog.setTextVisible(False)
        fv.addWidget(self.prog)
        self.prog.hide()
        self.prog_text.hide()
        row = QHBoxLayout()
        self.preview_btn = None
        if t.preview:
            self.preview_btn = QPushButton("Preview")
            self.preview_btn.setToolTip("Before / after on the page you are looking at")
            self.preview_btn.clicked.connect(lambda: self.previewClicked.emit(self.current, self.values()))
            row.addWidget(self.preview_btn)
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.clicked.connect(self.stopClicked.emit)
        self.stop_btn.hide()
        row.addWidget(self.stop_btn)
        fv.addLayout(row)
        self.run_btn = QPushButton(t.button)
        self.run_btn.setObjectName("dgPrimary")
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.clicked.connect(lambda: self.runClicked.emit(self.current, self.values()))
        fv.addWidget(self.run_btn)
        if t.kind == "ui" and t.key != "redact":
            self.run_btn.hide()
        outer.addWidget(foot)
        self.stack.setCurrentIndex(1)
        self._update_button()
        self.toolOpened.emit(key)

    def values(self):
        return self.form.values() if self.form else {}

    def _update_button(self):
        if not self.current or not hasattr(self, "run_btn"):
            return
        t = ops.TOOL[self.current]
        n = self.context["batch"]
        label = t.button
        if n and t.batch and t.kind in ("doc", "files"):
            label = f"{t.button} - {n} files" if n > 1 else t.button
        self.run_btn.setText(label)

    def back(self):
        key = self.current
        self.current = None
        self.stack.setCurrentIndex(0)
        self.fill()
        if key:
            self.toolClosed.emit(key)

    # ---- running
    def busy(self, flag):
        if not hasattr(self, "run_btn"):
            return
        self.run_btn.setEnabled(not flag)
        self.stop_btn.setVisible(flag)
        self.prog.setVisible(flag)
        self.prog_text.setVisible(flag)
        if self.preview_btn:
            self.preview_btn.setEnabled(not flag)
        if flag:
            self.prog.setValue(0)
            self.result.setText("")

    def show_progress(self, frac, text=""):
        if hasattr(self, "prog"):
            self.prog.setValue(int(frac * 1000))
            self.prog_text.setText(text)

    def show_result(self, text, error=False):
        if hasattr(self, "result"):
            self.result.setStyleSheet("color:#B3261E;" if error else "")
            self.result.setText(text)


# --------------------------------------------------------------------------------------------- batch view --
class BatchView(QWidget):
    """many files at once: one row per file with before → after, then Save all"""
    saveAll = Signal()
    stop = Signal()
    clear = Signal()
    addMore = Signal()

    def __init__(self):
        super().__init__()
        self.setObjectName("dgStart")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setStyleSheet(ui.SHELL_QSS + f"QFrame#dgCard {{ background: white; border: 1px solid {LINE}; "
                           "border-radius: 12px; } QLabel { background: transparent; }")
        self.files = []                 # dicts: path, name, pages, size, result, msg, err
        v = QVBoxLayout(self)
        v.setContentsMargins(28, 20, 28, 20)
        v.setSpacing(10)
        self.title = QLabel()
        self.title.setStyleSheet("font-size:16px; font-weight:700;")
        v.addWidget(self.title)
        sub = QLabel("Choose a tool on the right - it runs on every file in the list. Nothing is changed until you "
                     "press Save all.")
        sub.setStyleSheet(f"color:{SOFT};")
        sub.setWordWrap(True)
        v.addWidget(sub)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setStyleSheet("QScrollArea, QScrollArea > QWidget > QWidget { background: transparent; }")
        self.card = QFrame()
        self.card.setObjectName("dgCard")
        self.rows = QVBoxLayout(self.card)
        self.rows.setContentsMargins(0, 6, 0, 6)
        self.rows.setSpacing(0)
        holder = QWidget()
        hv = QVBoxLayout(holder)
        hv.setContentsMargins(0, 0, 0, 0)
        hv.addWidget(self.card)
        hv.addStretch(1)
        scroll.setWidget(holder)
        v.addWidget(scroll, 1)
        self.total = QLabel("")
        v.addWidget(self.total)
        row = QHBoxLayout()
        b_add = QPushButton("  Add more")
        b_add.setObjectName("dgPlain")
        b_add.setIcon(ui.line_icon("new", 18))
        b_add.clicked.connect(self.addMore.emit)
        row.addWidget(b_add)
        b_clear = QPushButton("Close list")
        b_clear.setObjectName("dgPlain")
        b_clear.clicked.connect(self.clear.emit)
        row.addWidget(b_clear)
        row.addStretch(1)
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setObjectName("dgPlain")
        self.stop_btn.clicked.connect(self.stop.emit)
        self.stop_btn.hide()
        row.addWidget(self.stop_btn)
        self.save_btn = QPushButton("  Save all …")
        self.save_btn.setObjectName("dgPrimary")
        self.save_btn.setIcon(ui.line_icon("save", 18))
        self.save_btn.clicked.connect(self.saveAll.emit)
        self.save_btn.setEnabled(False)
        row.addWidget(self.save_btn)
        v.addLayout(row)
        self.labels = []
        self.folder_name = ""

    def set_files(self, paths, folder_name=""):
        known = {f["path"] for f in self.files}
        for p in paths:
            if p in known:
                continue
            try:
                d = fitz.open(p)
                pages = d.page_count
                locked = d.needs_pass
                d.close()
            except Exception:                                  # noqa: BLE001
                continue
            self.files.append(dict(path=p, name=os.path.basename(p), pages=pages, size=os.path.getsize(p),
                                   result=None, msg="", err="", locked=locked))
        if folder_name:
            self.folder_name = folder_name
        self._rebuild()

    def reset_results(self):
        for f in self.files:
            f.update(result=None, msg="", err="")
        self._rebuild()

    def _rebuild(self):
        while self.rows.count():
            w = self.rows.takeAt(0).widget()
            if w:
                w.hide()
                w.setParent(None)
                w.deleteLater()
        self.labels = []
        for i, f in enumerate(self.files):
            r = QWidget()
            h = QHBoxLayout(r)
            h.setContentsMargins(16, 9, 16, 9)
            ic = QLabel()
            ic.setPixmap(ui.line_icon("lock" if f["locked"] else "file", 20, ui.ICON_INK).pixmap(20, 20))
            h.addWidget(ic)
            pg = f"{f['pages']} page" + ("s" if f["pages"] != 1 else "")
            nl = QLabel(f"<b>{f['name']}</b>  <span style='color:{SOFT}'>{pg}</span>")
            nl.setToolTip(f["path"])
            h.addWidget(nl, 1)
            st = QLabel()
            h.addWidget(st)
            self.labels.append(st)
            self.rows.addWidget(r)
            if i < len(self.files) - 1:
                s = QFrame()
                s.setFixedHeight(1)
                s.setStyleSheet(f"background:{LINE}; margin: 0 16px;")
                self.rows.addWidget(s)
            self._show(i)
        where = f" from “{self.folder_name}”" if self.folder_name else ""
        self.title.setText(f"{len(self.files)} files{where}")
        self._totals()

    def _show(self, i, working=None):
        f = self.files[i]
        lab = self.labels[i]
        size = ops.human(f["size"])
        if working is not None:
            lab.setText(f"<span style='color:{SOFT}'>{size}</span>  →  working … {int(working * 100)} %")
        elif f["err"]:
            lab.setText(f"<span style='color:#B3261E'>{f['err'][:60]}</span>")
        elif isinstance(f["result"], (bytes, bytearray)):
            lab.setText(f"<span style='color:{SOFT}'>{size}</span>  →  <b>{ops.human(len(f['result']))}</b>"
                        "   <span style='color:#2E9E62'>✓</span>")
        elif isinstance(f["result"], list):
            lab.setText(f"<b>{len(f['result'])} file(s)</b>   <span style='color:#2E9E62'>✓</span>")
        elif f["msg"]:
            lab.setText(f"<span style='color:{SOFT}'>{f['msg']}</span>")
        else:
            lab.setText(f"<span style='color:{SOFT}'>{size}</span>")

    def progress(self, i, frac):
        if 0 <= i < len(self.labels):
            self._show(i, frac)

    def job_done(self, i, result, msg, err):
        f = self.files[i]
        f.update(result=result, msg=msg, err=err)
        self._show(i)
        self._totals()

    def _totals(self):
        done = [f for f in self.files if isinstance(f["result"], (bytes, bytearray))]
        if done:
            before = sum(f["size"] for f in done)
            after = sum(len(f["result"]) for f in done)
            if after < before:
                self.total.setText(f"<span style='color:{SOFT}'>Saved so far:</span> <b>{ops.human(before - after)}</b>"
                                   f"  ({100 - after * 100 // max(before, 1)} % smaller)")
            else:
                self.total.setText(f"<span style='color:{SOFT}'>{len(done)} file(s) ready to save</span>")
        else:
            self.total.setText("")
        self.save_btn.setEnabled(any(f["result"] for f in self.files))

    def busy(self, flag):
        self.stop_btn.setVisible(flag)
        self.save_btn.setEnabled(not flag and any(f["result"] for f in self.files))


# -------------------------------------------------------------------------------------- before / after --
class BeforeAfter(QWidget):
    """two pictures of the same page; drag the line to compare"""
    closed = Signal()

    def __init__(self):
        super().__init__()
        self.before = self.after = None
        self.split = 0.5
        self.setMouseTracking(True)
        self.setCursor(Qt.SplitHCursor)
        self.close_btn = QPushButton("  Close preview", self)
        self.close_btn.setObjectName("dgPlain")
        self.close_btn.setStyleSheet(ui.SHELL_QSS)
        self.close_btn.setIcon(ui.line_icon("close", 16))
        self.close_btn.clicked.connect(self.closed.emit)
        self.note = ""

    def set_images(self, before, after, note=""):
        self.before, self.after, self.note = before, after, note
        self.split = 0.5
        self.update()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.close_btn.adjustSize()
        self.close_btn.move(self.width() - self.close_btn.width() - 14, 12)

    def _rect(self):
        if self.before is None:
            return QRectF()
        w, h = self.before.width(), self.before.height()
        avail = QRectF(20, 54, self.width() - 40, self.height() - 74)
        s = min(avail.width() / w, avail.height() / h)
        rw, rh = w * s, h * s
        return QRectF(avail.center().x() - rw / 2, avail.center().y() - rh / 2, rw, rh)

    def mousePressEvent(self, e):
        self._drag(e.position())

    def mouseMoveEvent(self, e):
        if e.buttons() & Qt.LeftButton:
            self._drag(e.position())

    def _drag(self, pos):
        r = self._rect()
        if r.width() > 0:
            self.split = min(1.0, max(0.0, (pos.x() - r.x()) / r.width()))
            self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.SmoothPixmapTransform)
        p.fillRect(self.rect(), QColor("#EEE9F4"))
        if self.before is None:
            return
        r = self._rect()
        sx = r.x() + r.width() * self.split
        p.drawImage(r, self.after)
        p.save()
        p.setClipRect(QRectF(r.x(), r.y(), sx - r.x(), r.height()))
        p.drawImage(r, self.before)
        p.restore()
        p.setPen(QPen(QColor("#DDD5E8"), 1))
        p.drawRect(r)
        p.setPen(QPen(QColor(PINK), 2.5))
        p.drawLine(QPointF(sx, r.y()), QPointF(sx, r.bottom()))
        p.setBrush(QColor("white"))
        p.drawEllipse(QPointF(sx, r.center().y()), 15, 15)
        p.setPen(QPen(QColor(PINK), 2))
        f = QFont()
        f.setBold(True)
        p.setFont(f)
        p.drawText(QRectF(sx - 15, r.center().y() - 15, 30, 30), Qt.AlignCenter, "‹ ›")
        f.setPointSizeF(9)
        p.setFont(f)
        for txt, x in (("BEFORE", r.x() + 10), ("AFTER", r.right() - 70)):
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(42, 18, 65, 190))
            p.drawRoundedRect(QRectF(x, r.y() + 10, 60, 22), 6, 6)
            p.setPen(QColor("white"))
            p.drawText(QRectF(x, r.y() + 10, 60, 22), Qt.AlignCenter, txt)
        if self.note:
            p.setPen(QColor(INK))
            f.setPointSizeF(11)
            p.setFont(f)
            p.drawText(QRectF(20, 12, self.width() - 200, 30), Qt.AlignVCenter | Qt.AlignLeft, self.note)


def page_qimage(page, width_px=1100):
    z = width_px / max(page.rect.width, 1)
    pix = page.get_pixmap(matrix=fitz.Matrix(z, z), alpha=False)
    return QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()


def preview_images(tool_key, doc, pno, opts):
    """(before, after, note) for the before/after slider - only the one page is worked on"""
    one = fitz.open()
    one.insert_pdf(doc, from_page=pno, to_page=pno)
    before = page_qimage(one[0])
    data = one.tobytes()
    if tool_key == "compress":
        out = ops.shrink_bytes(data, opts.get("preset", ops.PRESETS[1][0]))
        note = f"This page: {ops.human(len(data))} → {ops.human(len(out))}"
    else:
        _, out = ops.run_doc_tool(ops.TOOL[tool_key], data, dict(opts, which="all"))
        note = "Drag the line to compare"
    after = page_qimage(fitz.open(stream=out, filetype="pdf")[0])
    if after.size() != before.size():
        after = after.scaled(before.size(), Qt.IgnoreAspectRatio, Qt.SmoothTransformation)
    return before, after, note
