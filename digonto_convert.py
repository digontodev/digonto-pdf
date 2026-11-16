"""
digonto_convert.py - Digonto PDF's own converters. No Office, no LibreOffice.

    images (png, jpg, bmp, gif, tif, webp) -> PDF
    svg                                     -> PDF (stays vector)
    Word  .docx                             -> PDF (own reader + PyMuPDF layout)
    Excel .xlsx / .xlsm                     -> PDF (own reader + PyMuPDF layout)
    text  .txt                              -> PDF
    pdf                                     -> copied in as it is

Only needs PyMuPDF. Word/Excel files are read directly from their ZIP/XML
contents. The goal is a clean, readable PDF with the same text, tables,
pictures, bold/italic/colors and headings - not a pixel copy of Office.
Optional: convert_with_office() uses LibreOffice / MS Office when the user
asks for an exact copy and one of them is installed.
"""
import datetime
import html
import io
import os
import posixpath
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

try:
    import pymupdf as fitz
except ImportError:
    import fitz

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".jfif"}
SUPPORTED_EXT = IMAGE_EXT | {".svg", ".docx", ".docm", ".xlsx", ".xlsm", ".txt", ".pdf"}

PAPER = {"A4": (595.28, 841.89), "A5": (419.53, 595.28), "A3": (841.89, 1190.55),
         "Letter": (612.0, 792.0), "Legal": (612.0, 1008.0)}


def kind_of(path):
    ext = os.path.splitext(path)[1].lower()
    if ext in IMAGE_EXT:
        return "image"
    return {".svg": "svg", ".docx": "word", ".docm": "word", ".xlsx": "excel", ".xlsm": "excel",
            ".txt": "text", ".pdf": "pdf"}.get(ext)


# --------------------------------------------------------------------------
# HTML -> PDF with PyMuPDF Story (used by Word, Excel, text)
# --------------------------------------------------------------------------
BASE_CSS = """
body { font-family: sans-serif; font-size: 11pt; line-height: 1.25; }
p { margin: 0 0 6pt 0; }
h1 { font-size: 20pt; margin: 10pt 0 6pt 0; }
h2 { font-size: 16pt; margin: 10pt 0 5pt 0; }
h3 { font-size: 13pt; margin: 8pt 0 4pt 0; }
h4, h5, h6 { font-size: 11.5pt; margin: 6pt 0 3pt 0; }
table { border-collapse: collapse; margin: 4pt 0 8pt 0; }
td, th { padding: 2pt 4pt; vertical-align: top; }
.pb { page-break-before: always; }
"""


def html_to_pdf(html_body, page_w, page_h, margins=(56, 56, 56, 56), images=None, css=""):
    """Lay out HTML on as many pages as needed. Returns a fitz.Document."""
    arch = fitz.Archive()
    for name, data in (images or {}).items():
        arch.add((data, name))
    story = fitz.Story(html=f"<body>{html_body}</body>", user_css=BASE_CSS + css, archive=arch)
    buf = io.BytesIO()
    writer = fitz.DocumentWriter(buf)
    mediabox = fitz.Rect(0, 0, page_w, page_h)
    left, top, right, bottom = margins
    where = fitz.Rect(left, top, page_w - right, page_h - bottom)
    more = 1
    pages = 0
    while more:
        dev = writer.begin_page(mediabox)
        more, _ = story.place(where)
        story.draw(dev)
        writer.end_page()
        pages += 1
        if pages > 5000:
            break
    writer.close()
    return fitz.open(stream=buf.getvalue(), filetype="pdf")


def html_to_pdf_fit_width(html_body, pw, ph, margin, css="", images=None, est_width=0):
    """Like html_to_pdf, but content wider than the paper is scaled down to
    fit the page width (like Excel's 'fit all columns on one page')."""
    avail_w, avail_h = pw - 2 * margin, ph - 2 * margin
    vw = max(avail_w, est_width)
    tmp = None
    for _ in range(4):
        vh = vw * avail_h / avail_w
        tmp = html_to_pdf(html_body, vw + 4, vh + 4, (2, 2, 2, 2), images, css)
        right = 0
        for p in tmp:
            for item in p.get_bboxlog():
                r = fitz.Rect(item[1])
                if not r.is_empty and not r.is_infinite:
                    right = max(right, r.x1)
        if right <= vw + 6:
            break
        vw = right + 12
    out = fitz.open()
    for i in range(tmp.page_count):
        page = out.new_page(width=pw, height=ph)
        page.show_pdf_page(fitz.Rect(margin, margin, pw - margin, ph - margin), tmp, i)
    return out


def esc(t):
    return html.escape(t, quote=True)


# --------------------------------------------------------------------------
# Word .docx
# --------------------------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
V = "urn:schemas-microsoft-com:vml"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"


def w(tag):
    return f"{{{W}}}{tag}"


def wval(el, tag, attr="val"):
    if el is None:
        return None
    c = el.find(w(tag))
    if c is None:
        return None
    return c.get(w(attr))


def on_off(el, tag):
    if el is None:
        return None
    c = el.find(w(tag))
    if c is None:
        return None
    v = c.get(w("val"))
    return v not in ("0", "false", "off", "none")


HIGHLIGHT = {"yellow": "#ffff00", "green": "#00ff00", "cyan": "#00ffff", "magenta": "#ff00ff",
             "blue": "#0000ff", "red": "#ff0000", "darkBlue": "#000080", "darkCyan": "#008080",
             "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
             "darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#c0c0c0", "black": "#000000"}


def read_rels(z, path):
    rels = {}
    rp = posixpath.join(posixpath.dirname(path), "_rels", posixpath.basename(path) + ".rels")
    if rp not in z.namelist():
        return rels
    root = ET.fromstring(z.read(rp))
    for r in root.findall(f"{{{PR}}}Relationship"):
        tgt = r.get("Target")
        if r.get("TargetMode") == "External":
            rels[r.get("Id")] = ("ext", tgt)
        elif tgt.startswith("/"):
            rels[r.get("Id")] = ("int", tgt.lstrip("/"))
        else:
            rels[r.get("Id")] = ("int", posixpath.normpath(posixpath.join(posixpath.dirname(path), tgt)))
    return rels


class DocxConverter:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.rels = read_rels(self.z, "word/document.xml")
        self.images = {}
        self.styles = {}        # styleId -> dict(name, basedOn, ppr, rpr)
        self.num_fmt = {}       # (numId, ilvl) -> (fmt, lvlText, start)
        self.counters = {}
        self.default_size = 11.0
        self.default_font = None
        self._load_styles()
        self._load_numbering()

    # ---- styles & numbering
    def _load_styles(self):
        if "word/styles.xml" not in self.z.namelist():
            return
        root = ET.fromstring(self.z.read("word/styles.xml"))
        dd = root.find(w("docDefaults"))
        if dd is not None:
            rpr = dd.find(f"{w('rPrDefault')}/{w('rPr')}")
            sz = wval(rpr, "sz")
            if sz:
                self.default_size = int(sz) / 2
            fonts = rpr.find(w("rFonts")) if rpr is not None else None
            if fonts is not None:
                self.default_font = fonts.get(w("ascii"))
        for st in root.findall(w("style")):
            sid = st.get(w("styleId"))
            name = wval(st, "name") or sid
            self.styles[sid] = dict(name=name, basedOn=wval(st, "basedOn"), ppr=st.find(w("pPr")),
                                    rpr=st.find(w("rPr")), type=st.get(w("type")))

    def style_chain(self, sid):
        out = []
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            out.append(self.styles[sid])
            sid = self.styles[sid]["basedOn"]
        return out

    def _load_numbering(self):
        if "word/numbering.xml" not in self.z.namelist():
            return
        root = ET.fromstring(self.z.read("word/numbering.xml"))
        abstract = {}
        for an in root.findall(w("abstractNum")):
            lv = {}
            for l in an.findall(w("lvl")):
                lv[int(l.get(w("ilvl")))] = (wval(l, "numFmt") or "bullet", wval(l, "lvlText") or "",
                                             int(wval(l, "start") or 1))
            abstract[an.get(w("abstractNumId"))] = lv
        for num in root.findall(w("num")):
            aid = wval(num, "abstractNumId")
            for ilvl, v in abstract.get(aid, {}).items():
                self.num_fmt[(num.get(w("numId")), ilvl)] = v

    # ---- page setup
    def page_setup(self, body):
        sp = body.find(w("sectPr"))
        pw, ph = PAPER["A4"]
        m = [72, 72, 72, 72]
        if sp is not None:
            pg = sp.find(w("pgSz"))
            if pg is not None and pg.get(w("w")):
                pw, ph = int(pg.get(w("w"))) / 20, int(pg.get(w("h"))) / 20
            pm = sp.find(w("pgMar"))
            if pm is not None:
                def g(k, d):
                    v = pm.get(w(k))
                    try:
                        return max(18, abs(int(v)) / 20)
                    except (TypeError, ValueError):
                        return d
                m = [g("left", 72), g("top", 72), g("right", 72), g("bottom", 72)]
        return pw, ph, m

    # ---- runs
    def run_props(self, rpr_list):
        """Merge run properties (later entries win)."""
        p = {}
        for rpr in rpr_list:
            if rpr is None:
                continue
            for k in ("b", "i", "strike", "dstrike", "caps", "smallCaps", "vanish"):
                v = on_off(rpr, k)
                if v is not None:
                    p[k] = v
            u = wval(rpr, "u")
            if u is not None:
                p["u"] = u != "none"
            c = wval(rpr, "color")
            if c and c != "auto":
                p["color"] = "#" + c
            sz = wval(rpr, "sz")
            if sz:
                p["size"] = int(sz) / 2
            hl = wval(rpr, "highlight")
            if hl and hl in HIGHLIGHT:
                p["bg"] = HIGHLIGHT[hl]
            shd = rpr.find(w("shd"))
            if shd is not None and shd.get(w("fill")) not in (None, "auto", "FFFFFF"):
                p["bg"] = "#" + shd.get(w("fill"))
            va = wval(rpr, "vertAlign")
            if va:
                p["va"] = va
            f = rpr.find(w("rFonts"))
            if f is not None and f.get(w("ascii")):
                p["font"] = f.get(w("ascii"))
        return p

    @staticmethod
    def font_family(name):
        n = (name or "").lower()
        if any(k in n for k in ("courier", "consol", "mono")):
            return "monospace"
        if any(k in n for k in ("times", "georgia", "garamond", "cambria", "serif", "palatino", "book")) \
                and "sans" not in n:
            return "serif"
        return "sans-serif"

    def css_for(self, p, para_size=None):
        s = []
        if p.get("b"):
            s.append("font-weight:bold")
        if p.get("i"):
            s.append("font-style:italic")
        deco = []
        if p.get("u"):
            deco.append("underline")
        if p.get("strike") or p.get("dstrike"):
            deco.append("line-through")
        if deco:
            s.append("text-decoration:" + " ".join(deco))
        if p.get("color"):
            s.append(f"color:{p['color']}")
        if p.get("bg"):
            s.append(f"background-color:{p['bg']}")
        if p.get("size") and p.get("size") != para_size:
            s.append(f"font-size:{p['size']}pt")
        if p.get("font"):
            s.append(f"font-family:{self.font_family(p['font'])}")
        if p.get("caps"):
            s.append("text-transform:uppercase")
        if p.get("va") == "superscript":
            s.append("vertical-align:super;font-size:70%")
        elif p.get("va") == "subscript":
            s.append("vertical-align:sub;font-size:70%")
        return ";".join(s)

    def image_html(self, el):
        blip = el.find(f".//{{{A}}}blip")
        rid = None
        if blip is not None:
            rid = blip.get(f"{{{R}}}embed") or blip.get(f"{{{R}}}link")
        if rid is None:
            im = el.find(f".//{{{V}}}imagedata")
            if im is not None:
                rid = im.get(f"{{{R}}}id")
        if not rid or rid not in self.rels or self.rels[rid][0] != "int":
            return ""
        target = self.rels[rid][1]
        if target not in self.z.namelist():
            return ""
        ext = os.path.splitext(target)[1].lower()
        data = self.z.read(target)
        if ext in (".emf", ".wmf"):
            return ""   # Windows vector formats: not supported
        name = f"img{len(self.images)}{ext}"
        self.images[name] = data
        style = ""
        extent = el.find(f".//{{{WP}}}extent")
        if extent is not None:
            wpt = int(extent.get("cx", "0")) / 12700
            hpt = int(extent.get("cy", "0")) / 12700
            if wpt > 0 and hpt > 0:
                style = f' style="width:{wpt:.1f}pt;height:{hpt:.1f}pt"'
        return f'<img src="{name}"{style}/>'

    def runs_html(self, parent, para_rpr, para_size, page_break_flag):
        out = []
        for el in parent:
            tag = el.tag
            if tag == w("r"):
                rpr = el.find(w("rPr"))
                chain = [para_rpr]
                rs = wval(rpr, "rStyle")
                if rs:
                    chain += [s["rpr"] for s in reversed(self.style_chain(rs))]
                chain.append(rpr)
                p = self.run_props(chain)
                if p.get("vanish"):
                    continue
                parts = []
                for c in el:
                    if c.tag == w("t"):
                        parts.append(esc(c.text or ""))
                    elif c.tag == w("tab"):
                        parts.append("&#160;&#160;&#160;&#160;")
                    elif c.tag in (w("br"), w("cr")):
                        if c.get(w("type")) == "page":
                            page_break_flag.append(True)
                        else:
                            parts.append("<br/>")
                    elif c.tag == w("noBreakHyphen"):
                        parts.append("-")
                    elif c.tag in (w("drawing"), w("pict")) or c.tag.endswith("}object"):
                        parts.append(self.image_html(c))
                    elif c.tag == w("lastRenderedPageBreak"):
                        pass
                    elif c.tag == w("sym"):
                        try:
                            parts.append(chr(int(c.get(w("char")), 16)))
                        except (TypeError, ValueError):
                            pass
                text = "".join(parts)
                if not text:
                    continue
                css = self.css_for(p, para_size)
                out.append(f'<span style="{css}">{text}</span>' if css else text)
            elif tag == w("hyperlink"):
                inner = self.runs_html(el, para_rpr, para_size, page_break_flag)
                rid = el.get(f"{{{R}}}id")
                if rid and rid in self.rels and self.rels[rid][0] == "ext":
                    out.append(f'<a href="{esc(self.rels[rid][1])}">{inner}</a>')
                else:
                    out.append(inner)
            elif tag in (w("ins"), w("smartTag"), w("customXml"), w("fldSimple"), w("sdt"), w("sdtContent")):
                out.append(self.runs_html(el, para_rpr, para_size, page_break_flag))
            elif tag == "{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent":
                fb = el.find("{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback")
                ch = el.find("{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice")
                use = ch if ch is not None else fb
                if use is not None:
                    out.append(self.runs_html(use, para_rpr, para_size, page_break_flag))
        return "".join(out)

    def list_prefix(self, num_id, ilvl):
        fmt = self.num_fmt.get((num_id, ilvl))
        if not fmt or num_id == "0":
            return ""
        kind, text, start = fmt
        # reset deeper levels
        for k in list(self.counters):
            if k[0] == num_id and k[1] > ilvl:
                del self.counters[k]
        key = (num_id, ilvl)
        self.counters[key] = self.counters.get(key, start - 1) + 1
        if kind == "bullet":
            sym = text if text and text.isprintable() and ord(text[0]) < 0xF000 else ""
            return {"o": "\u25e6", "": "\u2022"}.get(sym, "\u2022") if sym in ("o", "") else sym

        def fmt_num(n, kd):
            if kd in ("lowerLetter",):
                return chr(ord("a") + (n - 1) % 26)
            if kd in ("upperLetter",):
                return chr(ord("A") + (n - 1) % 26)
            if kd in ("lowerRoman", "upperRoman"):
                vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"),
                        (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
                r, x = "", n
                for v, s in vals:
                    while x >= v:
                        r += s
                        x -= v
                return r.upper() if kd == "upperRoman" else r
            return str(n)

        def repl(m):
            lv = int(m.group(1)) - 1
            k = self.num_fmt.get((num_id, lv), ("decimal", "", 1))
            return fmt_num(self.counters.get((num_id, lv), k[2]), k[0])
        return re.sub(r"%(\d)", repl, text or "%1.")

    def paragraph_html(self, p_el, in_table=False):
        ppr = p_el.find(w("pPr"))
        sid = wval(ppr, "pStyle") or "Normal"
        chain = self.style_chain(sid)
        sname = (chain[0]["name"] if chain else sid).lower().replace(" ", "")
        # paragraph-level run defaults: styles (base first) then pPr/rPr
        rpr_chain = [s["rpr"] for s in reversed(chain)]
        rpr_chain.append(ppr.find(w("rPr")) if ppr is not None else None)
        para_p = self.run_props(rpr_chain)
        tag = "p"
        m = re.match(r"heading(\d)", sname)
        if m:
            tag = f"h{min(6, int(m.group(1)))}"
        elif sname == "title":
            tag = "h1"
        elif sname == "subtitle":
            tag = "h3"
        # alignment / indentation / spacing, own pPr then style pPr
        pprs = [ppr] + [s["ppr"] for s in chain]
        jc = next((wval(x, "jc") for x in pprs if wval(x, "jc")), None)
        styles = []
        if jc in ("center",):
            styles.append("text-align:center")
        elif jc in ("right", "end"):
            styles.append("text-align:right")
        elif jc in ("both", "distribute"):
            styles.append("text-align:justify")
        ind = next((x.find(w("ind")) for x in pprs if x is not None and x.find(w("ind")) is not None), None)
        left = 0
        if ind is not None:
            try:
                left = int(ind.get(w("left")) or ind.get(w("start")) or 0) / 20
            except ValueError:
                left = 0
            try:
                first = int(ind.get(w("firstLine")) or 0) / 20 - int(ind.get(w("hanging")) or 0) / 20
            except ValueError:
                first = 0
            if first and not in_table:
                styles.append(f"text-indent:{first:.0f}pt")
        spacing = next((x.find(w("spacing")) for x in pprs if x is not None and x.find(w("spacing")) is not None),
                       None)
        if spacing is not None and tag == "p":
            try:
                if spacing.get(w("after")) is not None:
                    styles.append(f"margin-bottom:{int(spacing.get(w('after'))) / 20:.0f}pt")
                if spacing.get(w("before")) is not None:
                    styles.append(f"margin-top:{int(spacing.get(w('before'))) / 20:.0f}pt")
            except ValueError:
                pass
        shd = ppr.find(w("shd")) if ppr is not None else None
        if shd is not None and shd.get(w("fill")) not in (None, "auto", "FFFFFF"):
            styles.append(f"background-color:#{shd.get(w('fill'))}")
        # list numbering (own pPr or style)
        prefix = ""
        numpr = next((x.find(w("numPr")) for x in pprs if x is not None and x.find(w("numPr")) is not None), None)
        if numpr is not None:
            num_id = wval(numpr, "numId")
            ilvl = int(wval(numpr, "ilvl") or 0)
            if num_id:
                prefix = self.list_prefix(num_id, ilvl)
                if prefix:
                    left = max(left, 18 * (ilvl + 1))
        if left and not in_table:
            styles.append(f"margin-left:{left:.0f}pt")
        para_css = self.css_for(para_p)
        if tag != "p":
            # headings: keep their own size, only take colour/italic from style
            para_css = ";".join(x for x in para_css.split(";") if x.startswith(("color", "font-style")))
        if para_css:
            styles.append(para_css)
        page_break = []
        if on_off(ppr, "pageBreakBefore"):
            page_break.append(True)
        inner = self.runs_html(p_el, None, para_p.get("size"), page_break)
        # the run css already merges para style via css_for(p, para_size); runs inherit from <p> anyway
        if prefix:
            inner = f"{esc(prefix)}&#160;&#160;{inner}"
        if not inner.strip():
            inner = "&#160;"
        style_attr = f' style="{";".join(styles)}"' if styles else ""
        out = f"<{tag}{style_attr}>{inner}</{tag}>"
        if page_break and not in_table:
            out = '<div class="pb"></div>' + out
        return out

    def table_html(self, tbl):
        rows = tbl.findall(w("tr"))
        tblpr = tbl.find(w("tblPr"))
        borders = True
        tb = tblpr.find(w("tblBorders")) if tblpr is not None else None
        if tb is not None:
            vals = [c.get(w("val")) for c in tb]
            if vals and all(v in ("nil", "none") for v in vals):
                borders = False
        else:
            sid = wval(tblpr, "tblStyle") or ""
            if sid and "grid" not in sid.lower() and "table" not in sid.lower():
                borders = False
        grid = [int(g.get(w("w")) or 0) / 20 for g in tbl.findall(f"{w('tblGrid')}/{w('gridCol')}")]
        bcss = "border:0.75pt solid #808080;" if borders else ""
        out = ['<table style="width:100%">' if sum(grid) > 400 else "<table>"]
        for tr in rows:
            out.append("<tr>")
            col = 0
            for tc in tr.findall(w("tc")):
                tcpr = tc.find(w("tcPr"))
                span = int(wval(tcpr, "gridSpan") or 1)
                vm = tcpr.find(w("vMerge")) if tcpr is not None else None
                cell_style = bcss
                if vm is not None and vm.get(w("val")) != "restart":
                    cell_style += "border-top:none;"
                shd = tcpr.find(w("shd")) if tcpr is not None else None
                if shd is not None and shd.get(w("fill")) not in (None, "auto"):
                    cell_style += f"background-color:#{shd.get(w('fill'))};"
                if grid and col < len(grid):
                    wd = sum(grid[col:col + span])
                    if wd > 0:
                        cell_style += f"width:{wd:.0f}pt;"
                col += span
                inner = []
                if not (vm is not None and vm.get(w("val")) != "restart"):
                    for c in tc:
                        if c.tag == w("p"):
                            inner.append(self.paragraph_html(c, in_table=True))
                        elif c.tag == w("tbl"):
                            inner.append(self.table_html(c))
                cs = f' colspan="{span}"' if span > 1 else ""
                out.append(f'<td{cs} style="{cell_style}">{"".join(inner) or "&#160;"}</td>')
            out.append("</tr>")
        out.append("</table>")
        return "".join(out)

    def body_html(self, body):
        out = []
        for el in body:
            if el.tag == w("p"):
                out.append(self.paragraph_html(el))
            elif el.tag == w("tbl"):
                out.append(self.table_html(el))
            elif el.tag == w("sdt"):
                content = el.find(w("sdtContent"))
                if content is not None:
                    out.append(self.body_html(content))
        return "".join(out)

    def convert(self):
        root = ET.fromstring(self.z.read("word/document.xml"))
        body = root.find(w("body"))
        pw, ph, m = self.page_setup(body)
        html_body = self.body_html(body)
        fam = self.font_family(self.default_font)
        css = f"body {{ font-size:{self.default_size}pt; font-family:{fam}; }}"
        return html_to_pdf(html_body, pw, ph, (m[0], m[1], m[2], m[3]), self.images, css)


def docx_to_pdf(path):
    return DocxConverter(path).convert()


# --------------------------------------------------------------------------
# Excel .xlsx
# --------------------------------------------------------------------------
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def s_(tag):
    return f"{{{S}}}{tag}"


BUILTIN_FMT = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%",
               11: "0.00E+00", 12: "# ?/?", 13: "# ??/??", 14: "dd.mm.yyyy", 15: "d-mmm-yy", 16: "d-mmm",
               17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "hh:mm", 21: "hh:mm:ss",
               22: "dd.mm.yyyy hh:mm", 37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)",
               39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss",
               47: "mmss.0", 48: "##0.0E+0", 49: "@"}

THEME_COLORS = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5",
                "70AD47", "0563C1", "954F72"]


def col_index(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)) - 1, col - 1


def excel_color(el):
    if el is None:
        return None
    rgb = el.get("rgb")
    if rgb:
        return "#" + rgb[-6:]
    th = el.get("theme")
    if th is not None:
        try:
            return "#" + THEME_COLORS[int(th)]
        except (ValueError, IndexError):
            return None
    return None


def format_number(v, fmt, date1904=False):
    """Simple version of Excel's number formats (good for the usual cases)."""
    if fmt in (None, "General", "@"):
        if float(v).is_integer() and abs(v) < 1e15:
            return str(int(v))
        return f"{v:.10g}"
    section = fmt.split(";")
    f = section[0]
    if v < 0 and len(section) > 1:
        f = section[1]
        v = -v
        neg = ""
    else:
        neg = "-" if v < 0 else ""
        v = abs(v)
    f_clean = re.sub(r"\[[^\]]*\]", lambda m: "" if not m.group(0).startswith("[$") else
                     m.group(0)[2:-1].split("-")[0], f)
    lit = re.sub(r'"([^"]*)"', r"\1", f_clean).replace("\\", "").replace("_)", "").replace("_(", "")
    low = re.sub(r'"[^"]*"', "", f_clean).lower()
    if re.search(r"[dmyhs]", low) and not re.search(r"[#0]", low.replace("00", "")) or \
            re.search(r"(^|[^a-z])(d{1,4}|y{2,4})([^a-z]|$)", low):
        base = datetime.datetime(1904, 1, 1) if date1904 else datetime.datetime(1899, 12, 30)
        try:
            dt = base + datetime.timedelta(days=v)
        except OverflowError:
            return str(v)
        out = low
        has_date = re.search(r"[dy]", out)
        has_time = re.search(r"[hs]", out)
        if has_date and has_time:
            return dt.strftime("%d.%m.%Y %H:%M")
        if has_time:
            return dt.strftime("%H:%M:%S" if "s" in out else "%H:%M")
        if "mmm" in out and "d" not in out:
            return dt.strftime("%b %Y")
        if "mmm" in out:
            return dt.strftime("%d %b %Y")
        if out.find("y") < out.find("d") and out.find("y") != -1:
            return dt.strftime("%Y-%m-%d")
        if "/" in out:
            return dt.strftime("%m/%d/%Y") if out.startswith("m") else dt.strftime("%d/%m/%Y")
        return dt.strftime("%d.%m.%Y")
    pct = "%" in low
    if pct:
        v *= 100
    m = re.search(r"[0#,]*\.([0#]+)", low)
    decimals = len(m.group(1).replace("#", "0")) if m else 0
    thousands = "," in low.split(".")[0]
    if "e+" in low:
        s = f"{v:.{decimals}E}"
    else:
        s = f"{v:,.{decimals}f}" if thousands else f"{v:.{decimals}f}"
    # currency / text around the number
    numeric = re.search(r"[#0][#0,.]*%?", lit)
    if numeric:
        prefix = lit[:numeric.start()].strip()
        suffix = lit[numeric.end():].strip()
        prefix = prefix.replace("#", "").replace("0", "")
        suffix = suffix.replace("%", "").strip()
        s = f"{prefix}{' ' if prefix and prefix[-1].isalpha() else ''}{s}{'%' if pct else ''}" \
            f"{(' ' + suffix) if suffix else ''}"
    elif pct:
        s += "%"
    return (neg + s).strip()


class XlsxConverter:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self.names = self.z.namelist()
        self.shared = []
        self.xfs = []
        self.numfmts = dict(BUILTIN_FMT)
        self.fonts = []
        self.fills = []
        self.borders = []
        self.date1904 = False
        self._load_shared()
        self._load_styles()

    def _load_shared(self):
        if "xl/sharedStrings.xml" not in self.names:
            return
        root = ET.fromstring(self.z.read("xl/sharedStrings.xml"))
        for si in root.findall(s_("si")):
            self.shared.append("".join(t.text or "" for t in si.iter(s_("t"))
                                       if not self._in_phonetic(si, t)))

    @staticmethod
    def _in_phonetic(si, t):
        for rph in si.findall(s_("rPh")):
            if t in list(rph.iter()):
                return True
        return False

    def _load_styles(self):
        if "xl/styles.xml" not in self.names:
            return
        root = ET.fromstring(self.z.read("xl/styles.xml"))
        nf = root.find(s_("numFmts"))
        if nf is not None:
            for n in nf.findall(s_("numFmt")):
                self.numfmts[int(n.get("numFmtId"))] = n.get("formatCode")
        fonts = root.find(s_("fonts"))
        if fonts is not None:
            for f in fonts.findall(s_("font")):
                d = {}
                if f.find(s_("b")) is not None and f.find(s_("b")).get("val") not in ("0", "false"):
                    d["b"] = True
                if f.find(s_("i")) is not None and f.find(s_("i")).get("val") not in ("0", "false"):
                    d["i"] = True
                if f.find(s_("u")) is not None:
                    d["u"] = True
                sz = f.find(s_("sz"))
                if sz is not None:
                    d["size"] = float(sz.get("val"))
                c = excel_color(f.find(s_("color")))
                if c and c.lower() != "#000000":
                    d["color"] = c
                self.fonts.append(d)
        fills = root.find(s_("fills"))
        if fills is not None:
            for f in fills.findall(s_("fill")):
                pf = f.find(s_("patternFill"))
                c = None
                if pf is not None and pf.get("patternType") not in (None, "none"):
                    c = excel_color(pf.find(s_("fgColor")))
                self.fills.append(c)
        borders = root.find(s_("borders"))
        if borders is not None:
            for b in borders.findall(s_("border")):
                d = {}
                for side in ("left", "right", "top", "bottom"):
                    e = b.find(s_(side))
                    if e is not None and e.get("style"):
                        d[side] = e.get("style")
                self.borders.append(d)
        cx = root.find(s_("cellXfs"))
        if cx is not None:
            for x in cx.findall(s_("xf")):
                al = x.find(s_("alignment"))
                self.xfs.append(dict(numFmt=int(x.get("numFmtId", 0)), font=int(x.get("fontId", 0)),
                                     fill=int(x.get("fillId", 0)), border=int(x.get("borderId", 0)),
                                     h=al.get("horizontal") if al is not None else None,
                                     wrap=al is not None and al.get("wrapText") in ("1", "true")))

    def sheets(self):
        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        pr = wb.find(s_("workbookPr"))
        if pr is not None and pr.get("date1904") in ("1", "true"):
            self.date1904 = True
        rels = read_rels(self.z, "xl/workbook.xml")
        out = []
        for sh in wb.iter(s_("sheet")):
            if sh.get("state") in ("hidden", "veryHidden"):
                continue
            rid = sh.get(f"{{{R}}}id")
            if rid in rels and rels[rid][1] in self.names:
                out.append((sh.get("name"), rels[rid][1]))
        return out

    def cell_text(self, c):
        t = c.get("t")
        v = c.find(s_("v"))
        style = self.xfs[int(c.get("s", 0))] if self.xfs and int(c.get("s", 0)) < len(self.xfs) else {}
        if t == "s" and v is not None:
            try:
                return self.shared[int(v.text)], style, False
            except (ValueError, IndexError):
                return "", style, False
        if t == "inlineStr":
            return "".join(x.text or "" for x in c.iter(s_("t"))), style, False
        if t in ("str", "e"):
            return (v.text or "") if v is not None else "", style, False
        if t == "b":
            return ("TRUE" if v is not None and v.text == "1" else "FALSE"), style, False
        if v is None or v.text is None:
            return "", style, False
        try:
            num = float(v.text)
        except ValueError:
            return v.text, style, False
        fmt = self.numfmts.get(style.get("numFmt", 0), "General")
        try:
            return format_number(num, fmt, self.date1904), style, True
        except Exception:
            return v.text, style, True

    def sheet_html(self, path, sheet_name, show_grid):
        root = ET.fromstring(self.z.read(path))
        cells = {}
        max_r = max_c = -1
        for row in root.iter(s_("row")):
            if row.get("hidden") in ("1", "true"):
                continue
            for c in row.findall(s_("c")):
                if not c.get("r"):
                    continue
                r, col = col_index(c.get("r"))
                text, style, is_num = self.cell_text(c)
                fill = self.fills[style.get("fill", 0)] if style and style.get("fill", 0) < len(self.fills) else None
                cells[(r, col)] = (text, style, is_num)
                if text.strip() or fill:
                    max_r, max_c = max(max_r, r), max(max_c, col)
        merges = {}
        covered = set()
        mc = root.find(s_("mergeCells"))
        if mc is not None:
            for m in mc.findall(s_("mergeCell")):
                a, b = m.get("ref").split(":") if ":" in m.get("ref") else (m.get("ref"), m.get("ref"))
                r1, c1 = col_index(a)
                r2, c2 = col_index(b)
                merges[(r1, c1)] = (r2 - r1 + 1, c2 - c1 + 1)
                for rr in range(r1, r2 + 1):
                    for cc in range(c1, c2 + 1):
                        if (rr, cc) != (r1, c1):
                            covered.add((rr, cc))
        widths = {}
        hidden_cols = set()
        cols = root.find(s_("cols"))
        if cols is not None:
            for col in cols.findall(s_("col")):
                for i in range(int(col.get("min")) - 1, min(int(col.get("max")), 16384)):
                    if col.get("hidden") in ("1", "true"):
                        hidden_cols.add(i)
                    if col.get("width"):
                        widths[i] = float(col.get("width")) * 5.6  # chars -> points (approx.)
        hidden_rows = {int(r.get("r")) - 1 for r in root.iter(s_("row")) if r.get("hidden") in ("1", "true")}
        if max_r < 0:
            return f"<h3>{esc(sheet_name)}</h3><p>(empty sheet)</p>", 0
        col_list = [c for c in range(max_c + 1) if c not in hidden_cols]
        need = {}
        for (r, c), (text, style, _n) in cells.items():
            if (r, c) in merges or (r, c) in covered or not text:
                continue
            longest = max(text.split("\n"), key=len)
            need[c] = max(need.get(c, 0), fitz.get_text_length(longest, "helv", 9) + 10)
        for c in col_list:
            widths[c] = max(widths.get(c, 48), need.get(c, 0))
        total_w = sum(widths.get(c, 48) for c in col_list)
        grid_css = "border:0.5pt solid #c8c8c8;" if show_grid else ""
        out = [f"<h3>{esc(sheet_name)}</h3><table>"]
        for r in range(max_r + 1):
            if r in hidden_rows:
                continue
            out.append("<tr>")
            for c in col_list:
                if (r, c) in covered:
                    continue
                text, style, is_num = cells.get((r, c), ("", {}, False))
                css = grid_css
                font = self.fonts[style.get("font", 0)] if style and style.get("font", 0) < len(self.fonts) else {}
                if font.get("b"):
                    css += "font-weight:bold;"
                if font.get("i"):
                    css += "font-style:italic;"
                if font.get("u"):
                    css += "text-decoration:underline;"
                if font.get("color"):
                    css += f"color:{font['color']};"
                if font.get("size") and abs(font["size"] - 11) > 0.6:
                    css += f"font-size:{font['size'] * 0.82:.1f}pt;"
                fill = self.fills[style.get("fill", 0)] if style and style.get("fill", 0) < len(self.fills) else None
                if fill:
                    css += f"background-color:{fill};"
                bd = self.borders[style.get("border", 0)] if style and style.get("border", 0) < len(self.borders) else {}
                for side, st in bd.items():
                    wpt = {"thin": 0.75, "medium": 1.5, "thick": 2.25, "double": 2.25}.get(st, 0.75)
                    css += f"border-{side}:{wpt}pt solid #000000;"
                h = style.get("h") if style else None
                if h in ("center", "centerContinuous"):
                    css += "text-align:center;"
                elif h == "right" or (h in (None, "general") and is_num):
                    css += "text-align:right;"
                span = merges.get((r, c))
                attrs = ""
                if span:
                    if span[0] > 1:
                        attrs += f' rowspan="{span[0]}"'
                    if span[1] > 1:
                        attrs += f' colspan="{span[1]}"'
                wd = widths.get(c, 48)
                if not span or span[1] == 1:
                    css += f"width:{wd:.0f}pt;"
                body = esc(text).replace("\n", "<br/>") or "&#160;"
                out.append(f'<td{attrs} style="{css}">{body}</td>')
            out.append("</tr>")
        out.append("</table>")
        return "".join(out), total_w

    def convert(self, paper="A4", show_grid=True):
        """Each visible sheet starts on a new page; wide sheets go landscape
        and are scaled down to fit the page width."""
        css = "body{font-size:9pt;} td{padding:1pt 3pt;} h3{font-size:12pt;}"
        out = fitz.open()
        margin = 36
        for name, path in self.sheets():
            h, wdt = self.sheet_html(path, name, show_grid)
            pw, ph = PAPER.get(paper, PAPER["A4"])
            if wdt > pw - 2 * margin:
                pw, ph = ph, pw
            out.insert_pdf(html_to_pdf_fit_width(h, pw, ph, margin, css, None, wdt))
        if out.page_count == 0:
            pw, ph = PAPER.get(paper, PAPER["A4"])
            out = html_to_pdf("<p>(empty workbook)</p>", pw, ph)
        return out


def xlsx_to_pdf(path, paper="A4", show_grid=True):
    return XlsxConverter(path).convert(paper, show_grid)


# --------------------------------------------------------------------------
# text, images, svg
# --------------------------------------------------------------------------
def txt_to_pdf(path, paper="A4"):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    pw, ph = PAPER.get(paper, PAPER["A4"])
    body = f'<div style="white-space:pre-wrap;font-family:monospace;font-size:10pt">{esc(text)}</div>'
    return html_to_pdf(body, pw, ph, (50, 50, 50, 50))


def page_size_for(content_w, content_h, page_mode, margin):
    """page_mode: 'fit' (page = picture size) or a PAPER name. Orientation follows the picture."""
    if page_mode == "fit":
        return content_w + 2 * margin, content_h + 2 * margin
    pw, ph = PAPER.get(page_mode, PAPER["A4"])
    if content_w > content_h:
        pw, ph = ph, pw
    return pw, ph


def place_rect(pw, ph, cw, ch, margin):
    avail_w, avail_h = pw - 2 * margin, ph - 2 * margin
    s = min(avail_w / cw, avail_h / ch)
    w_, h_ = cw * s, ch * s
    x = (pw - w_) / 2
    y = (ph - h_) / 2
    return fitz.Rect(x, y, x + w_, y + h_)


def image_to_pdf(path, page_mode="A4", margin=20, image_bytes=None):
    """image_bytes: optional already-corrected data (e.g. EXIF-rotated photo)."""
    data = image_bytes if image_bytes is not None else open(path, "rb").read()
    img_doc = fitz.open(stream=data, filetype=os.path.splitext(path)[1][1:].lower() or "png")
    out = fitz.open()
    for i in range(img_doc.page_count):   # multi-page TIFF / GIF frames
        src = img_doc[i]
        # the picture's size in points (uses its own dpi when known, else 96 dpi)
        cw, ch = src.rect.width, src.rect.height
        pw, ph = page_size_for(cw, ch, page_mode, margin)
        page = out.new_page(width=pw, height=ph)
        rect = place_rect(pw, ph, cw, ch, margin)
        pdf_bytes = img_doc.convert_to_pdf(i, i + 1)
        tmp = fitz.open("pdf", pdf_bytes)
        page.show_pdf_page(rect, tmp, 0)
    return out


def svg_to_pdf(path, page_mode="fit", margin=20):
    svg = fitz.open(path, filetype="svg")
    pdf = fitz.open("pdf", svg.convert_to_pdf())
    src = pdf[0]
    cw, ch = src.rect.width, src.rect.height
    pw, ph = page_size_for(cw, ch, page_mode, margin)
    out = fitz.open()
    page = out.new_page(width=pw, height=ph)
    page.show_pdf_page(place_rect(pw, ph, cw, ch, margin), pdf, 0)
    return out


# --------------------------------------------------------------------------
# optional: exact layout with an installed office suite
# --------------------------------------------------------------------------
def find_libreoffice():
    for n in ("soffice", "libreoffice"):
        p = shutil.which(n)
        if p:
            return p
    cands = ["/Applications/LibreOffice.app/Contents/MacOS/soffice",
             r"C:\Program Files\LibreOffice\program\soffice.exe",
             r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
    return next((c for c in cands if os.path.isfile(c)), None)


def has_ms_office():
    if sys.platform != "win32":
        return False
    try:
        import winreg
        for prog in ("Word.Application", "Excel.Application"):
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, prog))
        return True
    except OSError:
        return False


def convert_with_office(path):
    """Returns a fitz.Document or raises. Uses LibreOffice, or MS Office on Windows."""
    tmpdir = tempfile.mkdtemp(prefix="digonto_")
    try:
        out_pdf = os.path.join(tmpdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
        kw = {"stdout": subprocess.PIPE, "stderr": subprocess.PIPE, "timeout": 300}
        if sys.platform == "win32":
            kw["creationflags"] = 0x08000000
        lo = find_libreoffice()
        if lo:
            subprocess.run([lo, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", tmpdir, path], **kw)
        elif has_ms_office():
            src = os.path.abspath(path).replace("'", "''")
            dst = out_pdf.replace("'", "''")
            if kind_of(path) == "word":
                ps = (f"$a=New-Object -ComObject Word.Application;$a.Visible=$false;"
                      f"$d=$a.Documents.Open('{src}',$false,$true);$d.ExportAsFixedFormat('{dst}',17);"
                      f"$d.Close(0);$a.Quit()")
            else:
                ps = (f"$a=New-Object -ComObject Excel.Application;$a.Visible=$false;$a.DisplayAlerts=$false;"
                      f"$w=$a.Workbooks.Open('{src}',0,$true);$w.ExportAsFixedFormat(0,'{dst}');"
                      f"$w.Close($false);$a.Quit()")
            subprocess.run(["powershell", "-NoProfile", "-Command", ps], **kw)
        else:
            raise RuntimeError("no LibreOffice or Microsoft Office found")
        if not os.path.exists(out_pdf):
            raise RuntimeError("the office program did not create a PDF")
        with open(out_pdf, "rb") as f:
            return fitz.open(stream=f.read(), filetype="pdf")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def office_available():
    return bool(find_libreoffice() or has_ms_office())


# --------------------------------------------------------------------------
def convert_file(path, page_mode="A4", margin=20, use_office=False, excel_grid=True, image_bytes=None):
    """Convert one file to a fitz.Document."""
    k = kind_of(path)
    paper = page_mode if page_mode in PAPER else "A4"
    if k in ("word", "excel") and use_office and office_available():
        return convert_with_office(path)
    if k == "image":
        return image_to_pdf(path, page_mode, margin, image_bytes)
    if k == "svg":
        return svg_to_pdf(path, page_mode, margin)
    if k == "word":
        return docx_to_pdf(path)
    if k == "excel":
        return xlsx_to_pdf(path, paper, excel_grid)
    if k == "text":
        return txt_to_pdf(path, paper)
    if k == "pdf":
        with open(path, "rb") as f:
            return fitz.open(stream=f.read(), filetype="pdf")
    if os.path.splitext(path)[1].lower() in (".doc", ".xls"):
        raise RuntimeError("old .doc/.xls format: open it in Word/Excel/LibreOffice and save as .docx/.xlsx")
    raise RuntimeError("file type not supported")


if __name__ == "__main__":   # quick command line use: python digonto_convert.py in.docx out.pdf
    d = convert_file(sys.argv[1])
    d.save(sys.argv[2], garbage=3, deflate=True)
    print(f"{sys.argv[2]}: {d.page_count} page(s)")
