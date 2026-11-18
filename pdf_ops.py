"""Digonto PDF tools as plain functions.

Every tool works on a PyMuPDF document and knows nothing about windows, so the same code runs for the open
file and for many files at once (batch).  A tool is described by a Tool record: its name, group, a symbol for
the list, the options it asks for and what it gives back:
    kind "doc"     changes the document (rotate, OCR, compress ...)
    kind "files"   makes new files and leaves the document as it is (split, PDF to images ...)
    kind "ui"      needs the page on screen (sign, add text, redact by drawing ...) - handled by the window
"""
import difflib
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

try:
    import pymupdf as fitz
except ImportError:                                      # older PyMuPDF
    import fitz

import pdf_scan as scan


class Stopped(Exception):
    pass


# ---------------------------------------------------------------------------------------------- helpers --
def parse_ranges(text, count):
    """'1-3, 5, 8-' -> [0, 1, 2, 4, 7, 8 ...] (0-based, in the given order, no doubles)"""
    out = []
    for part in re.split(r"[,;\s]+", (text or "").strip()):
        if not part:
            continue
        m = re.fullmatch(r"(\d*)\s*-\s*(\d*)", part)
        if m:
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else count
        elif part.isdigit():
            a = b = int(part)
        else:
            raise ValueError(f"'{part}' is not a page or a range like 2-5")
        step = 1 if b >= a else -1
        for p in range(a, b + step, step):
            if 1 <= p <= count and p - 1 not in out:
                out.append(p - 1)
    return out


def target_pages(doc, opts):
    """the pages a tool should work on: 'all', the selected ones, or typed ranges"""
    which = opts.get("which", "all")
    if which == "selected" and opts.get("_selected"):
        return [p for p in opts["_selected"] if p < doc.page_count]
    if which == "ranges":
        return parse_ranges(opts.get("ranges", ""), doc.page_count)
    return list(range(doc.page_count))


def _tick(progress, stop, frac, text=""):
    if stop and stop():
        raise Stopped()
    if progress:
        progress(frac, text)


def mm(v):
    return float(v) * 72.0 / 25.4


def human(n):
    for unit in ("bytes", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:.0f} {unit}" if unit == "bytes" else f"{n:.1f} {unit}"
        n /= 1024.0


# ------------------------------------------------------------------------------------------------ tools --
def op_rotate(doc, o, progress=None, stop=None):
    deg = {"Right (90°)": 90, "Upside down (180°)": 180, "Left (90°)": 270}[o.get("angle", "Right (90°)")]
    pages = target_pages(doc, o)
    for k, i in enumerate(pages):
        p = doc[i]
        p.set_rotation((p.rotation + deg) % 360)
        _tick(progress, stop, (k + 1) / len(pages))
    return f"Turned {len(pages)} page(s)"


def op_remove_pages(doc, o, progress=None, stop=None):
    pages = target_pages(doc, o)
    if len(pages) >= doc.page_count:
        raise ValueError("A PDF needs at least one page - not every page can be removed.")
    doc.delete_pages(sorted(pages))
    return f"Removed {len(pages)} page(s)"


def op_page_numbers(doc, o, progress=None, stop=None):
    pos = o.get("where", "Bottom centre")
    look = o.get("look", "1 / N")
    first = int(o.get("start", 1))
    size = float(o.get("size", 10))
    n = doc.page_count
    for i in range(n):
        p = doc[i]
        k = i + first
        last = n + first - 1
        txt = {"1": f"{k}", "Page 1": f"Page {k}", "1 / N": f"{k} / {last}", "Page 1 of N": f"Page {k} of {last}"}[look]
        r = p.rect                                    # displayed (rotated) size
        w = fitz.get_text_length(txt, fontname="helv", fontsize=size)
        x = {"centre": (r.width - w) / 2, "right": r.width - w - 36, "left": 36}[pos.split()[1]]
        y = r.height - 24 if pos.startswith("Bottom") else 24 + size
        p.insert_text(fitz.Point(x, y) * p.derotation_matrix, txt, fontsize=size, fontname="helv",
                      color=(0.2, 0.2, 0.2), rotate=p.rotation)
        _tick(progress, stop, (i + 1) / n)
    return f"Numbered {n} pages"


def content_box(page):
    """the part of the page that is not white (for auto crop)"""
    a = scan.page_gray(page, dpi=60)
    ys, xs = (a < 235).nonzero()
    if len(xs) == 0:
        return None
    sx = page.rect.width / a.shape[1]
    sy = page.rect.height / a.shape[0]
    pad = 6
    return fitz.Rect(xs.min() * sx - pad, ys.min() * sy - pad, (xs.max() + 1) * sx + pad, (ys.max() + 1) * sy + pad)


def op_crop(doc, o, progress=None, stop=None):
    pages = target_pages(doc, o)
    auto = o.get("mode", "Cut white edges") == "Cut white edges"
    m = mm(o.get("margin", 10))
    done = 0
    for k, i in enumerate(pages):
        p = doc[i]
        disp = p.rect
        if auto:
            box = content_box(p)
            if box is None:
                continue
            box &= disp
        else:
            box = fitz.Rect(disp.x0 + m, disp.y0 + m, disp.x1 - m, disp.y1 - m)
        if box.width < 20 or box.height < 20:
            continue
        unrot = box * p.derotation_matrix
        unrot.normalize()
        unrot &= p.mediabox
        p.set_cropbox(unrot)
        done += 1
        _tick(progress, stop, (k + 1) / len(pages))
    return f"Cropped {done} page(s)"


def op_ocr(doc, o, progress=None, stop=None):
    langs = tuple(o.get("languages") or ("eng",))
    skip = o.get("skip_text", True)
    pages = target_pages(doc, o)
    words = done = 0
    for k, i in enumerate(pages):
        _tick(progress, stop, k / len(pages), f"Reading page {i + 1} of {doc.page_count}")
        p = doc[i]
        if skip and scan.has_text(p):
            continue
        words += scan.ocr_page(p, langs)
        done += 1
    _tick(progress, stop, 1.0)
    if done == 0:
        return "Every page already had text - nothing to do"
    return f"Made {done} page(s) searchable ({words} words found)"


def op_clean(doc, o, progress=None, stop=None):
    pages = target_pages(doc, o)
    opts = dict(straighten=o.get("straighten", True), specks=o.get("specks", True), darker=o.get("darker", True),
                white=o.get("white", True), edges=o.get("edges", True), strength=float(o.get("strength", 60)) / 100)
    for k, i in enumerate(pages):
        _tick(progress, stop, k / len(pages), f"Cleaning page {i + 1} of {doc.page_count}")
        scan.clean_page(doc, i, dpi=int(o.get("dpi", 200)), **opts)
    _tick(progress, stop, 1.0)
    return f"Cleaned {len(pages)} page(s)"


def op_remove_blank(doc, o, progress=None, stop=None):
    limit = {"Only empty pages": 0.0008, "Normal": 0.003, "Also nearly empty pages": 0.01}[o.get("sense", "Normal")]
    blank = []
    for i in range(doc.page_count):
        _tick(progress, stop, i / doc.page_count, f"Checking page {i + 1}")
        if scan.is_blank(doc[i], limit):
            blank.append(i)
    if len(blank) == doc.page_count:
        raise ValueError("Every page looks blank - nothing was removed.")
    if blank:
        doc.delete_pages(blank)
    return f"Removed {len(blank)} blank page(s)" if blank else "No blank pages found"


PRESETS = [
    # label, ghostscript setting, image dpi, jpeg quality
    ("Smallest - screen, WhatsApp, email", "screen", 72, 45),
    ("Small - good quality (recommended)", "ebook", 150, 70),
    ("Larger - for printing", "printer", 300, 85),
]


def find_ghostscript():
    for name in ("gs", "gswin64c", "gswin32c"):
        p = shutil.which(name)
        if p:
            return p
    for c in ("/opt/homebrew/bin/gs", "/usr/local/bin/gs", "/usr/bin/gs"):
        if os.path.isfile(c):
            return c
    return None


def shrink_bytes(data, preset_label):
    """the smallest of: PyMuPDF rewrite, Ghostscript (when installed). Returns bytes (maybe the original)."""
    _, gs_setting, dpi, quality = next(p for p in PRESETS if p[0] == preset_label)
    results = []
    d = fitz.open(stream=data, filetype="pdf")
    if hasattr(d, "rewrite_images"):
        d.rewrite_images(dpi_threshold=int(dpi * 1.25), dpi_target=dpi, quality=quality,
                         lossy=True, lossless=True, bitonal=False, color=True, gray=True)
    try:
        d.subset_fonts()
    except Exception:                                  # noqa: BLE001
        pass
    results.append(d.tobytes(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True,
                             use_objstms=1))
    gs = find_ghostscript()
    if gs:
        tmp = tempfile.mkdtemp(prefix="digonto_")
        try:
            src, dst = os.path.join(tmp, "a.pdf"), os.path.join(tmp, "b.pdf")
            with open(src, "wb") as f:
                f.write(data)
            kw = {"creationflags": 0x08000000} if sys.platform == "win32" else {}
            r = subprocess.run([gs, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.5", f"-dPDFSETTINGS=/{gs_setting}",
                                "-dNOPAUSE", "-dBATCH", "-dQUIET", f"-sOutputFile={dst}", src],
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=600, **kw)
            if r.returncode == 0 and os.path.getsize(dst) > 0:
                with open(dst, "rb") as f:
                    results.append(f.read())
        except Exception:                              # noqa: BLE001 - Ghostscript is only a bonus
            pass
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
    best = min(results, key=len)
    return best if len(best) < len(data) else data


def op_compress(doc, o, progress=None, stop=None):
    """special: replaces the whole document - returns (message, new bytes)"""
    _tick(progress, stop, 0.1, "Making it smaller ...")
    data = doc.tobytes()
    out = shrink_bytes(data, o.get("preset", PRESETS[1][0]))
    _tick(progress, stop, 1.0)
    if len(out) >= len(data):
        return "This PDF is already as small as it gets", None
    return f"{human(len(data))} → {human(len(out))}  (-{100 - len(out) * 100 // max(len(data), 1)} %)", out


def op_watermark(doc, o, progress=None, stop=None):
    txt = (o.get("text") or "").strip()
    if not txt:
        raise ValueError("Type the watermark text first.")
    opacity = float(o.get("opacity", 30)) / 100
    angle = {"Diagonal": 45, "Straight": 0}[o.get("angle", "Diagonal")]
    pages = target_pages(doc, o)
    for k, i in enumerate(pages):
        p = doc[i]
        r = p.rect
        size = min(r.width, r.height) / max(len(txt), 3) * (1.5 if angle else 1.2)
        w = fitz.get_text_length(txt, fontname="hebo", fontsize=size)
        c = fitz.Point(r.width / 2, r.height / 2)
        start = fitz.Point(c.x - w / 2, c.y + size / 3)
        p.insert_text(start * p.derotation_matrix, txt, fontsize=size, fontname="hebo", color=(0.55, 0.55, 0.55),
                      fill_opacity=opacity, morph=(c * p.derotation_matrix, fitz.Matrix(angle - p.rotation)),
                      overlay=True)
        _tick(progress, stop, (k + 1) / len(pages))
    return f"Watermark on {len(pages)} page(s)"


def op_redact_words(doc, o, progress=None, stop=None):
    """black out every place where the words appear - the text is really removed, not just covered"""
    words = [w.strip() for w in re.split(r"[\n,;]+", o.get("words", "")) if w.strip()]
    if not words:
        raise ValueError("Type the word(s) to black out first.")
    hits = 0
    for i in range(doc.page_count):
        p = doc[i]
        for w in words:
            for r in p.search_for(w):
                p.add_redact_annot(r, fill=(0, 0, 0))
                hits += 1
        p.apply_redactions(images=fitz.PDF_REDACT_IMAGE_PIXELS if hasattr(fitz, "PDF_REDACT_IMAGE_PIXELS") else 2)
        _tick(progress, stop, (i + 1) / doc.page_count)
    return f"Blacked out {hits} place(s)" if hits else "The words were not found"


def op_info(doc, o, progress=None, stop=None):
    m = dict(doc.metadata or {})
    for k in ("title", "author", "subject", "keywords"):
        if k in o:
            m[k] = o[k]
    doc.set_metadata({k: v for k, v in m.items() if k in ("title", "author", "subject", "keywords", "creator",
                                                          "producer", "creationDate", "modDate")})
    return "Document info saved"


def op_repair(doc, o, progress=None, stop=None):
    _tick(progress, stop, 0.3, "Rebuilding the file ...")
    data = doc.tobytes(garbage=4, deflate=True, clean=True)
    fixed = fitz.open(stream=data, filetype="pdf")
    _tick(progress, stop, 1.0)
    return f"Repaired - {fixed.page_count} pages are readable", data


# --- tools that make files ------------------------------------------------------------------------------
def files_split(doc, o, base, progress=None, stop=None):
    n = doc.page_count
    mode = o.get("mode", "Every page on its own")
    groups = []
    if mode == "Every page on its own":
        groups = [[i] for i in range(n)]
    elif mode == "Every N pages":
        step = max(1, int(o.get("every", 2)))
        groups = [list(range(i, min(i + step, n))) for i in range(0, n, step)]
    else:
        for part in re.split(r"[;\n]+|,(?=\s*\d+\s*-)", o.get("ranges", "")):
            if part.strip():
                groups.append(parse_ranges(part, n))
        groups = [g for g in groups if g]
        if not groups:
            raise ValueError("Type the parts, for example: 1-3; 4-9; 10-")
    out = []
    width = len(str(len(groups)))
    for k, g in enumerate(groups):
        part = fitz.open()
        for i in g:
            part.insert_pdf(doc, from_page=i, to_page=i)
        out.append((f"{base} - part {str(k + 1).zfill(width)}.pdf", part.tobytes(garbage=3, deflate=True)))
        _tick(progress, stop, (k + 1) / len(groups))
    return out


def files_extract_pages(doc, o, base, progress=None, stop=None):
    pages = target_pages(doc, o)
    part = fitz.open()
    for i in pages:
        part.insert_pdf(doc, from_page=i, to_page=i)
    return [(f"{base} - pages.pdf", part.tobytes(garbage=3, deflate=True))]


def files_images(doc, o, base, progress=None, stop=None):
    fmt = o.get("format", "PNG").lower()
    dpi = int(o.get("dpi", 150))
    pages = target_pages(doc, o)
    width = len(str(doc.page_count))
    out = []
    for k, i in enumerate(pages):
        pix = doc[i].get_pixmap(dpi=dpi, alpha=False)
        data = pix.tobytes("jpeg", jpg_quality=88) if fmt == "jpg" else pix.tobytes("png")
        out.append((f"{base} - page {str(i + 1).zfill(width)}.{fmt}", data))
        _tick(progress, stop, (k + 1) / len(pages), f"Page {i + 1}")
    return out


def files_text(doc, o, base, progress=None, stop=None):
    parts = []
    for i in range(doc.page_count):
        parts.append(doc[i].get_text("text", sort=True))
        _tick(progress, stop, (i + 1) / doc.page_count)
    text = "\n\f\n".join(parts)
    if not text.strip():
        raise ValueError("No text found. If this is a scan, use OCR first.")
    return [(f"{base}.txt", text.encode("utf-8"))]


def files_extract_images(doc, o, base, progress=None, stop=None):
    seen, out = set(), []
    min_px = int(o.get("min", 64))
    for i in range(doc.page_count):
        for img in doc[i].get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                info = doc.extract_image(xref)
            except Exception:                          # noqa: BLE001
                continue
            if not info or min(info.get("width", 0), info.get("height", 0)) < min_px:
                continue
            ext = info.get("ext", "png")
            out.append((f"{base} - image {len(out) + 1} (page {i + 1}).{ext}", info["image"]))
        _tick(progress, stop, (i + 1) / doc.page_count)
    if not out:
        raise ValueError("No pictures found in this PDF.")
    return out


def files_word(doc, o, base, progress=None, stop=None):
    return [(f"{base}.docx", pdf_to_docx(doc, progress, stop))]


# --- PDF -> Word (Digonto's own, small .docx writer) ------------------------------------------------------
def _run_xml(text, size, bold, italic, color):
    props = []
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if color and color != "000000":
        props.append(f'<w:color w:val="{color}"/>')
    props.append(f'<w:sz w:val="{max(8, int(round(size * 2)))}"/>')
    return (f'<w:r><w:rPr>{"".join(props)}</w:rPr>'
            f'<w:t xml:space="preserve">{escape(text)}</w:t></w:r>')


def pdf_to_docx(doc, progress=None, stop=None):
    """text with sizes, bold/italic and colours, paragraphs and page breaks, plus the pictures of each page"""
    body, media, rels = [], [], []
    for i in range(doc.page_count):
        page = doc[i]
        d = page.get_text("dict", sort=True)
        for b in d["blocks"]:
            if b.get("type") == 1 and b.get("image"):
                try:
                    pix = fitz.Pixmap(b["image"])
                    if pix.alpha or pix.n > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    png = pix.tobytes("png")
                except Exception:                      # noqa: BLE001
                    continue
                k = len(media) + 1
                media.append((f"image{k}.png", png))
                rid = f"rIdImg{k}"
                rels.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                            f'relationships/image" Target="media/image{k}.png"/>')
                bw = min(b["bbox"][2] - b["bbox"][0], 470)
                bh = (b["bbox"][3] - b["bbox"][1]) * bw / max(b["bbox"][2] - b["bbox"][0], 1)
                cx, cy = int(bw * 12700), int(bh * 12700)
                body.append(
                    f'<w:p><w:r><w:drawing><wp:inline><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{k}" name="p{k}"/>'
                    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                    '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                    f'<pic:nvPicPr><pic:cNvPr id="{k}" name="p{k}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
                    f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
                    f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
                    '</wp:inline></w:drawing></w:r></w:p>')
                continue
            if b.get("type") != 0:
                continue
            runs = []
            for ln in b["lines"]:
                for sp in ln["spans"]:
                    t = sp["text"]
                    if not t:
                        continue
                    f = sp.get("flags", 0)
                    bold = bool(f & 16) or "bold" in sp.get("font", "").lower()
                    italic = bool(f & 2) or "italic" in sp.get("font", "").lower()
                    runs.append(_run_xml(t, sp["size"], bold, italic, f'{sp.get("color", 0):06X}'))
                if runs:
                    runs.append(_run_xml(" ", ln["spans"][-1]["size"] if ln["spans"] else 11, False, False, None))
            if runs:
                body.append(f'<w:p><w:pPr><w:spacing w:after="120"/></w:pPr>{"".join(runs)}</w:p>')
        if i < doc.page_count - 1:
            body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        _tick(progress, stop, (i + 1) / doc.page_count, f"Page {i + 1}")
    w = doc[0].rect.width if doc.page_count else 595
    h = doc[0].rect.height if doc.page_count else 842
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
        f'<w:body>{"".join(body)}<w:sectPr><w:pgSz w:w="{int(w * 20)}" w:h="{int(h * 20)}"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/'
                   '2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-'
                   'package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" '
                   'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/'
                   'package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                   'officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/'
                   f'package/2006/relationships">{"".join(rels)}</Relationships>')
        z.writestr("word/document.xml", document)
        for name, data in media:
            z.writestr("word/media/" + name, data)
    return buf.getvalue()


# --- compare ------------------------------------------------------------------------------------------------
def compare_html(doc_a, doc_b, name_a, name_b):
    """the text of two PDFs side by side, page by page, changes marked"""
    rows = []
    changed = 0
    for i in range(max(doc_a.page_count, doc_b.page_count)):
        ta = doc_a[i].get_text("text", sort=True).splitlines() if i < doc_a.page_count else []
        tb = doc_b[i].get_text("text", sort=True).splitlines() if i < doc_b.page_count else []
        if ta == tb:
            continue
        changed += 1
        rows.append(f'<tr><td colspan="2" class="pg">Page {i + 1}</td></tr>')
        for tag, a0, a1, b0, b1 in difflib.SequenceMatcher(None, ta, tb, autojunk=False).get_opcodes():
            if tag == "equal":
                continue
            left = "<br>".join(escape(x) for x in ta[a0:a1]) or "&nbsp;"
            right = "<br>".join(escape(x) for x in tb[b0:b1]) or "&nbsp;"
            rows.append(f'<tr><td class="old">{left}</td><td class="new">{right}</td></tr>')
    if not changed:
        return 0, "<p style='font-size:15px'>The text of both PDFs is the same.</p>"
    head = (f"<style>td{{vertical-align:top;padding:5px 8px;font-size:12px}} .pg{{background:#EDE6F4;font-weight:700}}"
            f" .old{{background:#FBE3E8}} .new{{background:#E2F3E6}}</style>"
            f"<table width='100%' cellspacing='2'><tr><th>{escape(name_a)}</th><th>{escape(name_b)}</th></tr>")
    return changed, head + "".join(rows) + "</table>"


# ------------------------------------------------------------------------------------------- tool list --
@dataclass
class Tool:
    key: str
    name: str
    group: str
    symbol: str                      # icon name (digonto_ui.line_icon)
    tip: str
    kind: str = "doc"                # doc | files | ui | action
    fields: list = field(default_factory=list)
    run: object = None
    button: str = "Apply"
    batch: bool = True               # can run on many files at once
    preview: bool = False            # before/after slider


WHICH = ("choice", "which", "Pages", ["all", "selected", "ranges"], "all")
RANGES = ("text", "ranges", "Page numbers (when 'ranges')", "1-3, 5")
LANGS = [code for _, code in scan.LANGUAGES]

TOOLS = [
    # Organise
    Tool("merge", "Merge", "Organise", "merge", "Put several PDFs together into one", "action", button="Choose PDFs …",
         batch=True),
    Tool("split", "Split", "Organise", "split", "Cut a PDF into several files", "files",
         [("choice", "mode", "How", ["Every page on its own", "Every N pages", "By ranges"], "Every page on its own"),
          ("int", "every", "N (pages per file)", 1, 500, 2),
          ("text", "ranges", "Parts (when 'by ranges')", "1-3; 4-9; 10-")], files_split, "Split"),
    Tool("rotate", "Rotate", "Organise", "rotate", "Turn pages", "doc",
         [("choice", "angle", "Turn", ["Right (90°)", "Left (90°)", "Upside down (180°)"], "Right (90°)"), WHICH, RANGES],
         op_rotate, "Rotate"),
    Tool("organise", "Organise pages", "Organise", "organise", "Move, copy, insert and reverse pages", "ui", batch=False),
    Tool("remove", "Remove pages", "Organise", "remove", "Delete the pages you don't need", "doc",
         [("choice", "which", "Pages", ["selected", "ranges"], "selected"), RANGES], op_remove_pages, "Remove",
         batch=True),
    Tool("extract", "Extract pages", "Organise", "extract", "Save some pages as a new PDF", "files",
         [("choice", "which", "Pages", ["selected", "ranges"], "selected"), RANGES], files_extract_pages, "Extract",
         batch=False),
    Tool("numbers", "Page numbers", "Organise", "numbers", "Number every page", "doc",
         [("choice", "where", "Where", ["Bottom centre", "Bottom right", "Bottom left", "Top centre", "Top right"],
           "Bottom centre"),
          ("choice", "look", "Look", ["1", "Page 1", "1 / N", "Page 1 of N"], "1 / N"),
          ("int", "start", "First number", 0, 9999, 1), ("int", "size", "Size", 6, 30, 10)], op_page_numbers, "Add numbers"),
    Tool("crop", "Crop", "Organise", "crop", "Cut away white edges or a margin", "doc",
         [("choice", "mode", "How", ["Cut white edges", "Same margin all round"], "Cut white edges"),
          ("int", "margin", "Margin (mm)", 1, 100, 10), WHICH, RANGES], op_crop, "Crop"),
    # Scans
    Tool("ocr", "OCR (make searchable)", "Scans", "ocr", "Read the text of scanned pages so you can search and copy it",
         "doc", [("langs", "languages", "Languages", LANGS, ["eng"]),
                 ("check", "skip_text", "Skip pages that already have text", True), WHICH, RANGES], op_ocr, "Run OCR"),
    Tool("clean", "Clean up scans", "Scans", "clean", "Straighten, remove dots, white paper, dark edges away", "doc",
         [("check", "straighten", "Straighten crooked pages", True), ("check", "specks", "Remove dots and dust", True),
          ("check", "white", "Make the paper white (no shadows)", True), ("check", "darker", "Darker, sharper text", True),
          ("check", "edges", "Remove dark scanner edges", True), ("slider", "strength", "Strength", 0, 100, 60),
          WHICH, RANGES], op_clean, "Clean up", preview=True),
    Tool("blank", "Remove blank pages", "Scans", "blank", "Find and delete empty pages (also empty backs of scans)", "doc",
         [("choice", "sense", "Which pages count as blank", ["Only empty pages", "Normal", "Also nearly empty pages"],
           "Normal")], op_remove_blank, "Remove blank pages"),
    Tool("compress", "Compress", "Scans", "compress", "Make the file smaller", "doc",
         [("choice", "preset", "Size", [p[0] for p in PRESETS], PRESETS[1][0])], op_compress, "Compress",
         preview=True),
    # Convert
    Tool("img2pdf", "Images to PDF", "Convert", "img2pdf", "Pictures, Word, Excel or text files into a PDF", "action",
         button="Choose files …", batch=False),
    Tool("pdf2img", "PDF to images", "Convert", "images", "Every page as a picture", "files",
         [("choice", "format", "Format", ["PNG", "JPG"], "PNG"), ("int", "dpi", "Quality (dpi)", 50, 600, 150),
          WHICH, RANGES], files_images, "Save pictures"),
    Tool("pdf2word", "PDF to Word", "Convert", "word", "An editable .docx file", "files", [], files_word, "Save as Word"),
    Tool("pdf2text", "PDF to text", "Convert", "textfile", "All text as a .txt file", "files", [], files_text, "Save as text"),
    # Edit & sign
    Tool("text", "Add or edit text", "Edit & sign", "text", "Change, add or delete text on the page", "ui", batch=False),
    Tool("sign", "Sign", "Edit & sign", "sign", "Put your signature on a page", "ui", batch=False),
    Tool("watermark", "Watermark", "Edit & sign", "watermark", "Text across every page", "doc",
         [("text", "text", "Text", "DRAFT"), ("slider", "opacity", "See-through", 5, 90, 30),
          ("choice", "angle", "Direction", ["Diagonal", "Straight"], "Diagonal"), WHICH, RANGES], op_watermark,
         "Add watermark"),
    Tool("redact", "Redact (black out)", "Edit & sign", "redact", "Remove words for good - draw boxes or search", "ui",
         [("text", "words", "Words to black out everywhere (comma between)", "")], op_redact_words, "Black out words"),
    # Security
    Tool("lock", "Add password", "Security", "lock", "Nobody can open the PDF without the password", "action",
         [("password", "pw", "Password", ""), ("password", "pw2", "Again", "")], button="Set password"),
    Tool("unlock", "Remove password", "Security", "unlock", "Save the PDF without a password", "action",
         button="Remove password"),
    # More
    Tool("compare", "Compare", "More", "compare", "Show what changed between two PDFs", "action", button="Choose the other PDF …",
         batch=False),
    Tool("images", "Extract images", "More", "images", "Save every picture inside the PDF", "files",
         [("int", "min", "Skip pictures smaller than (pixels)", 1, 2000, 64)], files_extract_images, "Save pictures"),
    Tool("info", "Document info", "More", "info", "Title, author, pages, size", "action", button="Save info", batch=False),
    Tool("repair", "Repair", "More", "repair", "Fix a damaged PDF", "doc", [], op_repair, "Repair"),
]
TOOL = {t.key: t for t in TOOLS}
GROUPS = ["Organise", "Scans", "Convert", "Edit & sign", "Security", "More"]


def run_doc_tool(tool, data, opts, progress=None, stop=None, password=None):
    """run a 'doc' tool on PDF bytes; returns (message, new bytes or None when nothing changed)"""
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.needs_pass and not (password and doc.authenticate(password)):
        raise ValueError("This PDF has a password.")
    r = tool.run(doc, opts, progress, stop)
    if isinstance(r, tuple):
        return r
    return r, doc.tobytes(garbage=3, deflate=True)


def run_files_tool(tool, data, opts, base, progress=None, stop=None, password=None):
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.needs_pass and not (password and doc.authenticate(password)):
        raise ValueError("This PDF has a password.")
    return tool.run(doc, opts, base, progress, stop)
