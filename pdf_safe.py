"""Digonto PDF - PDFs that offices accept.

Government upload portals (Arbeitsagentur, Finanzamt, ...) convert every upload to PDF/A,
the archive format. Two things break that conversion, and Digonto used to add both:
  * text in a font that is not inside the file (the built-in "Helvetica")
  * pictures with a see-through background (a signature on a soft mask)
So everything Digonto writes now uses fonts that travel inside the file, and signatures are
drawn as solid ink shapes instead of see-through pictures.
`upload_copy` goes further for a file you are about to upload: forms and notes are flattened,
scripts, attachments, bookmarks and tags removed, and any page that would still break the
conversion is turned into a clean picture page.
"""
import os

try:
    import pymupdf as fitz
except ImportError:
    import fitz
import numpy as np

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(HERE, "assets", "fonts")

# the PDF built-in names used in the code -> a free font that is embedded instead
_FONTS = {
    "helv": ("DgSans", "NotoSans-Regular.ttf"),
    "hebo": ("DgSansBold", "NotoSans-Bold.ttf"),
    "tiro": ("DgSerif", "DejaVuSerif.ttf"),
    "tibo": ("DgSerifBold", "DejaVuSerif-Bold.ttf"),
    "cour": ("DgMono", "FiraMono-Regular.ttf"),
    "cobo": ("DgMonoBold", "FiraMono-Bold.ttf"),
}
_cache = {}


def font_kw(name):
    """keyword arguments for page.insert_text / insert_textbox: an embedded font instead of a built-in one"""
    nm, fn = _FONTS.get(name, (None, None))
    path = os.path.join(FONT_DIR, fn) if fn else None
    if path and os.path.exists(path):
        return {"fontname": nm, "fontfile": path}
    return {"fontname": name}


def _font(name):
    if name not in _cache:
        kw = font_kw(name)
        _cache[name] = fitz.Font(fontfile=kw["fontfile"]) if "fontfile" in kw else fitz.Font(name)
    return _cache[name]


def text_length(text, name="helv", size=11):
    return _font(name).text_length(text, fontsize=size)


def finish(doc):
    """before saving: keep only the letters really used of the embedded fonts (small files)"""
    try:
        doc.subset_fonts()
    except Exception:
        pass


# ---------------------------------------------------------------- signatures
def place_signature(page, rect_displayed, png):
    """rect_displayed: (x, y, w, h) in page coordinates as you see the page (rotation applied).
    A signature with a see-through background is drawn as solid ink (vector shapes), which every
    office system accepts; a picture without see-through parts is placed as it is."""
    x, y, w, h = rect_displayed
    pix = fitz.Pixmap(png)
    if not pix.alpha:
        r = fitz.Rect(x, y, x + w, y + h) * page.derotation_matrix
        page.insert_image(r, stream=png, rotate=page.rotation, keep_proportion=False, overlay=True)
        return
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)
    # about 4 dots per point (~290 dpi) is plenty for a signature, and keeps the file small
    step = max(1, int(min(pix.width / max(w * 4, 1), pix.height / max(h * 4, 1))))
    a = a[::step, ::step]
    ink = a[..., 3] >= 128
    if not ink.any():
        return
    rgb = a[..., :3][ink].astype(np.float32).mean(axis=0) / 255.0
    rows, cols = ink.shape
    sx, sy = w / cols, h / rows
    der = page.derotation_matrix
    sh = page.new_shape()
    edge = np.diff(np.pad(ink.astype(np.int8), ((0, 0), (1, 1))), axis=1)
    for r_ in range(rows):
        starts = np.flatnonzero(edge[r_] == 1)
        ends = np.flatnonzero(edge[r_] == -1)
        for c0, c1 in zip(starts, ends):
            q = fitz.Rect(float(x + c0 * sx), float(y + r_ * sy), float(x + c1 * sx), float(y + (r_ + 1.05) * sy)).quad * der
            sh.draw_quad(q)
    sh.finish(color=None, fill=tuple(float(v) for v in rgb), width=0, closePath=True)
    sh.commit(overlay=True)


# ---------------------------------------------------------------- upload copy
def page_problems(doc, pno):
    """what on this page would stop an office system from accepting it"""
    out = []
    page = doc[pno]
    for f in page.get_fonts(full=True):
        if f[1] in ("n/a", "") and f[2] not in ("Type3",):
            out.append(f"font {f[3]} is not inside the file")
            break
    for im in page.get_images(full=True):
        if im[1]:                                   # has a soft mask: see-through picture
            out.append("see-through picture")
            break
    return out


def upload_copy(data, as_picture=False, dpi=200, password=None):
    """make a copy that government upload portals accept. Returns (pdf bytes, pages turned into pictures)"""
    doc = fitz.open(stream=data, filetype="pdf")
    if doc.needs_pass:
        if not password or not doc.authenticate(password):
            raise RuntimeError("This PDF has a password. Remove it first (Unlock).")
    try:
        doc.bake(annots=True, widgets=True)          # form fields and notes become normal page content
    except Exception:
        pass
    try:
        doc.scrub(attached_files=True, clean_pages=False, embedded_files=True, hidden_text=False, javascript=True,
                  metadata=False, redactions=False, remove_links=False, reset_fields=False, reset_responses=False,
                  thumbnails=True, xml_metadata=True)
    except Exception:
        pass
    try:
        doc.set_toc([])                              # no bookmarks
    except Exception:
        pass
    cat = doc.pdf_catalog()
    for key in ("AcroForm", "StructTreeRoot", "MarkInfo", "OpenAction", "AA", "Names", "OCProperties", "Outlines"):
        try:
            doc.xref_set_key(cat, key, "null")
        except Exception:
            pass
    finish(doc)
    out = fitz.open()
    pictures = []
    for i in range(doc.page_count):
        if as_picture or page_problems(doc, i):
            src = doc[i]
            pg = out.new_page(width=src.rect.width, height=src.rect.height)
            pix = src.get_pixmap(dpi=dpi, alpha=False)
            pg.insert_image(pg.rect, stream=pix.tobytes("jpeg", jpg_quality=90))
            pictures.append(i + 1)
        else:
            out.insert_pdf(doc, from_page=i, to_page=i, annots=False, links=False)
    meta = {k: v for k, v in (doc.metadata or {}).items() if k in ("title", "author", "subject", "keywords")}
    meta.update(producer="Digonto PDF", creator="Digonto PDF")
    out.set_metadata(meta)
    res = out.tobytes(garbage=4, deflate=True, clean=True, use_objstms=0)
    if res.startswith(b"%PDF-1.7"):                  # nothing newer than PDF 1.4 is used: say so
        res = b"%PDF-1.4" + res[8:]
    return res, pictures
