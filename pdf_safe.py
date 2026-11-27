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
def _stencil_page(ink, rgb, w_pt, h_pt):
    """a one-page PDF holding the signature as a 1-bit stencil image (an /ImageMask) painted in its ink colour:
    one small image, no see-through picture, allowed in every PDF/A level"""
    rows, cols = ink.shape
    tmp = fitz.open()
    tp = tmp.new_page(width=w_pt, height=h_pt)
    x = tmp.get_new_xref()
    tmp.update_object(x, f"<< /Type /XObject /Subtype /Image /Width {cols} /Height {rows} /ImageMask true /BitsPerComponent 1 >>")
    tmp.update_stream(x, np.packbits(~ink, axis=1).tobytes())          # 0 = ink, 1 = leave the page as it is
    tmp.xref_set_key(tp.xref, "Resources", f"<< /XObject << /Sig {x} 0 R >> >>")
    c = tmp.get_new_xref()
    tmp.update_object(c, "<< >>")
    tmp.update_stream(c, f"q {rgb[0]:.4f} {rgb[1]:.4f} {rgb[2]:.4f} rg {w_pt:.4f} 0 0 {h_pt:.4f} 0 0 cm /Sig Do Q".encode())
    tmp.xref_set_key(tp.xref, "Contents", f"{c} 0 R")
    return tmp


def place_signature(page, rect_displayed, png):
    """rect_displayed: (x, y, w, h) in page coordinates as you see the page (rotation applied).
    A signature with a see-through background becomes one small stencil image in its ink colour, which office
    systems accept; a picture without see-through parts is placed as it is."""
    x, y, w, h = rect_displayed
    pix = fitz.Pixmap(png)
    r = fitz.Rect(x, y, x + w, y + h) * page.derotation_matrix
    if not pix.alpha:
        page.insert_image(r, stream=png, rotate=page.rotation, keep_proportion=False, overlay=True)
        return
    a = np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width, pix.n)
    # about 5 dots per point (~360 dpi) is plenty for a signature
    step = max(1, int(min(pix.width / max(w * 5, 1), pix.height / max(h * 5, 1))))
    a = a[::step, ::step]
    ink = a[..., 3] >= 128
    if not ink.any():
        return
    rgb = a[..., :3][ink].astype(np.float32).mean(axis=0) / 255.0
    page.show_pdf_page(r, _stencil_page(ink, rgb, w, h), 0, rotate=page.rotation, keep_proportion=False, overlay=True)


def unembedded_fonts(doc):
    """names of fonts used in the document that are not inside the file (offices often refuse those)"""
    out = set()
    for pg in doc:
        for f in pg.get_fonts(full=True):
            if f[1] in ("n/a", "") and f[2] != "Type3":
                out.add(f[3] or f[4])
    return sorted(out)


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
