"""Digonto PDF - PDFs that official upload portals accept.

Government and other official upload portals convert every upload to PDF/A,
the archive format. Two things break that conversion, and Digonto used to add both:
  * text in a font that is not inside the file (the built-in "Helvetica")
  * pictures with a see-through background (a signature on a soft mask)
So everything Digonto writes now uses fonts that travel inside the file, and signatures are
drawn as solid ink shapes instead of see-through pictures.
`upload_copy` goes further for a file you are about to upload: forms and notes are flattened,
scripts, attachments, bookmarks and tags removed, and any page that would still break the
conversion is turned into a clean picture page.
"""
import math
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
# a new resource name every run: a saved file holds a trimmed copy of the font (only the letters used) under the
# name of that run; reusing such a copy for new text would show empty boxes for letters it does not have
_RUN = f"{os.getpid() % 100000:05d}{int.from_bytes(os.urandom(2), 'big') % 1000:03d}"


def font_kw(name):
    """keyword arguments for page.insert_text / insert_textbox: an embedded font instead of a built-in one"""
    nm, fn = _FONTS.get(name, (None, None))
    path = os.path.join(FONT_DIR, fn) if fn else None
    if path and os.path.exists(path):
        return {"fontname": nm + _RUN, "fontfile": path}
    return {"fontname": name}


def is_picture_page(page):
    """a page that is one big picture with invisible text on top (a scan with OCR, or a page Save for upload made)"""
    try:
        imgs = page.get_image_info()
        area = abs(page.rect)
        if not any(abs(fitz.Rect(i["bbox"]) & page.rect) > 0.9 * area for i in imgs):
            return False
        spans = page.get_texttrace()
        return bool(spans) and all(sp.get("type") == 3 for sp in spans)
    except Exception:
        return False


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
    """names of fonts used in the document that are not inside the file (upload portals often refuse those)"""
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


def _picture_page(out, src, dpi):
    """the page as a clean picture, with its text kept as invisible text (embedded font) so it can still be
    searched, selected and copied - like a scanned page with OCR"""
    pg = out.new_page(width=src.rect.width, height=src.rect.height)
    pix = src.get_pixmap(dpi=dpi, alpha=False)
    pg.insert_image(pg.rect, stream=pix.tobytes("jpeg", jpg_quality=92))
    try:
        kw = font_kw("helv")
        for b in src.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                dx, dy = ln.get("dir", (1, 0))
                ang = round(math.degrees(math.atan2(-dy, dx)) / 90.0) * 90 % 360
                for sp in ln["spans"]:
                    t = sp["text"]
                    if not t.strip():
                        continue
                    size = max(1.0, sp["size"])
                    if ang == 0:                       # stretch so the invisible words sit exactly over the visible ones
                        want = sp["bbox"][2] - sp["bbox"][0]
                        have = text_length(t, "helv", size) or 1
                        size = max(1.0, min(size * 3, size * want / have)) if want > 0 else size
                    pg.insert_text(fitz.Point(sp["origin"]), t, fontsize=size, render_mode=3, rotate=ang, **kw)
    except Exception:
        pass
    return pg


def needs_office_fix(data):
    """True when this PDF has something upload portals refuse (fonts not in the file, see-through pictures, forms, scripts)"""
    try:
        doc = fitz.open(stream=data, filetype="pdf")
    except Exception:
        return False
    if doc.needs_pass or doc.is_encrypted:
        return False
    if doc.is_form_pdf:
        return True
    return any(page_problems(doc, i) for i in range(doc.page_count))


def office_ready(data, dpi=200):
    """for Save / Save as / tool results: exactly the steps of Save for upload (that is what upload portals accepted
    in tests; a lighter version that kept links and bookmarks was refused). Files with a password stay as they are."""
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        if doc.needs_pass or doc.is_encrypted:
            return data
        doc.close()
        return upload_copy(data, dpi=dpi)[0]
    except Exception:
        return data


def shrink_images(doc, max_dpi=215, target=200):
    """pictures shown sharper than max_dpi are scaled to target dpi (JPEG, RGB, no colour profile)"""
    done = set()
    for page in doc:
        for im in page.get_images(full=True):
            xref, smask, w, h, bpc = im[0], im[1], im[2], im[3], im[4]
            if xref in done or smask or bpc == 1 or w * h < 400 * 400:
                continue
            rects = page.get_image_rects(xref)
            if not rects:
                continue
            r = max(rects, key=lambda q: abs(q))
            if r.width <= 0 or r.height <= 0:
                continue
            dpi = max(w / (r.width / 72.0), h / (r.height / 72.0))
            if dpi <= max_dpi:
                continue
            try:
                pix = fitz.Pixmap(doc, xref)
                if pix.alpha or pix.colorspace is None or pix.colorspace.n not in (1, 3):
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                    if pix.alpha:
                        pix = fitz.Pixmap(pix, 0)
                k = target / dpi
                small = fitz.Pixmap(pix, max(1, int(w * k)), max(1, int(h * k)), None)
                page.replace_image(xref, stream=small.tobytes("jpeg", jpg_quality=90))
                done.add(xref)
            except Exception:
                continue


def upload_copy(data, as_picture=False, dpi=200, password=None, keep_nav=False):
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
    # pictures sharper than 200 dpi are scaled down: upload portals refused a page picture at 300 dpi and
    # accepted the same page at 200 dpi (also keeps the file small)
    shrink_images(doc)
    toc = doc.get_toc(simple=False) if keep_nav else []
    if not keep_nav:
        try:
            doc.set_toc([])                          # no bookmarks
        except Exception:
            pass
    cat = doc.pdf_catalog()
    for key in ("AcroForm", "StructTreeRoot", "MarkInfo", "OpenAction", "AA", "Names", "OCProperties") + (() if keep_nav else ("Outlines",)):
        try:
            doc.xref_set_key(cat, key, "null")
        except Exception:
            pass
    finish(doc)
    out = fitz.open()
    pictures = []
    for i in range(doc.page_count):
        if as_picture or page_problems(doc, i):
            pg = _picture_page(out, doc[i], dpi)
            if keep_nav:
                for ln in doc[i].get_links():
                    if ln.get("kind") == fitz.LINK_URI:
                        pg.insert_link(ln)
            pictures.append(i + 1)
        else:
            out.insert_pdf(doc, from_page=i, to_page=i, annots=False, links=keep_nav)
    if keep_nav and toc:
        try:
            out.set_toc(toc)
        except Exception:
            pass
    finish(out)                                       # the invisible text layer: only the letters used
    meta = {k: v for k, v in (doc.metadata or {}).items() if k in ("title", "author", "subject", "keywords")}
    meta.update(producer="Digonto PDF", creator="Digonto PDF")
    out.set_metadata(meta)
    res = out.tobytes(garbage=4, deflate=True, clean=True, use_objstms=0)
    if res.startswith(b"%PDF-1.7"):                  # nothing newer than PDF 1.4 is used: say so
        res = b"%PDF-1.4" + res[8:]
    return res, pictures
