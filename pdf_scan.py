"""Scanned pages: clean up (straighten, specks, contrast, white background, dark edges), find blank pages, and OCR
(make the text searchable).  Everything runs on this computer: OCR uses the Tesseract engine built into PyMuPDF
with the language files in assets/tessdata.
"""
import os
import sys

import numpy as np
import pymupdf as fitz

HERE = os.path.dirname(os.path.abspath(__file__))
TESSDATA = os.path.join(HERE, "assets", "tessdata")
BENGALI_FONT = os.path.join(HERE, "assets", "NotoSansBengali-Regular.ttf")
LANGUAGES = [("English", "eng"), ("Deutsch", "deu"), ("Italiano", "ita"), ("বাংলা (Bengali)", "ben")]


# ------------------------------------------------------------------------------------------------ pictures --
def page_gray(page, dpi=200):
    """the page as a grey picture (numpy, 0 = black .. 255 = white)"""
    pix = page.get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
    a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)[:, :pix.width]
    return a.copy()


def box_blur(a, r):
    """fast blur with a (2r+1) square, using running sums"""
    if r < 1:
        return a.astype(np.float32)
    f = np.pad(a.astype(np.float32), r + 1, mode="edge")
    c = f.cumsum(0).cumsum(1)
    k = 2 * r + 1
    s = c[k:, k:] - c[:-k, k:] - c[k:, :-k] + c[:-k, :-k]
    return s[:a.shape[0], :a.shape[1]] / (k * k)


def rotate_gray(a, angle):
    """turn a grey picture by a small angle (degrees), white fills the corners"""
    if abs(angle) < 0.05:
        return a
    h, w = a.shape
    t = np.deg2rad(angle)
    cy, cx = (h - 1) / 2, (w - 1) / 2
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    xs = np.cos(t) * (xx - cx) + np.sin(t) * (yy - cy) + cx
    ys = -np.sin(t) * (xx - cx) + np.cos(t) * (yy - cy) + cy
    xi, yi = np.rint(xs).astype(np.int32), np.rint(ys).astype(np.int32)
    ok = (xi >= 0) & (xi < w) & (yi >= 0) & (yi < h)
    out = np.full_like(a, 255)
    out[ok] = a[yi[ok], xi[ok]]
    return out


def skew_angle(a, limit=5.0, step=0.25):
    """the angle that makes the text lines most horizontal (projection profile)"""
    small = a[::4, ::4]
    ink = small < 140
    if ink.mean() < 0.002:
        return 0.0
    best, best_score = 0.0, -1.0
    h, w = ink.shape
    ys, xs = np.nonzero(ink)
    cy, cx = h / 2, w / 2
    for ang in np.arange(-limit, limit + 1e-9, step):
        t = np.deg2rad(ang)
        r = np.rint(-np.sin(t) * (xs - cx) + np.cos(t) * (ys - cy) + cy).astype(np.int32)
        r = r[(r >= 0) & (r < h)]
        prof = np.bincount(r, minlength=h).astype(np.float64)
        score = float(np.sum(np.diff(prof) ** 2))
        if score > best_score:
            best, best_score = ang, score
    return float(best)


def despeckle(a, size=2):
    """remove small dark dots that have (almost) no dark neighbours"""
    ink = a < 128
    n = box_blur(ink.astype(np.float32), size) * ((2 * size + 1) ** 2)
    lonely = ink & (n <= 2.5 * size)
    out = a.copy()
    out[lonely] = 255
    return out


def clean_gray(a, straighten=True, specks=True, darker=True, white=True, edges=True, strength=0.6):
    """clean a scanned page picture; strength 0..1"""
    out = a.astype(np.float32)
    if white:
        # even out the paper: divide by the (blurred) paper colour, so shadows and yellow paper become white
        bg = box_blur(np.maximum(out, box_blur(out, 2)), max(8, a.shape[1] // 60))
        out = np.clip(out / np.maximum(bg, 1) * 255.0, 0, 255)
    if darker:
        lo = 60 + 50 * strength                 # everything darker than this becomes black ...
        hi = 235 - 25 * strength                # ... and lighter than this white
        out = np.clip((out - lo) / max(hi - lo, 1) * 255.0, 0, 255)
    out = out.astype(np.uint8)
    if specks:
        out = despeckle(out, 1 if strength < 0.5 else 2)
    if edges:
        # dark bands along the border of a scan (the scanner lid) become white
        h, w = out.shape
        m = max(4, int(min(h, w) * 0.06))
        for sl in ((slice(0, m), slice(None)), (slice(h - m, h), slice(None)),
                   (slice(None), slice(0, m)), (slice(None), slice(w - m, w))):
            band = out[sl]
            rows = band.mean(axis=1 if sl[0] != slice(None) else 0)
            dark = rows < 110
            if sl[0] != slice(None):
                band[dark, :] = 255
            else:
                band[:, dark] = 255
    if straighten:
        out = rotate_gray(out, -skew_angle(out))
    return out


def gray_to_png(a):
    pix = fitz.Pixmap(fitz.csGRAY, a.shape[1], a.shape[0], np.ascontiguousarray(a).tobytes(), False)
    return pix.tobytes("png")


def gray_to_jpeg(a, quality=82):
    pix = fitz.Pixmap(fitz.csGRAY, a.shape[1], a.shape[0], np.ascontiguousarray(a).tobytes(), False)
    return pix.tobytes("jpeg", jpg_quality=quality)


def clean_page(doc, pno, dpi=200, **opts):
    """replace page pno with its cleaned picture (same size); the text layer of a scan is kept out"""
    page = doc[pno]
    rect = page.rect
    rot = page.rotation
    a = clean_gray(page_gray(page, dpi), **opts)
    data = gray_to_jpeg(a) if opts.get("small", True) else gray_to_png(a)
    doc.delete_page(pno)
    new = doc.new_page(pno, width=rect.width, height=rect.height)
    new.insert_image(new.rect, stream=data)
    if rot:
        new.set_rotation(0)


# ------------------------------------------------------------------------------------------------ blank pages --
def is_blank(page, ink_limit=0.003):
    """True when the page has (almost) nothing on it - also for scans with a little dust"""
    a = page_gray(page, dpi=50)
    h, w = a.shape
    inner = a[int(h * .05):int(h * .95), int(w * .05):int(w * .95)]
    return float((inner < 200).mean()) < ink_limit and not page.get_text().strip()


# ----------------------------------------------------------------------------------------------------- OCR --
def has_text(page):
    return len(page.get_text().strip()) > 20


def _font_for(word, has_bn):
    if has_bn and any("ঀ" <= ch <= "৿" for ch in word):
        return "bn"
    return "helv"


def ocr_page(page, languages=("eng",), dpi=300):
    """read the page with OCR and put the words on it as invisible text: the page looks the same, but the text
    can now be searched, selected and copied"""
    lang = "+".join(languages) or "eng"
    tp = page.get_textpage_ocr(language=lang, dpi=dpi, full=True, tessdata=TESSDATA)
    words = page.get_text("words", textpage=tp)
    has_bn = "ben" in languages and os.path.isfile(BENGALI_FONT)
    if has_bn:
        page.insert_font(fontname="bn", fontfile=BENGALI_FONT)
    for w in words:
        x0, y0, x1, y1, text = w[:5]
        if not text.strip():
            continue
        font = _font_for(text, has_bn)
        h = max(y1 - y0, 1)
        size = h * 0.8
        try:
            fobj = fitz.Font(fontfile=BENGALI_FONT) if font == "bn" else fitz.Font("helv")
            tw = fobj.text_length(text, fontsize=size)
            if tw > 0 and (x1 - x0) > 0:
                size *= min(4.0, max(0.25, (x1 - x0) / tw))
            page.insert_text((x0, y1 - h * 0.18), text, fontsize=size, fontname=font, render_mode=3)
        except Exception:                                      # noqa: BLE001 - one odd word must not stop the page
            continue
    return len(words)


def tessdata_ok():
    return os.path.isdir(TESSDATA) and any(f.endswith(".traineddata") for f in os.listdir(TESSDATA))
