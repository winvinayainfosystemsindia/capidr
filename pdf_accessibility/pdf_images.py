"""PDF image extraction: embedded rasters first, then a caption-anchored crop
of vector-drawn figures (the common case for geometry/statistics diagrams in
textbook PDFs, which are drawn as PDF vector paths rather than embedded
images), with a full-page-screenshot fallback only for genuinely scanned
pages.

Why not just crop the union of every vector drawing on the page? Real-world
PDFs draw page furniture -- a full-bleed decorative border, printer's crop
marks, colour-registration swatches -- as vector paths too, indistinguishable
from a real figure by shape alone. And these textbook pages also carry
colour-block callout boxes (definition/note boxes) that are themselves
vector-drawn rectangles but aren't figures. Reliably telling the two apart
by geometry alone is fragile. Instead, this module anchors on the figure's
own caption ("படம் N.N" -- Tamil for "Figure N.N", printed under every
diagram in this document) and crops to the vector cluster that sits directly
above it, which is both simpler and matches the "Figures: sequential
captions only for images that carry captions in the source" rule the LLM
extraction step is already instructed to follow -- so the figure this module
crops lines up with the figure the JSON describes.
"""

import io
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
    print("[WARNING] PyMuPDF not installed. Images will NOT be extracted from PDF.")
    print("[WARNING] Install it with: pip install PyMuPDF")

# A page is only screenshotted whole when it has less than this much
# extractable text -- i.e. it looks like a genuine scanned page-as-image,
# not a text page that simply has no captioned figure on it.
SCANNED_PAGE_TEXT_THRESHOLD = 20

# A single vector-drawing rect covering at least this fraction of the page
# in BOTH dimensions is page furniture (the printed border/frame) rather
# than a distinct figure, and is dropped before any clustering happens.
FURNITURE_MIN_PAGE_FRACTION = 0.85

# Points of padding added around a detected figure so the crop doesn't clip
# anti-aliased edges -- capped per-side so it never eats into adjacent text
# (see _pad_without_clipping_text).
FIGURE_BBOX_PADDING = 10

# How far (points) a vector cluster's bottom edge may sit above a caption's
# top edge and still count as "directly above" it.
CAPTION_VERTICAL_SLACK = 6

# How far (points) a vector cluster may sit outside a caption's left/right
# edges and still count as being in the same column as it.
CAPTION_HORIZONTAL_SLACK = 60

# Caption words for "படம்" (Tamil for "figure") sometimes extract as
# "படடம்" -- an extra consonant glyph -- depending on how the source PDF
# shaped/ordered the glyphs. Match the stable prefix instead of the exact
# word, and keep it short so this doesn't also match unrelated body text.
FIGURE_CAPTION_PREFIX = "பட"
FIGURE_CAPTION_MAX_LEN = 6


def _is_furniture(rect, page_rect):
    """True for a vector-drawing rect that is really page furniture (the
    printed border/frame) rather than a distinct figure."""
    return (
        rect.width >= FURNITURE_MIN_PAGE_FRACTION * page_rect.width
        and rect.height >= FURNITURE_MIN_PAGE_FRACTION * page_rect.height
    )


def _rects_close(a, b, pad):
    """Manual overlap/proximity test, padded by `pad` points on every side.

    fitz.Rect.intersects() answers via `not (a & b).is_empty`, and
    `is_empty` is True for any zero-width OR zero-height rect -- so it
    reports "no intersection" for two rects that plainly touch or overlap
    if either one is a perfectly horizontal/vertical line. Figure diagrams
    are full of exactly such lines (axis strokes, drop-lines to a point),
    so clustering needs its own interval-overlap check instead of relying
    on intersects().
    """
    if a.x0 - pad > b.x1 or b.x0 - pad > a.x1:
        return False
    if a.y0 - pad > b.y1 or b.y0 - pad > a.y1:
        return False
    return True


def _safe_union(a, b):
    """Union of two rects by raw coordinates. fitz's `|=`/`|` silently
    no-ops when either operand is zero-width or zero-height (its
    `is_empty` is True for those, and the built-in union skips "empty"
    operands) -- which would drop every straight axis/drop-line a diagram
    is drawn from. Every merge in this module goes through this instead.
    """
    return fitz.Rect(
        min(a.x0, b.x0), min(a.y0, b.y0),
        max(a.x1, b.x1), max(a.y1, b.y1),
    )


def _cluster_rects(rects, pad=12):
    """Merge overlapping/nearby vector-drawing rects into one bounding box
    per distinct figure. A diagram is made of many small, separate path
    rects (lines, bezier arcs, point dots) sitting close together --
    clustering recombines them into a single box per figure."""
    clusters = [fitz.Rect(r) for r in rects]
    changed = True
    while changed:
        changed = False
        merged = []
        used = [False] * len(clusters)
        for i, base in enumerate(clusters):
            if used[i]:
                continue
            base = fitz.Rect(base)
            for j in range(i + 1, len(clusters)):
                if used[j]:
                    continue
                if _rects_close(base, clusters[j], pad):
                    base = _safe_union(base, clusters[j])
                    used[j] = True
                    changed = True
            used[i] = True
            merged.append(base)
        clusters = merged
    return clusters


def _find_figure_captions(page):
    """Locate 'படம் N.N' caption words on the page, sorted top-to-bottom."""
    caps = []
    for w in page.get_text("words"):
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        if text.startswith(FIGURE_CAPTION_PREFIX) and len(text) <= FIGURE_CAPTION_MAX_LEN:
            caps.append(fitz.Rect(x0, y0, x1, y1))
    caps.sort(key=lambda r: r.y0)
    return caps


def _expand_with_labels(rect, words, band_gap=55, band_min_overlap=6):
    """Grow a figure's vector bbox to swallow nearby point/axis labels that
    are drawn as ordinary text (e.g. 'P(x, y)', 'A(x1, y1)', 'C') rather
    than vector paths, so they don't get clipped at the crop edge.

    A word is pulled in when it vertically overlaps the *original* figure
    rect by at least `band_min_overlap` points (it sits at the same height
    as part of the figure -- how a point label sits beside a circle) and is
    within `band_gap` points of it horizontally. Checking overlap against
    the original rect (not the growing one) is what keeps this from
    chaining sideways into a whole paragraph: a heading sitting just above
    or below the figure has no vertical overlap with it at all and is
    never pulled in, regardless of how close the gap is.
    """
    original = fitz.Rect(rect)
    grown = fitz.Rect(rect)
    for w in words:
        wr = fitz.Rect(w[0], w[1], w[2], w[3])
        if original.contains(wr):
            continue
        vert_overlap = min(wr.y1, original.y1) - max(wr.y0, original.y0)
        if vert_overlap >= min(band_min_overlap, wr.height):
            gap = max(original.x0 - wr.x1, wr.x0 - original.x1)
            if 0 <= gap <= band_gap:
                grown = _safe_union(grown, wr)
    return grown


def _pad_without_clipping_text(rect, words, pad):
    """Pad `rect` by up to `pad` points on each side, but never let the
    padding cross into a text word that isn't already part of the figure
    (e.g. a heading sitting just above the diagram, or an equation right
    next to it). Padding exists only to avoid clipping anti-aliased vector
    edges, not to eat into adjacent text.
    """
    x0, y0, x1, y1 = rect.x0, rect.y0, rect.x1, rect.y1

    def room(direction):
        limit = pad
        for w in words:
            wr = fitz.Rect(w[0], w[1], w[2], w[3])
            if rect.contains(wr):
                continue
            if direction == "top" and wr.y1 <= y0 and wr.x1 >= x0 and wr.x0 <= x1:
                limit = min(limit, max(0, y0 - wr.y1 - 1))
            elif direction == "bottom" and wr.y0 >= y1 and wr.x1 >= x0 and wr.x0 <= x1:
                limit = min(limit, max(0, wr.y0 - y1 - 1))
            elif direction == "left" and wr.x1 <= x0 and wr.y1 >= y0 and wr.y0 <= y1:
                limit = min(limit, max(0, x0 - wr.x1 - 1))
            elif direction == "right" and wr.x0 >= x1 and wr.y1 >= y0 and wr.y0 <= y1:
                limit = min(limit, max(0, wr.x0 - x1 - 1))
        return limit

    return fitz.Rect(
        x0 - room("left"), y0 - room("top"),
        x1 + room("right"), y1 + room("bottom"),
    )


def _vector_figure_bboxes(page):
    """Return one tightly-cropped bounding box per captioned figure on the
    page, in top-to-bottom order. Empty list if the page has no vector
    drawings or no 'படம் N.N' captions."""
    try:
        drawings = page.get_drawings()
    except Exception:
        return []
    if not drawings:
        return []

    page_rect = page.rect
    rects = []
    for d in drawings:
        r = d.get("rect")
        if r is None:
            continue
        # NOTE: fitz.Rect.is_empty is True for a zero-width OR zero-height
        # rect -- but a perfectly horizontal/vertical stroke (an axis line,
        # a drop line to a point) legitimately has zero width or height and
        # is still real figure content. Only drop true zero-area points.
        if r.width <= 0 and r.height <= 0:
            continue
        if _is_furniture(r, page_rect):
            continue
        rects.append(r)
    if not rects:
        return []

    clusters = _cluster_rects(rects)
    captions = _find_figure_captions(page)
    if not captions:
        return []

    used = [False] * len(clusters)
    bboxes = []
    for cap in captions:
        # Pick the nearest not-yet-used cluster sitting directly above this
        # caption, in the same column.
        best_idx, best_gap = None, None
        for idx, cl in enumerate(clusters):
            if used[idx]:
                continue
            if cl.y1 > cap.y0 + CAPTION_VERTICAL_SLACK:
                continue
            if cl.x1 < cap.x0 - CAPTION_HORIZONTAL_SLACK or cl.x0 > cap.x1 + CAPTION_HORIZONTAL_SLACK:
                continue
            gap = cap.y0 - cl.y1
            if best_gap is None or gap < best_gap:
                best_gap, best_idx = gap, idx

        if best_idx is not None:
            used[best_idx] = True
            r = clusters[best_idx]
            words = page.get_text("words")
            r = _expand_with_labels(r, words)
            r = _pad_without_clipping_text(r, words, FIGURE_BBOX_PADDING)
            r = r & page_rect
            bboxes.append(r)

    return bboxes


def extract_images_from_pdf(pdf_path: str) -> dict:
    """Extract images from a PDF, keyed by (page_index, image_index).

    Strategy, per page:
    1. Extract any embedded raster images as-is (photos, or a figure saved
       as an actual image XObject) -- these are already a clean crop.
    2. Otherwise, for diagrams drawn as PDF vector paths (the common case
       for geometry/statistics figures in textbook PDFs), find each
       'படம் N.N' caption and crop tightly to the drawing cluster sitting
       directly above it, expanded to include its point/axis labels.
    3. Only if a page has almost no extractable text at all (a genuine
       scanned page) is it captured as one full-page screenshot. Ordinary
       text pages with no embedded image and no captioned vector figure are
       left with no image, rather than being screenshotted whole.

    Returns:
        dict mapping (page_idx, img_idx) -> PNG image bytes
    """
    if fitz is None:
        print("[WARNING] PyMuPDF not available. Skipping image extraction.")
        return {}

    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        print(f"[WARNING] PDF not found for image extraction: {pdf_path}")
        return {}

    print("[INFO] Extracting images from PDF...")
    extracted = {}
    embedded_count = 0
    cropped_count = 0
    scanned_count = 0

    try:
        pdf_doc = fitz.open(str(pdf_path))
        num_pages = len(pdf_doc)

        for page_idx in range(num_pages):
            page = pdf_doc[page_idx]
            image_list = page.get_images(full=True)
            img_on_page = 0

            # --- Embedded raster images (already a clean crop) ---
            for img_info in image_list:
                xref = img_info[0]
                try:
                    base_image = pdf_doc.extract_image(xref)
                    if base_image is None:
                        continue

                    image_bytes = base_image["image"]
                    image_ext = base_image.get("ext", "png")

                    # Convert to PNG for consistency
                    if image_ext.lower() != "png":
                        try:
                            from PIL import Image as PILImage
                            pil_img = PILImage.open(io.BytesIO(image_bytes))
                            png_buffer = io.BytesIO()
                            pil_img.save(png_buffer, format="PNG")
                            image_bytes = png_buffer.getvalue()
                        except Exception:
                            pass

                    # Skip tiny images (icons, bullets)
                    width = base_image.get("width", 0)
                    height = base_image.get("height", 0)
                    if width < 50 and height < 50:
                        continue

                    extracted[(page_idx, img_on_page)] = image_bytes
                    img_on_page += 1
                    embedded_count += 1

                except Exception as e:
                    print(f"  [WARNING] Could not extract image from page {page_idx + 1}: {e}")

            if img_on_page > 0:
                # This page's images came from real embedded rasters -- don't
                # also try to crop vector figures on it.
                continue

            # --- No embedded raster: crop captioned vector figures ---
            bboxes = _vector_figure_bboxes(page)
            if bboxes:
                for bbox in bboxes:
                    pix = page.get_pixmap(dpi=200, clip=bbox)
                    extracted[(page_idx, img_on_page)] = pix.tobytes("png")
                    img_on_page += 1
                    cropped_count += 1
                continue

            # --- Last resort: only screenshot genuinely scanned pages ---
            page_text = page.get_text("text").strip()
            if len(page_text) < SCANNED_PAGE_TEXT_THRESHOLD:
                pix = page.get_pixmap(dpi=200)
                extracted[(page_idx, 0)] = pix.tobytes("png")
                scanned_count += 1

        pdf_doc.close()

        print(
            f"[INFO] Images: {embedded_count} embedded, {cropped_count} cropped "
            f"vector figures, {scanned_count} full-page scans "
            f"(from {num_pages} pages)."
        )

    except Exception as e:
        print(f"[ERROR] Image extraction failed: {e}")

    return extracted