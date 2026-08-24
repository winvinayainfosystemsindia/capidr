"""PDF image extraction: embedded rasters first, then a caption-anchored crop
of vector-drawn figures, with a complexity-based fallback for uncaptioned
vector diagrams, and a full-page screenshot only as a last resort for
genuinely scanned pages.

Generic by design -- nothing here is tied to one document's language,
caption wording, or layout:
  - A caption is identified structurally (a short, numbered, isolated text
    block), never by matching a specific word -- "படம் 3.4", "Figure 3.4",
    "Fig. 2", "Abbildung 5", a bare "3.4", all match the same rule.
  - A figure is told apart from decorative chrome (a colored callout/
    definition box, a section badge or logo, an equation's vector-drawn
    root sign or fraction bar) by how it's actually drawn, not by guessing
    at what it depicts: diagrams are built from unfilled stroked lines and
    curves; decorative panels are built from solid fills; a caption-like
    label drawn *on top of* a shape (a box header, a badge caption) marks
    it as chrome rather than a figure with a genuine adjacent caption.
  - Diagrams with no caption at all still get extracted, via the same
    shape-based scoring used to reject decorative elements, applied to
    every remaining vector cluster on the page.
"""

import io
import re
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
    print("[WARNING] PyMuPDF not installed. Images will NOT be extracted from PDF.")
    print("[WARNING] Install it with: pip install PyMuPDF")

# A page is only screenshotted whole when it has less than this much
# extractable text -- i.e. it looks like a genuine scanned page-as-image.
SCANNED_PAGE_TEXT_THRESHOLD = 20

# A single vector-drawing rect covering at least this fraction of the page
# in BOTH dimensions is page furniture (a printed border/frame) rather than
# a distinct figure, and is dropped before any clustering happens.
FURNITURE_MIN_PAGE_FRACTION = 0.85

# Points of padding added around a detected figure so the crop doesn't clip
# anti-aliased edges -- capped per-side so it never eats into adjacent text
# (see _pad_without_clipping_text).
FIGURE_BBOX_PADDING = 10

# How far (points) a vector cluster's near edge may sit from a caption's
# near edge and still count as "belonging to" it.
CAPTION_VERTICAL_SLACK = 8
CAPTION_HORIZONTAL_SLACK = 60
CAPTION_MAX_GAP = 40

# A caption candidate is a short, numbered, otherwise-isolated line of text
# -- independent of language or the word used for "figure". 2 words keeps
# out lone equation-number labels like "(1)"; 6 keeps out full sentences
# that merely happen to contain a number.
CAPTION_MIN_WORDS = 2
CAPTION_MAX_WORDS = 6
_HAS_DIGIT_RE = re.compile(r"\d")

# --- Fallback tier: pages with vector diagrams but no detected caption ---
# A genuine diagram (circle + radius line + labels, an axis system, a bar
# chart, ...) is drawn from many separate path items -- lines, bezier
# curves, point markers. A decorative callout/definition box is typically
# one or two fill/stroke paths for the box itself. This threshold is what
# tells them apart when there's no caption to anchor to.
MIN_FIGURE_STROKE_ITEMS = 3
MAX_FILL_AREA_FRACTION = 0.5
MIN_FIGURE_SIZE = 30  # points, both width and height
MAX_FIGURE_PAGE_FRACTION = 0.5  # both width and height, vs. the page

# A cluster with no curved (bezier) stroke at all -- straight lines only --
# and a wide, flat aspect ratio is far more often a piece of math notation
# (a root sign, a fraction bar, a bracket -- all drawn from straight
# lines) than a diagram. Real diagrams almost always contain at least one
# curve (a circle, an arc) or are closer to square. This only rejects the
# straight-AND-wide combination, so genuinely straight-edged diagrams
# (a square, a triangle, a modestly-proportioned bar chart) are unaffected.
MAX_ASPECT_WITHOUT_CURVE = 2.5

# Headers/footers (running titles, page numbers, file/print metadata) live
# in a thin band at the very top/bottom of the page and can otherwise look
# exactly like a short numbered caption ("Page 12", a date stamp). Caption
# candidates inside this margin band are ignored.
PAGE_MARGIN_BAND_FRACTION = 0.06


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


def _cluster_shapes(shapes, pad=12):
    """Merge overlapping/nearby vector-drawing shapes into one shape per
    distinct figure. `shapes` is a list of {"rect": fitz.Rect, "items": n}.
    A diagram is made of many small, separate path rects (lines, bezier
    arcs, point dots) sitting close together -- clustering recombines them
    into a single box (with a summed item count) per figure."""
    clusters = [dict(s) for s in shapes]
    changed = True
    while changed:
        changed = False
        merged = []
        used = [False] * len(clusters)
        for i, base in enumerate(clusters):
            if used[i]:
                continue
            base = dict(base)
            for j in range(i + 1, len(clusters)):
                if used[j]:
                    continue
                if _rects_close(base["rect"], clusters[j]["rect"], pad):
                    base["rect"] = _safe_union(base["rect"], clusters[j]["rect"])
                    base["items"] += clusters[j]["items"]
                    base["stroke_items"] += clusters[j]["stroke_items"]
                    base["fill_area"] += clusters[j]["fill_area"]
                    base["has_curve"] = base["has_curve"] or clusters[j]["has_curve"]
                    used[j] = True
                    changed = True
            used[i] = True
            merged.append(base)
        clusters = merged
    return clusters


def _find_caption_candidates(page):
    """Locate short, numbered, isolated text blocks -- caption candidates
    -- sorted top-to-bottom. Works for any language/caption word ("Figure
    3.4", "படம் 3.4", "Fig. 2", a bare "3.4", ...): a caption is identified
    structurally (short + numbered + its own text block), never by
    matching a specific word.
    """
    candidates = []
    try:
        page_dict = page.get_text("dict")
    except Exception:
        return candidates

    page_rect = page.rect
    top_band = page_rect.y0 + PAGE_MARGIN_BAND_FRACTION * page_rect.height
    bottom_band = page_rect.y1 - PAGE_MARGIN_BAND_FRACTION * page_rect.height

    for block in page_dict.get("blocks", []):
        if block.get("type") != 0:  # text blocks only
            continue
        words = []
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                words.extend(text.split())
        if not (CAPTION_MIN_WORDS <= len(words) <= CAPTION_MAX_WORDS):
            continue
        if not any(_HAS_DIGIT_RE.search(w) for w in words):
            continue
        bbox = block.get("bbox")
        if not bbox:
            continue
        r = fitz.Rect(bbox)
        if r.y0 < top_band or r.y1 > bottom_band:
            continue  # header/footer margin band
        candidates.append(r)

    candidates.sort(key=lambda r: r.y0)
    return candidates


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


def _collect_shapes(page):
    """Non-furniture vector-drawing shapes on the page, each as
    {"rect": fitz.Rect, "items": path-item count, "stroke_items": count of
    items that are pure unfilled strokes}.

    The stroke/fill split matters: a genuine diagram (a circle, an axis
    system, a plotted line) is drawn almost entirely from unfilled stroked
    paths -- lines and curves -- with at most a couple of small filled
    markers (an arrowhead, a point dot). Decorative chrome -- a colored
    callout-box background, an icon/badge/logo -- is drawn from filled
    shapes, rarely if ever a bare stroke. That split is what lets this
    module tell "Definition 3.2"'s colored box, or a section badge, apart
    from an actual figure, generically, in any document.
    """
    try:
        drawings = page.get_drawings()
    except Exception:
        return []
    if not drawings:
        return []

    page_rect = page.rect
    shapes = []
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
        item_count = max(1, len(d.get("items", [])))
        is_stroke_only = d.get("type") == "s"
        is_fillish = d.get("type") in ("f", "fs")
        has_curve = any(it[0] == "c" for it in d.get("items", []))
        shapes.append({
            "rect": r,
            "items": item_count,
            "stroke_items": item_count if is_stroke_only else 0,
            "fill_area": (r.width * r.height) if is_fillish else 0.0,
            "has_curve": has_curve,
        })
    return shapes


def _finalize_bbox(rect, page):
    """Expand a raw cluster rect with nearby labels and safe padding."""
    words = page.get_text("words")
    r = _expand_with_labels(rect, words)
    r = _pad_without_clipping_text(r, words, FIGURE_BBOX_PADDING)
    return r & page.rect


def _caption_mostly_inside(shape_rect, cap_rect, threshold=0.5):
    """True when a large majority of the caption's own area is covered by
    `shape_rect` -- i.e. the caption text is drawn essentially on top of
    the shape (a callout-box header, a badge/logo label), not merely
    touching its edge. A real figure's caption can legitimately graze the
    diagram's bounding box by a few points (an axis line ending close to
    where the caption starts) without being "inside" it in this sense, so
    this needs to be a fraction-of-area test, not a bare intersects().
    """
    ix0, iy0 = max(shape_rect.x0, cap_rect.x0), max(shape_rect.y0, cap_rect.y0)
    ix1, iy1 = min(shape_rect.x1, cap_rect.x1), min(shape_rect.y1, cap_rect.y1)
    iw, ih = max(0.0, ix1 - ix0), max(0.0, iy1 - iy0)
    cap_area = max(1e-6, cap_rect.width * cap_rect.height)
    return (iw * ih) / cap_area >= threshold


def _looks_like_figure(cluster, page_rect):
    """Universal sanity filter applied to every candidate cluster, whether
    it was matched via a caption or picked up by the no-caption fallback:
    - too small (a bullet, a stray symbol)
    - covering most of the page (a background panel)
    - not enough unfilled stroked paths (a diagram is line art -- a
      decorative color box, icon, or badge is drawn from fills, not
      strokes; see _collect_shapes)
    is never a figure.
    """
    r = cluster["rect"]
    if r.width < MIN_FIGURE_SIZE or r.height < MIN_FIGURE_SIZE:
        return False
    if r.width > MAX_FIGURE_PAGE_FRACTION * page_rect.width and r.height > MAX_FIGURE_PAGE_FRACTION * page_rect.height:
        return False
    if cluster["stroke_items"] < MIN_FIGURE_STROKE_ITEMS:
        return False
    cluster_area = max(1.0, r.width * r.height)
    if cluster["fill_area"] / cluster_area > MAX_FILL_AREA_FRACTION:
        return False
    if not cluster["has_curve"]:
        aspect = max(r.width, r.height) / max(1.0, min(r.width, r.height))
        if aspect > MAX_ASPECT_WITHOUT_CURVE:
            return False
    return True


def _vector_figure_bboxes(page):
    """Return one tightly-cropped bounding box per figure on the page, in
    top-to-bottom order. Prefers anchoring each figure to a nearby caption
    (works for any language/wording); if the page has vector diagrams but
    no caption at all, falls back to picking out clusters that look like
    real diagrams by path density/size rather than a decorative box.
    """
    shapes = _collect_shapes(page)
    if not shapes:
        return []

    clusters = _cluster_shapes(shapes)
    captions = _find_caption_candidates(page)

    bboxes = []
    used = [False] * len(clusters)
    page_rect = page.rect

    if captions:
        for cap in captions:
            best_idx, best_gap = None, None
            for idx, cl in enumerate(clusters):
                if used[idx] or not _looks_like_figure(cl, page_rect):
                    continue
                cr = cl["rect"]
                if _caption_mostly_inside(cr, cap):
                    continue  # caption drawn on top of the shape -> a label, not a figure caption
                # caption directly below the cluster, or directly above it
                below = cr.y1 <= cap.y0 + CAPTION_VERTICAL_SLACK
                above = cap.y1 <= cr.y0 + CAPTION_VERTICAL_SLACK
                if not (below or above):
                    continue
                if cr.x1 < cap.x0 - CAPTION_HORIZONTAL_SLACK or cr.x0 > cap.x1 + CAPTION_HORIZONTAL_SLACK:
                    continue
                gap = (cap.y0 - cr.y1) if below else (cr.y0 - cap.y1)
                if gap > CAPTION_MAX_GAP:
                    continue
                if best_gap is None or gap < best_gap:
                    best_gap, best_idx = gap, idx

            if best_idx is not None:
                used[best_idx] = True
                r = clusters[best_idx]["rect"]
                bboxes.append((r, r.y0))

    # Fallback: any remaining/unused cluster that looks like a genuine
    # diagram (not a decorative box) by path density and size, even
    # without a caption to anchor it to. A cluster that has ANY caption
    # candidate drawn on top of it (a callout-box header, a badge label)
    # is a labeled decorative element, not a bare figure -- skip it.
    for idx, cl in enumerate(clusters):
        if used[idx] or not _looks_like_figure(cl, page_rect):
            continue
        if any(_caption_mostly_inside(cl["rect"], cap) for cap in captions):
            continue
        used[idx] = True
        r = cl["rect"]
        bboxes.append((r, r.y0))

    bboxes.sort(key=lambda t: t[1])
    return [_finalize_bbox(r, page) for r, _ in bboxes]


def extract_images_from_pdf(pdf_path: str) -> dict:
    """Extract images from a PDF, keyed by (page_index, image_index).

    Strategy, per page:
    1. Extract any embedded raster images as-is (photos, or a figure saved
       as an actual image XObject) -- these are already a clean crop.
    2. Otherwise, for diagrams drawn as PDF vector paths (the common case
       for figures in textbook/scientific PDFs), find each figure caption
       -- in whatever language/wording, detected structurally rather than
       by matching specific words -- and crop tightly to the drawing
       cluster next to it, expanded to include its point/axis labels. If a
       page has vector diagrams but no caption at all, fall back to
       picking out clusters that look like real diagrams by path density.
    3. Only if a page has almost no extractable text at all (a genuine
       scanned page) is it captured as one full-page screenshot. Ordinary
       text pages with no embedded image and no detected figure are left
       with no image, rather than being screenshotted whole.

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

            # --- No embedded raster: crop vector figures ---
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