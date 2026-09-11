"""PDF image extraction: embedded rasters first, then AWS Textract
AnalyzeDocument (LAYOUT feature) for detecting figures drawn as vector
graphics or any other visual element the PDF viewer renders as a figure.

Strategy per page:
  1. Extract any embedded raster images as-is (photos, or a figure saved
     as an actual image XObject) -- these are already a clean crop.
  2. Otherwise, render the page to a PNG bitmap, send it to AWS Textract
     AnalyzeDocument with the LAYOUT feature, and crop each detected
     LAYOUT_FIGURE region from the rendered page image.
  3. Only if a page has almost no extractable text at all (a genuine
     scanned page) is it captured as one full-page screenshot.
"""

import io
import os
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
    print("[WARNING] PyMuPDF not installed. Images will NOT be extracted from PDF.")
    print("[WARNING] Install it with: pip install PyMuPDF")

try:
    import boto3
    from botocore.exceptions import ClientError, NoCredentialsError
except ImportError:
    boto3 = None
    print("[WARNING] boto3 not installed. Textract figure detection unavailable.")
    print("[WARNING] Install it with: pip install boto3")

try:
    from PIL import Image as PILImage
except ImportError:
    PILImage = None

# A page is only screenshotted whole when it has less than this much
# extractable text -- i.e. it looks like a genuine scanned page-as-image.
SCANNED_PAGE_TEXT_THRESHOLD = 20

# DPI used to render PDF pages as images for Textract analysis.
RENDER_DPI = 200

# Minimum figure dimensions in pixels (at RENDER_DPI) to skip tiny
# false-positive detections.
MIN_FIGURE_PX = 40

# Padding in pixels added around the cropped figure to avoid clipping
# anti-aliased edges.
CROP_PADDING_PX = 8


# ---------------------------------------------------------------------------
# AWS Textract helpers
# ---------------------------------------------------------------------------

def _get_textract_client():
    """Create a boto3 Textract client using env-var credentials or the
    default credential chain (AWS CLI profile, IAM role, etc.).

    Returns None (with a warning) if credentials are not configured.
    """
    if boto3 is None:
        return None

    region = os.environ.get("AWS_REGION", "us-east-1")
    access_key = os.environ.get("AWS_ACCESS_KEY_ID")
    secret_key = os.environ.get("AWS_SECRET_ACCESS_KEY")

    try:
        if access_key and secret_key:
            client = boto3.client(
                "textract",
                region_name=region,
                aws_access_key_id=access_key,
                aws_secret_access_key=secret_key,
            )
        else:
            # Fall back to default credential chain (AWS CLI, IAM role, ...)
            client = boto3.client("textract", region_name=region)
        return client
    except (NoCredentialsError, ClientError) as e:
        print(f"[WARNING] AWS Textract credentials not configured: {e}")
        return None


def _detect_figures_with_textract(client, page_image_bytes):
    """Call Textract AnalyzeDocument with the LAYOUT feature and return a
    list of normalised bounding boxes for every LAYOUT_FIGURE block,
    sorted in reading order (top-to-bottom, left-to-right).

    Each bbox is a dict with keys: Width, Height, Left, Top (all 0–1
    fractions of the page image dimensions).

    Returns an empty list if no figures are found or the API call fails.
    """
    try:
        response = client.analyze_document(
            Document={"Bytes": page_image_bytes},
            FeatureTypes=["LAYOUT"],
        )
    except (ClientError, Exception) as e:
        print(f"  [WARNING] Textract API call failed: {e}")
        return []

    bboxes = []
    for block in response.get("Blocks", []):
        if block.get("BlockType") == "LAYOUT_FIGURE":
            geometry = block.get("Geometry", {})
            bbox = geometry.get("BoundingBox")
            if bbox:
                bboxes.append(bbox)

    return bboxes


def _is_meaningful_figure(bbox, page_w, page_h):
    """Filter out non-content page elements detected by Textract as LAYOUT_FIGURE:
    - Running headers, page numbers, decorative footer ornaments
    - Tiny UI icons, bullets, and section heading banners
    - Extreme bottom or margin lines
    """
    top = bbox.get("Top", 0)
    left = bbox.get("Left", 0)
    w = bbox.get("Width", 0)
    h = bbox.get("Height", 0)
    w_px = w * page_w
    h_px = h * page_h
    area_px = w_px * h_px

    # 1. Header / footer / corner margin furniture
    if top < 0.10 and h < 0.08:
        return False
    if top > 0.85 and (left < 0.15 or (left + w) > 0.85) and h < 0.10:
        return False
    if top > 0.95:
        return False

    # 2. Tiny decorative icons, bullets, and thin section heading banners
    if w_px < 140 and h_px < 140:
        return False
    if h_px < 98:
        return False
    if area_px < 25000 and (w_px < 160 or h_px < 160):
        return False

    return True


def _merge_adjacent_bboxes(bboxes, x_gap_tolerance=0.005, y_overlap_threshold=0.7):
    """Merge adjacent bounding boxes that belong to the same composite figure
    (e.g., Textract splitting a horizontal screenshot into adjacent tiles).
    """
    if len(bboxes) < 2:
        return bboxes
    changed = True
    merged = list(bboxes)
    while changed:
        changed = False
        new_merged = []
        skip_indices = set()
        for i in range(len(merged)):
            if i in skip_indices:
                continue
            merged_b = dict(merged[i])
            for j in range(i + 1, len(merged)):
                if j in skip_indices:
                    continue
                b2 = merged[j]
                top1, bottom1 = merged_b["Top"], merged_b["Top"] + merged_b["Height"]
                top2, bottom2 = b2["Top"], b2["Top"] + b2["Height"]
                overlap_y = max(0, min(bottom1, bottom2) - max(top1, top2))
                min_h = min(merged_b["Height"], b2["Height"])
                if min_h > 0 and (overlap_y / min_h) >= y_overlap_threshold:
                    left1, right1 = merged_b["Left"], merged_b["Left"] + merged_b["Width"]
                    left2, right2 = b2["Left"], b2["Left"] + b2["Width"]
                    gap_x = max(0, max(left1, left2) - min(right1, right2))
                    if gap_x <= x_gap_tolerance:
                        new_left = min(left1, left2)
                        new_top = min(top1, top2)
                        new_right = max(right1, right2)
                        new_bottom = max(bottom1, bottom2)
                        merged_b["Left"] = new_left
                        merged_b["Top"] = new_top
                        merged_b["Width"] = new_right - new_left
                        merged_b["Height"] = new_bottom - new_top
                        skip_indices.add(j)
                        changed = True
            new_merged.append(merged_b)
        merged = new_merged
    return merged


def _sort_figures_reading_order(bboxes, y_tolerance=0.04):
    """Sort bounding boxes in visual reading order:
    groups bboxes into horizontal rows (within y_tolerance) and sorts each row
    left-to-right, then sequences the rows top-to-bottom.
    """
    if not bboxes:
        return []
    sorted_by_top = sorted(bboxes, key=lambda b: b.get("Top", 0))
    rows = []
    current_row = [sorted_by_top[0]]
    current_row_y = sorted_by_top[0].get("Top", 0)
    for b in sorted_by_top[1:]:
        top = b.get("Top", 0)
        if abs(top - current_row_y) <= y_tolerance:
            current_row.append(b)
        else:
            current_row.sort(key=lambda x: x.get("Left", 0))
            rows.extend(current_row)
            current_row = [b]
            current_row_y = top
    if current_row:
        current_row.sort(key=lambda x: x.get("Left", 0))
        rows.extend(current_row)
    return rows


def _crop_figure(page_pil_image, bbox, padding=CROP_PADDING_PX):
    """Crop a single figure from the full-page PIL image using a
    normalised bounding box (0–1 fractions), and return PNG bytes.

    Args:
        page_pil_image: PIL.Image of the full rendered page.
        bbox: dict with keys Width, Height, Left, Top (0–1 normalised).
        padding: extra pixels around the crop to avoid clipping edges.

    Returns:
        PNG image bytes, or None if the crop is too small.
    """
    img_w, img_h = page_pil_image.size

    # Convert normalised coords to pixel coords
    left = int(bbox["Left"] * img_w)
    top = int(bbox["Top"] * img_h)
    width = int(bbox["Width"] * img_w)
    height = int(bbox["Height"] * img_h)

    # Skip tiny detections (icons, bullets, decorative marks)
    if width < MIN_FIGURE_PX or height < MIN_FIGURE_PX:
        return None

    # Apply padding, clamped to image bounds
    x0 = max(0, left - padding)
    y0 = max(0, top - padding)
    x1 = min(img_w, left + width + padding)
    y1 = min(img_h, top + height + padding)

    cropped = page_pil_image.crop((x0, y0, x1, y1))

    buf = io.BytesIO()
    cropped.save(buf, format="PNG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract_images_from_pdf(pdf_path: str) -> dict:
    """Extract images from a PDF, keyed by (page_index, image_index).

    Strategy, per page:
    1. Primary: Render the page to a high-DPI image, run AWS Textract
       AnalyzeDocument with the LAYOUT feature, and crop each identified
       LAYOUT_FIGURE region. This ensures full-fidelity capture of all
       figures, including composite figures (photos with vector arrows/text
       annotations on top) and pure vector diagrams.
    2. Fallback (if Textract is unavailable or finds no figures on the page):
       Extract embedded raster images as-is with PyMuPDF Pixmap RGB conversion,
       sorted by their vertical position on the page.
    3. Scanned Page Fallback: If a page has almost no extractable text and
       no figures detected, capture the full-page screenshot.

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
    textract_count = 0
    embedded_count = 0
    scanned_count = 0

    # Initialise Textract client once (may be None if not configured)
    textract_client = _get_textract_client()
    if textract_client is None:
        print("[WARNING] AWS Textract not available. Falling back to embedded raster extraction.")
        print("[WARNING] Configure AWS credentials in .env to enable layout figure detection.")

    try:
        pdf_doc = fitz.open(str(pdf_path))
        num_pages = len(pdf_doc)

        for page_idx in range(num_pages):
            page = pdf_doc[page_idx]
            img_on_page = 0

            # Render page to PNG bytes once (used for Textract analysis & cropping)
            pix = page.get_pixmap(dpi=RENDER_DPI)
            page_png_bytes = pix.tobytes("png")
            page_pil = PILImage.open(io.BytesIO(page_png_bytes)) if PILImage else None

            # --- Tier 1: AWS Textract LAYOUT_FIGURE detection ---
            if textract_client is not None and page_pil is not None:
                figure_bboxes = _detect_figures_with_textract(
                    textract_client, page_png_bytes
                )

                if figure_bboxes:
                    # Filter out page furniture, margin decorations, tiny icons, and heading banners
                    meaningful_bboxes = [
                        b for b in figure_bboxes
                        if _is_meaningful_figure(b, page_pil.width, page_pil.height)
                    ]
                    # Merge split diagram parts (e.g. adjacent screenshots split by Textract)
                    merged_bboxes = _merge_adjacent_bboxes(meaningful_bboxes)
                    # Sort in visual reading order (rows top-to-bottom, columns left-to-right)
                    ordered_bboxes = _sort_figures_reading_order(merged_bboxes)

                    for bbox in ordered_bboxes:
                        figure_png = _crop_figure(page_pil, bbox)
                        if figure_png is not None:
                            extracted[(page_idx, img_on_page)] = figure_png
                            img_on_page += 1
                            textract_count += 1

                    if img_on_page > 0:
                        # Successfully extracted figures identified by Textract layout
                        continue

            # --- Tier 2: Fallback to Embedded raster images ---
            image_list = page.get_images(full=True)
            candidate_images = []

            for img_info in image_list:
                xref = img_info[0]
                try:
                    base_image = pdf_doc.extract_image(xref)
                    if base_image is None:
                        continue

                    # Skip tiny images (icons, bullets)
                    width = base_image.get("width", 0)
                    height = base_image.get("height", 0)
                    if width < 50 and height < 50:
                        continue

                    # Get position on page to sort by vertical reading order
                    rects = page.get_image_rects(xref)
                    top_y = rects[0].y0 if rects else 0
                    left_x = rects[0].x0 if rects else 0

                    # Convert CMYK/ICC to RGB PNG using PyMuPDF Pixmap
                    try:
                        img_pix = fitz.Pixmap(pdf_doc, xref)
                        if img_pix.n - img_pix.alpha > 3:
                            img_pix = fitz.Pixmap(fitz.csRGB, img_pix)
                        if img_pix.alpha:
                            img_pix = fitz.Pixmap(fitz.csRGB, img_pix, 0)
                        image_bytes = img_pix.tobytes("png")
                    except Exception:
                        image_bytes = base_image["image"]

                    candidate_images.append((top_y, left_x, image_bytes))

                except Exception as e:
                    print(f"  [WARNING] Could not extract embedded image from page {page_idx + 1}: {e}")

            if candidate_images:
                # Sort embedded images in reading order (rows top-to-bottom, columns left-to-right)
                candidate_images.sort(key=lambda x: x[0])
                sorted_embedded = []
                current_row = [candidate_images[0]]
                for item in candidate_images[1:]:
                    if abs(item[0] - current_row[0][0]) <= 30:  # ~30pt row tolerance
                        current_row.append(item)
                    else:
                        current_row.sort(key=lambda x: x[1])
                        sorted_embedded.extend(current_row)
                        current_row = [item]
                if current_row:
                    current_row.sort(key=lambda x: x[1])
                    sorted_embedded.extend(current_row)

                for _, _, img_bytes in sorted_embedded:
                    extracted[(page_idx, img_on_page)] = img_bytes
                    img_on_page += 1
                    embedded_count += 1

                if img_on_page > 0:
                    continue

            # --- Tier 3: Full-page screenshot for genuinely scanned pages ---
            page_text = page.get_text("text").strip()
            if len(page_text) < SCANNED_PAGE_TEXT_THRESHOLD:
                extracted[(page_idx, 0)] = page_png_bytes
                scanned_count += 1

        pdf_doc.close()

        print(
            f"[INFO] Images: {embedded_count} embedded, {textract_count} Textract "
            f"figures, {scanned_count} full-page scans "
            f"(from {num_pages} pages)."
        )

    except Exception as e:
        print(f"[ERROR] Image extraction failed: {e}")

    return extracted