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
    list of normalised bounding boxes for every LAYOUT_FIGURE block.

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
    1. Extract any embedded raster images as-is (photos, or a figure saved
       as an actual image XObject) -- these are already a clean crop.
    2. Otherwise, render the page to a PNG and use AWS Textract's LAYOUT
       analysis to detect LAYOUT_FIGURE regions, then crop each one.
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
    textract_count = 0
    scanned_count = 0

    # Initialise Textract client once (may be None if not configured)
    textract_client = _get_textract_client()
    if textract_client is None:
        print("[WARNING] AWS Textract not available. Only embedded raster images will be extracted.")
        print("[WARNING] Configure AWS credentials to enable figure detection.")

    try:
        pdf_doc = fitz.open(str(pdf_path))
        num_pages = len(pdf_doc)

        for page_idx in range(num_pages):
            page = pdf_doc[page_idx]
            image_list = page.get_images(full=True)
            img_on_page = 0

            # --- Tier 1: Embedded raster images (already a clean crop) ---
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
                # This page's images came from real embedded rasters -- skip
                # Textract analysis for it.
                continue

            # --- Tier 2: Textract LAYOUT_FIGURE detection ---
            if textract_client is not None:
                # Render page to PNG bytes for Textract
                pix = page.get_pixmap(dpi=RENDER_DPI)
                page_png_bytes = pix.tobytes("png")

                # Detect figures via Textract
                figure_bboxes = _detect_figures_with_textract(
                    textract_client, page_png_bytes
                )

                if figure_bboxes:
                    # Open the rendered page as a PIL image for cropping
                    page_pil = PILImage.open(io.BytesIO(page_png_bytes))

                    for bbox in figure_bboxes:
                        figure_png = _crop_figure(page_pil, bbox)
                        if figure_png is not None:
                            extracted[(page_idx, img_on_page)] = figure_png
                            img_on_page += 1
                            textract_count += 1

                    if img_on_page > 0:
                        continue

            # --- Tier 3: Full-page screenshot for genuinely scanned pages ---
            page_text = page.get_text("text").strip()
            if len(page_text) < SCANNED_PAGE_TEXT_THRESHOLD:
                pix = page.get_pixmap(dpi=RENDER_DPI)
                extracted[(page_idx, 0)] = pix.tobytes("png")
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