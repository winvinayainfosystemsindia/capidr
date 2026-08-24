"""PDF image extraction (embedded rasters, with a full-page-screenshot fallback)."""

import io
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
    print("[WARNING] PyMuPDF not installed. Images will NOT be extracted from PDF.")
    print("[WARNING] Install it with: pip install PyMuPDF")


def extract_images_from_pdf(pdf_path: str) -> dict:
    """Extract images from a PDF, keyed by (page_index, image_index).

    Strategy:
    1. Try extracting embedded raster images per page.
    2. If no embedded images found, render each page as a high-res PNG
       (handles vector graphics, scanned pages, diagrams).

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
    try:
        pdf_doc = fitz.open(str(pdf_path))
        num_pages = len(pdf_doc)
        total_images = 0

        # --- Pass 1: Try embedded raster images ---
        for page_idx in range(num_pages):
            page = pdf_doc[page_idx]
            image_list = page.get_images(full=True)
            img_on_page = 0

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
                    total_images += 1

                except Exception as e:
                    print(f"  [WARNING] Could not extract image from page {page_idx + 1}: {e}")

        # --- Pass 2: If no embedded images, render pages as screenshots ---
        if total_images == 0:
            print("[INFO] No embedded raster images found. Rendering pages as screenshots...")
            for page_idx in range(num_pages):
                page = pdf_doc[page_idx]
                # Render at 200 DPI for good quality
                pix = page.get_pixmap(dpi=200)
                png_bytes = pix.tobytes("png")
                # Store as (page_idx, 0) — one image per page
                extracted[(page_idx, 0)] = png_bytes
                total_images += 1

            print(f"[INFO] Rendered {total_images} page screenshots.")
        else:
            print(f"[INFO] Extracted {total_images} embedded images from {num_pages} pages.")

        pdf_doc.close()

    except Exception as e:
        print(f"[ERROR] Image extraction failed: {e}")

    return extracted
