"""
PDF Text Localizer
==================
Applies auditor text changes from an Excel (.xlsx) review sheet onto a PDF,
preserving all original layout, images, cartoons, and formatting.

Strategy
--------
1. Read the Excel sheet to extract (page_no, original_text, suggested_text) rows.
2. Render each PDF page to a high-resolution PNG via pdf2image (poppler).
3. Use pdfplumber to get exact word-level bounding boxes (in PDF coordinate space).
4. For each change:
   a. Locate every word in the original text on the target page.
   b. Compute the bounding box that covers all matched words.
   c. White-out that region on the image.
   d. Re-render the suggested text using a matching font/size via Pillow.
5. Assemble all modified page images back into a single PDF via ReportLab.

Dependencies
------------
    pip install pdf2image pillow pdfplumber reportlab pandas openpyxl

System dependency (Linux):
    apt-get install poppler-utils
"""

import os
import sys
import textwrap

import pandas as pd
import pdfplumber
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas as rl_canvas


# ---------------------------------------------------------------------------
# Configuration — edit these paths to run on your own files
# ---------------------------------------------------------------------------
PDF_INPUT  = "/mnt/user-data/uploads/FirstUnit1.pdf"
XLSX_INPUT = "/mnt/user-data/uploads/DOC-20260626-WA0002_.xlsx"
PDF_OUTPUT = "/mnt/user-data/outputs/FirstUnit1_Localized.pdf"

# Render DPI: higher = crisper output but larger file
RENDER_DPI = 200

# Liberation Sans is metrically identical to Arial (used in the original doc)
FONT_REGULAR = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
FONT_BOLD    = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"

# Column names in the Excel sheet
COL_PAGE     = "Page No."
COL_ORIGINAL = "Original Text"
COL_SUGGEST  = "Suggested Localized Text"


# ---------------------------------------------------------------------------
# Step 1 — Load auditor changes from Excel
# ---------------------------------------------------------------------------
def load_changes(xlsx_path: str) -> list[dict]:
    """
    Returns a list of dicts:
        {"page": int, "original": str, "suggested": str}
    Rows where suggested text is NaN are skipped (no change needed).
    """
    df = pd.read_excel(xlsx_path)
    # Normalise column names: strip whitespace, title-case
    df.columns = [c.strip() for c in df.columns]

    changes = []
    for _, row in df.iterrows():
        orig = str(row[COL_ORIGINAL]).strip() if pd.notna(row[COL_ORIGINAL]) else ""
        sugg = str(row[COL_SUGGEST]).strip()  if pd.notna(row[COL_SUGGEST])  else ""
        page = row[COL_PAGE]

        if not orig or not sugg or orig == sugg:
            continue                         # nothing to do
        try:
            page = int(page)
        except (ValueError, TypeError):
            continue

        changes.append({"page": page, "original": orig, "suggested": sugg})

    print(f"[load_changes] {len(changes)} actionable change(s) found")
    return changes


# ---------------------------------------------------------------------------
# Step 2 — Render PDF pages to PIL Images
# ---------------------------------------------------------------------------
def render_pages(pdf_path: str, dpi: int = RENDER_DPI) -> list[Image.Image]:
    """Render every page of the PDF to a RGB PIL Image."""
    images = convert_from_path(pdf_path, dpi=dpi)
    images = [img.convert("RGB") for img in images]
    print(f"[render_pages] {len(images)} page(s) rendered at {dpi} dpi")
    return images


# ---------------------------------------------------------------------------
# Step 3 — Extract word bounding boxes from PDF (all pages)
# ---------------------------------------------------------------------------
def extract_words(pdf_path: str) -> dict[int, list[dict]]:
    """
    Returns {page_number (1-based): [word_dict, ...]}
    Each word_dict has keys: text, x0, x1, top, bottom
    """
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            result[i] = page.extract_words()
    return result


# ---------------------------------------------------------------------------
# Helper — coordinate conversion (PDF pts -> image pixels)
# ---------------------------------------------------------------------------
def make_scaler(pdf_path: str, images: list[Image.Image]):
    """
    Returns a function  pdf_to_px(page_no, x0, top, x1, bottom)
    that maps PDF-point coordinates to image-pixel coordinates.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page_dims = [(p.width, p.height) for p in pdf.pages]

    def pdf_to_px(page_no, x0, top, x1, bottom):
        pw, ph = page_dims[page_no - 1]
        iw, ih = images[page_no - 1].size
        sx = iw / pw
        sy = ih / ph
        return (int(x0 * sx), int(top * sy),
                int(x1 * sx), int(bottom * sy))

    return pdf_to_px


# ---------------------------------------------------------------------------
# Step 4 — Locate original text on a page and return its bounding box
# ---------------------------------------------------------------------------
def find_text_bbox(words: list[dict], original: str):
    """
    Try to match the first sentence / key phrase of `original` within
    the page's word list.  Returns (x0, top, x1, bottom) in PDF pts,
    or None if not found.

    Matching strategy:
      - Build a sliding-window over consecutive words.
      - Normalise whitespace before comparing.
      - Try the full string, then progressively shorter leading fragments
        to handle line-wrap differences.
    """
    norm_original = " ".join(original.split())
    page_tokens   = [w["text"] for w in words]

    def bbox_of_slice(start, end):
        ws = words[start:end]
        return (
            min(w["x0"]    for w in ws),
            min(w["top"]   for w in ws),
            max(w["x1"]    for w in ws),
            max(w["bottom"]for w in ws),
        )

    # Try matching a leading fragment of increasing word count
    orig_tokens = norm_original.split()
    # Minimum meaningful match: first 4 words (or all if fewer)
    min_tokens = min(4, len(orig_tokens))

    for fragment_len in range(len(orig_tokens), min_tokens - 1, -1):
        target = " ".join(orig_tokens[:fragment_len])
        target_len = fragment_len

        for start in range(len(page_tokens) - target_len + 1):
            window = " ".join(page_tokens[start:start + target_len])
            if window == target:
                return bbox_of_slice(start, start + target_len)

    return None  # no match found


# ---------------------------------------------------------------------------
# Step 5 — Detect font style from matched region
# ---------------------------------------------------------------------------
def detect_font_style(words: list[dict], bbox) -> tuple[str, float]:
    """
    Heuristic: look at the height of matched words to estimate pt size.
    Returns (font_path, font_size_in_pt).
    Bold detection: if the matched word's fontname contains 'Bold' use bold.
    Falls back to regular if fontname isn't available.
    """
    x0, top, x1, bottom = bbox
    matched = [w for w in words
               if w["top"] >= top - 2 and w["bottom"] <= bottom + 2
               and w["x0"] >= x0 - 2]

    avg_height = 11.0  # fallback
    if matched:
        avg_height = sum(w["bottom"] - w["top"] for w in matched) / len(matched)

    # Rough pt size ≈ word height (pdfplumber reports in pts already)
    pt_size = round(avg_height)

    # Check if any matched word has Bold in its fontname (pdfplumber char level)
    is_bold = any("Bold" in w.get("fontname", "") for w in matched
                  if "fontname" in w)

    font_path = FONT_BOLD if is_bold else FONT_REGULAR
    return font_path, pt_size


# ---------------------------------------------------------------------------
# Step 6 — White-out a region and draw replacement text
# ---------------------------------------------------------------------------
def apply_text_patch(
    draw:      ImageDraw.ImageDraw,
    img_bbox:  tuple[int, int, int, int],  # pixels
    new_text:  str,
    font_path: str,
    pt_size:   float,
    scale_y:   float,
    bg_color:  tuple = (255, 255, 255),
    text_color:tuple = (0, 0, 0),
    padding:   int   = 4,
):
    """
    1. White-out the bounding box (with a small padding).
    2. Word-wrap the new text to fit the same width.
    3. Draw each wrapped line.
    """
    ix0, iy0, ix1, iy1 = img_bbox
    font_px = max(8, int(pt_size * scale_y))
    font    = ImageFont.truetype(font_path, font_px)

    # --- erase original ---
    draw.rectangle(
        [ix0 - padding, iy0 - padding, ix1 + padding * 3, iy1 + padding * 3],
        fill=bg_color,
    )

    # --- word-wrap to column width ---
    max_width = (ix1 - ix0) + padding * 2
    words_    = new_text.split()
    lines     = []
    line      = ""
    for word in words_:
        candidate = (line + " " + word).strip()
        bbox_test = draw.textbbox((0, 0), candidate, font=font)
        if bbox_test[2] <= max_width:
            line = candidate
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)

    # --- draw lines ---
    line_height = int(font_px * 1.35)
    for i, ln in enumerate(lines):
        draw.text((ix0, iy0 + i * line_height), ln, fill=text_color, font=font)


# ---------------------------------------------------------------------------
# Step 7 — Orchestrate all changes on all pages
# ---------------------------------------------------------------------------
def apply_changes(
    images:   list[Image.Image],
    changes:  list[dict],
    words_by_page: dict[int, list[dict]],
    pdf_to_px,
):
    """Mutates `images` in-place applying every auditor change."""
    draws = [ImageDraw.Draw(img) for img in images]

    for ch in changes:
        page_no   = ch["page"]
        original  = ch["original"]
        suggested = ch["suggested"]

        if page_no < 1 or page_no > len(images):
            print(f"  [skip] page {page_no} out of range")
            continue

        words = words_by_page.get(page_no, [])
        bbox  = find_text_bbox(words, original)

        if bbox is None:
            print(f"  [miss] page {page_no}: could not locate [{original[:60]}...]")
            continue

        x0, top, x1, bottom = bbox
        img_bbox = pdf_to_px(page_no, x0, top, x1, bottom)

        # Determine scale_y for this page
        with pdfplumber.open(PDF_INPUT) as pdf:
            ph = pdf.pages[page_no - 1].height
        scale_y = images[page_no - 1].size[1] / ph

        font_path, pt_size = detect_font_style(words, bbox)

        apply_text_patch(
            draw       = draws[page_no - 1],
            img_bbox   = img_bbox,
            new_text   = suggested,
            font_path  = font_path,
            pt_size    = pt_size,
            scale_y    = scale_y,
        )
        print(f"  [done] page {page_no}: [{original[:50]}...]  ->  [{suggested[:50]}...]")


# ---------------------------------------------------------------------------
# Step 8 — Save modified images as a single PDF
# ---------------------------------------------------------------------------
def save_pdf(images: list[Image.Image], output_path: str, pdf_input: str):
    """Embed each modified page image into a new PDF at the original page size."""
    with pdfplumber.open(pdf_input) as pdf:
        page_dims = [(p.width, p.height) for p in pdf.pages]

    # Save temp PNGs
    tmp_paths = []
    for i, img in enumerate(images):
        p = f"/tmp/_localized_page_{i+1}.png"
        img.save(p, "PNG")
        tmp_paths.append(p)

    # Build PDF
    c = rl_canvas.Canvas(output_path)
    for (pw, ph), tmp in zip(page_dims, tmp_paths):
        c.setPageSize((pw, ph))
        c.drawImage(tmp, 0, 0, width=pw, height=ph, preserveAspectRatio=False)
        c.showPage()
    c.save()

    # Cleanup temp files
    for p in tmp_paths:
        os.remove(p)

    print(f"[save_pdf] Written -> {output_path}  ({os.path.getsize(output_path)//1024} KB)")


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------
def main():
    print("=== PDF Localizer ===")

    # 1. Load changes
    changes = load_changes(XLSX_INPUT)
    if not changes:
        print("No changes to apply. Exiting.")
        sys.exit(0)

    # 2. Render pages
    images = render_pages(PDF_INPUT)

    # 3. Extract word bounding boxes
    words_by_page = extract_words(PDF_INPUT)

    # 4. Build coordinate scaler
    pdf_to_px = make_scaler(PDF_INPUT, images)

    # 5. Apply changes
    print("\nApplying changes:")
    apply_changes(images, changes, words_by_page, pdf_to_px)

    # 6. Write output PDF
    print()
    save_pdf(images, PDF_OUTPUT, PDF_INPUT)
    print("\nDone.")


if __name__ == "__main__":
    main()
