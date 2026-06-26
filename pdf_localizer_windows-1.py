"""
PDF Text Localizer  —  Windows-compatible version
==================================================
Applies auditor text changes from an Excel (.xlsx) review sheet onto a PDF,
preserving all original layout, images, cartoons, and formatting.

WINDOWS SETUP (one-time)
-------------------------
1. Install Python 3.10+ from https://python.org

2. Install Python packages:
       pip install pdf2image pillow pdfplumber reportlab pandas openpyxl

3. Install Poppler for Windows (required by pdf2image):
   a. Download the latest release from:
      https://github.com/oschwartz10612/poppler-windows/releases
      (e.g. Release-24.08.0-0.zip)
   b. Extract to a folder, e.g.  C:\poppler
   c. Set POPPLER_PATH below to the "bin" subfolder inside that extracted folder,
      e.g.  r"C:\poppler\Library\bin"
   d. You do NOT need to add it to system PATH — the script passes it directly.

4. Font note:
   Windows ships with Arial (metrically identical to Liberation Sans).
   The script auto-detects common Windows font paths.
   If your system has Arial, no changes needed.
   Otherwise set FONT_REGULAR / FONT_BOLD to any .ttf you like.

USAGE
-----
   python pdf_localizer_windows.py

   Or pass paths as arguments:
   python pdf_localizer_windows.py input.pdf changes.xlsx output.pdf
"""

import os
import sys
import tempfile

import pandas as pd
import pdfplumber
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas as rl_canvas


# ===========================================================================
# CONFIGURATION  —  edit these three lines
# ===========================================================================
PDF_INPUT   = r"C:\Users\YourName\Documents\FirstUnit1.pdf"
XLSX_INPUT  = r"C:\Users\YourName\Documents\DOC-20260626-WA0002_.xlsx"
PDF_OUTPUT  = r"C:\Users\YourName\Documents\FirstUnit1_Localized.pdf"

# Path to the Poppler "bin" folder you extracted (Step 3c above)
POPPLER_PATH = r"C:\poppler\Library\bin"

# Render DPI: 200 is a good balance of quality vs file size
RENDER_DPI = 200

# Column names in the Excel sheet (must match exactly)
COL_PAGE     = "Page No."
COL_ORIGINAL = "Original Text"
COL_SUGGEST  = "Suggested Localized Text"
# ===========================================================================


# ---------------------------------------------------------------------------
# Auto-detect a usable font on Windows (Arial preferred, fallbacks provided)
# ---------------------------------------------------------------------------
def _find_windows_font(bold: bool = False) -> str:
    """
    Returns the path to a suitable .ttf font file.
    Tries Arial first (ships with Windows), then common fallbacks.
    """
    windir = os.environ.get("WINDIR", r"C:\Windows")
    fonts_dir = os.path.join(windir, "Fonts")

    candidates = (
        ["arialbd.ttf", "calibrib.ttf", "trebucbd.ttf"]
        if bold else
        ["arial.ttf",   "calibri.ttf",  "trebuc.ttf"]
    )
    for name in candidates:
        path = os.path.join(fonts_dir, name)
        if os.path.exists(path):
            return path

    # Last resort: Pillow's built-in bitmap font (no TTF needed, but ugly)
    return None   # caller will use ImageFont.load_default()


FONT_REGULAR = _find_windows_font(bold=False)
FONT_BOLD    = _find_windows_font(bold=True)


# ---------------------------------------------------------------------------
# Step 1 — Load auditor changes from Excel
# ---------------------------------------------------------------------------
def load_changes(xlsx_path: str) -> list:
    """
    Returns a list of dicts:
        {"page": int, "original": str, "suggested": str}
    Rows where suggested text is NaN / identical to original are skipped.
    """
    df = pd.read_excel(xlsx_path)
    df.columns = [c.strip() for c in df.columns]   # normalise column names

    changes = []
    for _, row in df.iterrows():
        orig = str(row[COL_ORIGINAL]).strip() if pd.notna(row[COL_ORIGINAL]) else ""
        sugg = str(row[COL_SUGGEST]).strip()  if pd.notna(row[COL_SUGGEST])  else ""
        page = row[COL_PAGE]

        if not orig or not sugg or orig == sugg:
            continue
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
def render_pages(pdf_path: str, dpi: int = RENDER_DPI) -> list:
    """
    Render every page of the PDF to a high-resolution RGB PIL Image.
    Uses Poppler's pdftoppm under the hood via pdf2image.
    """
    kwargs = {"dpi": dpi}
    if POPPLER_PATH and os.path.isdir(POPPLER_PATH):
        kwargs["poppler_path"] = POPPLER_PATH

    images = convert_from_path(pdf_path, **kwargs)
    images = [img.convert("RGB") for img in images]
    print(f"[render_pages] {len(images)} page(s) rendered at {dpi} dpi")
    return images


# ---------------------------------------------------------------------------
# Step 3 — Extract word bounding boxes from the original PDF
# ---------------------------------------------------------------------------
def extract_words(pdf_path: str) -> dict:
    """
    Returns {page_number (1-based): [word_dict, ...]}
    Each word_dict has keys: text, x0, x1, top, bottom  (in PDF points)
    """
    result = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            result[i] = page.extract_words()
    return result


# ---------------------------------------------------------------------------
# Helper — coordinate conversion  (PDF pts -> image pixels)
# ---------------------------------------------------------------------------
def make_scaler(pdf_path: str, images: list):
    """
    Returns a callable  pdf_to_px(page_no, x0, top, x1, bottom)
    that converts PDF-point coordinates into image-pixel coordinates.
    The scale is computed per-page so it handles mixed page sizes correctly.
    """
    with pdfplumber.open(pdf_path) as pdf:
        page_dims = [(p.width, p.height) for p in pdf.pages]

    def pdf_to_px(page_no, x0, top, x1, bottom):
        pw, ph = page_dims[page_no - 1]
        iw, ih = images[page_no - 1].size
        sx, sy = iw / pw, ih / ph
        return (int(x0 * sx), int(top * sy), int(x1 * sx), int(bottom * sy))

    return pdf_to_px


# ---------------------------------------------------------------------------
# Step 4 — Locate original text and return its bounding box
# ---------------------------------------------------------------------------
def find_text_bbox(words: list, original: str):
    """
    Sliding-window token match across the page word list.

    Returns (x0, top, x1, bottom) in PDF pts covering all matched words,
    or None if no match is found.

    Algorithm:
    - Normalise whitespace in both the original string and each window.
    - Try the full token sequence first, then shorter leading fragments
      (down to 4 tokens) to handle text that wraps across lines differently
      between pdfplumber's extraction and the Excel sheet.
    """
    norm_original = " ".join(original.split())
    page_tokens   = [w["text"] for w in words]

    def bbox_of_slice(start, end):
        ws = words[start:end]
        return (
            min(w["x0"]     for w in ws),
            min(w["top"]    for w in ws),
            max(w["x1"]     for w in ws),
            max(w["bottom"] for w in ws),
        )

    orig_tokens = norm_original.split()
    min_tokens  = min(4, len(orig_tokens))

    for fragment_len in range(len(orig_tokens), min_tokens - 1, -1):
        target = " ".join(orig_tokens[:fragment_len])
        n      = fragment_len

        for start in range(len(page_tokens) - n + 1):
            window = " ".join(page_tokens[start : start + n])
            if window == target:
                return bbox_of_slice(start, start + n)

    return None


# ---------------------------------------------------------------------------
# Step 5 — Detect font style (size + bold) from matched region
# ---------------------------------------------------------------------------
def detect_font_style(words: list, bbox) -> tuple:
    """
    Returns (font_path, pt_size) by inspecting words within the bbox.

    Font size  : average word height (pdfplumber already returns pts).
    Bold flag  : True if any matched word's fontname contains "Bold".
    """
    x0, top, x1, bottom = bbox
    matched = [
        w for w in words
        if w["top"] >= top - 2 and w["bottom"] <= bottom + 2 and w["x0"] >= x0 - 2
    ]

    avg_height = 11.0
    if matched:
        avg_height = sum(w["bottom"] - w["top"] for w in matched) / len(matched)

    pt_size = round(avg_height)
    is_bold = any("Bold" in w.get("fontname", "") for w in matched if "fontname" in w)
    font_path = FONT_BOLD if is_bold else FONT_REGULAR
    return font_path, pt_size


# ---------------------------------------------------------------------------
# Step 6 — White-out region and draw replacement text
# ---------------------------------------------------------------------------
def apply_text_patch(
    draw,
    img_bbox,
    new_text,
    font_path,
    pt_size,
    scale_y,
    bg_color   = (255, 255, 255),
    text_color = (0, 0, 0),
    padding    = 4,
):
    """
    1. Erase the original text area with a white rectangle.
    2. Word-wrap the new text to fit exactly the same column width.
    3. Render each wrapped line with the detected font/size.
    """
    ix0, iy0, ix1, iy1 = img_bbox
    font_px = max(8, int(pt_size * scale_y))

    # Load font — fall back to Pillow default if TTF not found
    if font_path and os.path.exists(font_path):
        font = ImageFont.truetype(font_path, font_px)
    else:
        font = ImageFont.load_default()

    # Erase
    draw.rectangle(
        [ix0 - padding, iy0 - padding, ix1 + padding * 3, iy1 + padding * 3],
        fill=bg_color,
    )

    # Word-wrap
    max_width = (ix1 - ix0) + padding * 2
    words_    = new_text.split()
    lines, line = [], ""
    for word in words_:
        candidate = (line + " " + word).strip()
        if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
            line = candidate
        else:
            if line:
                lines.append(line)
            line = word
    if line:
        lines.append(line)

    # Draw
    line_height = int(font_px * 1.35)
    for i, ln in enumerate(lines):
        draw.text((ix0, iy0 + i * line_height), ln, fill=text_color, font=font)


# ---------------------------------------------------------------------------
# Step 7 — Orchestrate all changes across all pages
# ---------------------------------------------------------------------------
def apply_changes(images, changes, words_by_page, pdf_to_px, pdf_path):
    """Mutates `images` in-place, applying every row from the auditor sheet."""
    draws = [ImageDraw.Draw(img) for img in images]

    with pdfplumber.open(pdf_path) as pdf:
        page_heights = [p.height for p in pdf.pages]

    for ch in changes:
        page_no   = ch["page"]
        original  = ch["original"]
        suggested = ch["suggested"]

        if page_no < 1 or page_no > len(images):
            print(f"  [skip] page {page_no} — not in PDF (only {len(images)} page(s))")
            continue

        words = words_by_page.get(page_no, [])
        bbox  = find_text_bbox(words, original)

        if bbox is None:
            print(f"  [miss] page {page_no} — text not found: [{original[:60]}]")
            continue

        img_bbox   = pdf_to_px(page_no, *bbox)
        scale_y    = images[page_no - 1].size[1] / page_heights[page_no - 1]
        font_path, pt_size = detect_font_style(words, bbox)

        apply_text_patch(
            draw       = draws[page_no - 1],
            img_bbox   = img_bbox,
            new_text   = suggested,
            font_path  = font_path,
            pt_size    = pt_size,
            scale_y    = scale_y,
        )
        print(f"  [done] page {page_no}: [{original[:45]}]  ->  [{suggested[:45]}]")


# ---------------------------------------------------------------------------
# Step 8 — Assemble modified images back into a PDF
# ---------------------------------------------------------------------------
def save_pdf(images, output_path, pdf_input):
    """
    Embeds each modified page image into a new PDF at the original page size.
    Uses tempfile.gettempdir() for cross-platform temp file storage
    (works on both Windows and Linux/macOS).
    """
    with pdfplumber.open(pdf_input) as pdf:
        page_dims = [(p.width, p.height) for p in pdf.pages]

    tmp_dir   = tempfile.gettempdir()
    tmp_paths = []
    for i, img in enumerate(images):
        p = os.path.join(tmp_dir, f"_localized_page_{i+1}.png")
        img.save(p, "PNG")
        tmp_paths.append(p)

    c = rl_canvas.Canvas(output_path)
    for (pw, ph), tmp in zip(page_dims, tmp_paths):
        c.setPageSize((pw, ph))
        c.drawImage(tmp, 0, 0, width=pw, height=ph, preserveAspectRatio=False)
        c.showPage()
    c.save()

    for p in tmp_paths:
        os.remove(p)

    size_kb = os.path.getsize(output_path) // 1024
    print(f"[save_pdf] Written -> {output_path}  ({size_kb} KB)")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    # Allow overriding paths via command-line arguments
    global PDF_INPUT, XLSX_INPUT, PDF_OUTPUT
    if len(sys.argv) == 4:
        PDF_INPUT, XLSX_INPUT, PDF_OUTPUT = sys.argv[1], sys.argv[2], sys.argv[3]

    print("=== PDF Localizer (Windows) ===")
    print(f"  PDF   : {PDF_INPUT}")
    print(f"  XLSX  : {XLSX_INPUT}")
    print(f"  Output: {PDF_OUTPUT}")
    print(f"  Fonts : regular={FONT_REGULAR}  bold={FONT_BOLD}")
    print()

    changes = load_changes(XLSX_INPUT)
    if not changes:
        print("No changes to apply. Exiting.")
        sys.exit(0)

    images        = render_pages(PDF_INPUT)
    words_by_page = extract_words(PDF_INPUT)
    pdf_to_px     = make_scaler(PDF_INPUT, images)

    print("\nApplying changes:")
    apply_changes(images, changes, words_by_page, pdf_to_px, PDF_INPUT)

    print()
    save_pdf(images, PDF_OUTPUT, PDF_INPUT)
    print("\nDone.")


if __name__ == "__main__":
    main()
