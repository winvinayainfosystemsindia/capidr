"""
PDF → Accessible Word Document Converter using Claude Opus API

Usage:
    python pdf_to_word.py <input.pdf> [output.docx]

Requirements:
    pip install anthropic python-docx Pillow

Set your API key:
    set ANTHROPIC_API_KEY=sk-ant-...
"""

import sys
import os
import json
import base64
import re
import io
import argparse
from pathlib import Path

import anthropic
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
MODEL = "claude-opus-5"
FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(12)
MAX_TOKENS = 128_000  # Claude Opus 5 supports large output

# The skill content is embedded as the system prompt
SYSTEM_PROMPT = r"""
You are an expert document remediation assistant. Your task is to convert a PDF
document into a structured JSON representation that will be used to build an
accessible Word (.docx) document.

<skill>
# PDF → Accessible Word Remediation

A repeatable spec for turning a PDF into a Word file that meets WCAG 2.1 AA /
Section 508 / PDF-UA expectations, without altering the substance of the source.

## Rules

### Page-by-page assembly
For each physical PDF page, in order:
1. A Heading 6 paragraph containing the page's label — the printed page number
   if shown, otherwise "Page N" using PDF sequential position.
2. That page's content with proper formatting.
3. A page break — except after the very last page.

### Heading levels
- Heading 1: document title (once).
- Heading 2/3/4: section headings mapped to the source's hierarchy.

### Global formatting
- Font: Times New Roman, 12pt, for ALL text.
- Preserve heading hierarchy through bold/italics/case, not point size.

### Tables and figures
- Tables: real tables, header row marked, NO merged cells — split merged cells
  into repeated plain cells.
- Figures: sequential "Figure N" captions only for images that carry captions
  in the source. Don't invent captions for decorative images.
- Alt text: every meaningful image gets descriptive alt text. Decorative images
  get empty alt.

### Footnotes and endnotes
Match the source's own choice:
- Bottom-of-page notes → footnotes
- Collected end-of-document notes → body content with own heading

### Never fabricate
If content is illegible or unclear, note it rather than guessing.
</skill>

<output_format>
You MUST respond with a single JSON object. Do NOT include any text before or
after the JSON. The JSON structure is:

{
  "title": "Document Title",
  "pages": [
    {
      "page_label": "1",
      "elements": [
        {
          "type": "heading",
          "level": 2,
          "text": "Section Title"
        },
        {
          "type": "paragraph",
          "text": "Body text content...",
          "bold": false,
          "italic": false
        },
        {
          "type": "list_item",
          "text": "A bullet point",
          "level": 0,
          "ordered": false
        },
        {
          "type": "table",
          "caption": "Table 1: Description",
          "headers": ["Col1", "Col2", "Col3"],
          "rows": [
            ["cell1", "cell2", "cell3"]
          ]
        },
        {
          "type": "figure",
          "figure_number": 1,
          "caption": "Figure 1: Description of the figure",
          "alt_text": "Descriptive alt text for accessibility",
          "description": "Visual description of what the figure shows"
        },
        {
          "type": "footnote",
          "marker": "1",
          "text": "Footnote text"
        },
        {
          "type": "blockquote",
          "text": "Quoted text"
        },
        {
          "type": "code_block",
          "text": "code content"
        }
      ]
    }
  ],
  "endnotes": [
    {
      "marker": "1",
      "text": "Endnote text"
    }
  ]
}

Rules for the JSON output:
1. Every page in the PDF MUST have an entry in "pages" with the correct "page_label".
2. page_label should be the printed page number if visible, otherwise "Page N".
3. Heading levels: 1 = doc title, 2/3/4 = content headings. Do NOT use level 5 or 6 in elements (level 6 is reserved for page markers).
4. For tables, split any merged cells into repeated plain cells. Every row must have the same number of columns as headers.
5. For figures with captions in the source, include figure_number in sequence.
6. Include ALL text content — do not skip or summarize anything.
7. For footnotes at the bottom of pages, use type "footnote". For collected endnotes, use the "endnotes" array.
8. Preserve the exact text — this is remediation, not paraphrase.
9. Running headers/footers are NOT content — do not include repeated page furniture.
10. Mid-sentence page breaks: resolve hyphenated words to whole words, split at word boundary.
</output_format>
"""

USER_PROMPT = """Consider the attached PDF and convert it to a structured JSON format
following the output_format specification in your instructions.

Do all these steps precisely:
- Convert ALL content from the PDF — do not skip, summarize, or change any content.
- For each PDF page, create a page entry with the correct page_label.
- Format content headings as Heading 2, 3, and 4 based on the source hierarchy.
- Include and properly format all tables (without merged cells).
- If figures have captions, add figure captions with sequential Figure numbers (Figure 1, Figure 2, etc.).
- Provide appropriate Alt Text for all meaningful images.
- Handle footnotes/endnotes as they appear in the source PDF.
- Ensure page labels are in sequential order.
- Preserve ALL original content verbatim."""


# ---------------------------------------------------------------------------
# Claude API interaction
# ---------------------------------------------------------------------------

def upload_pdf_and_get_response(pdf_path: str, api_key: str) -> dict:
    """Upload PDF to Claude and get structured JSON response."""
    client = anthropic.Anthropic(api_key=api_key)

    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    file_size_mb = pdf_path.stat().st_size / (1024 * 1024)
    print(f"[INFO] PDF file: {pdf_path.name} ({file_size_mb:.1f} MB)")

    # Use Files API for larger files, base64 for smaller ones
    if file_size_mb > 20:
        print("[INFO] Using Files API for large file upload...")
        with open(pdf_path, "rb") as f:
            file_upload = client.files.upload(
                file=(pdf_path.name, f, "application/pdf")
            )
        content = [
            {
                "type": "document",
                "source": {
                    "type": "file",
                    "file_id": file_upload.id,
                },
            },
            {
                "type": "text",
                "text": USER_PROMPT,
            },
        ]
        print(f"[INFO] File uploaded. ID: {file_upload.id}")
    else:
        print("[INFO] Using base64 inline upload...")
        with open(pdf_path, "rb") as f:
            pdf_data = base64.standard_b64encode(f.read()).decode("utf-8")
        content = [
            {
                "type": "document",
                "source": {
                    "type": "base64",
                    "media_type": "application/pdf",
                    "data": pdf_data,
                },
            },
            {
                "type": "text",
                "text": USER_PROMPT,
            },
        ]

    print(f"[INFO] Sending request to {MODEL}...")
    print("[INFO] This may take several minutes for large documents...")

    message = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        system=SYSTEM_PROMPT,
        messages=[
            {
                "role": "user",
                "content": content,
            }
        ],
    )

    print(f"[INFO] Response received. Usage: {message.usage}")

    # Extract text from response
    response_text = ""
    for block in message.content:
        if hasattr(block, "text"):
            response_text += block.text

    # Check if response was truncated
    if message.stop_reason == "max_tokens":
        print("[WARNING] Response was truncated due to max_tokens limit.")
        print("[WARNING] The output document may be incomplete.")
        print("[WARNING] Consider splitting the PDF into smaller parts.")

    # Parse JSON from response
    json_data = extract_json(response_text)
    return json_data


def extract_json(text: str) -> dict:
    """Extract JSON object from Claude's response text."""
    # Try direct parse first
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Try to find JSON within markdown code blocks
    json_match = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1).strip())
        except json.JSONDecodeError:
            pass

    # Try to find first { ... } block
    brace_start = text.find("{")
    if brace_start != -1:
        depth = 0
        for i in range(brace_start, len(text)):
            if text[i] == "{":
                depth += 1
            elif text[i] == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(text[brace_start : i + 1])
                    except json.JSONDecodeError:
                        break

    raise ValueError(
        "Could not parse JSON from Claude's response.\n"
        f"Response preview: {text[:500]}..."
    )


# ---------------------------------------------------------------------------
# Word document builder
# ---------------------------------------------------------------------------

def set_cell_font(cell, text, bold=False):
    """Set font properties for a table cell."""
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.bold = bold
    # Set East Asian font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)


def add_paragraph_with_font(doc, text, style=None, bold=False, italic=False,
                            alignment=None):
    """Add a paragraph with consistent Times New Roman 12pt formatting."""
    para = doc.add_paragraph(style=style)
    if alignment:
        para.alignment = alignment

    run = para.add_run(text)
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    run.font.bold = bold
    run.font.italic = italic
    # Ensure East Asian font fallback
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), FONT_NAME)

    return para


def set_heading_font(paragraph):
    """Ensure heading paragraphs use Times New Roman 12pt."""
    for run in paragraph.runs:
        run.font.name = FONT_NAME
        run.font.size = FONT_SIZE
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.get_or_add_rFonts()
        rfonts.set(qn("w:eastAsia"), FONT_NAME)


def add_page_break(doc):
    """Add a page break to the document."""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    fld_break = parse_xml(f'<w:br {nsdecls("w")} w:type="page"/>')
    run._element.append(fld_break)


def add_table(doc, headers, rows, caption=None):
    """Add a formatted table to the document."""
    if caption:
        cap_para = add_paragraph_with_font(doc, caption, bold=True, italic=True)
        cap_para.space_after = Pt(6)

    num_cols = len(headers) if headers else (len(rows[0]) if rows else 0)
    if num_cols == 0:
        return

    # Determine total rows (header + data)
    total_rows = len(rows) + (1 if headers else 0)
    table = doc.add_table(rows=total_rows, cols=num_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set header row
    if headers:
        hdr_row = table.rows[0]
        for i, header_text in enumerate(headers):
            if i < len(hdr_row.cells):
                set_cell_font(hdr_row.cells[i], header_text, bold=True)
        # Mark header row to repeat on new pages
        trpr = hdr_row._element.get_or_add_trPr()
        tbl_header = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
        trpr.append(tbl_header)

    # Fill data rows
    data_start = 1 if headers else 0
    for row_idx, row_data in enumerate(rows):
        actual_row_idx = row_idx + data_start
        if actual_row_idx < len(table.rows):
            row = table.rows[actual_row_idx]
            for col_idx, cell_text in enumerate(row_data):
                if col_idx < len(row.cells):
                    set_cell_font(row.cells[col_idx], cell_text)

    # Add spacing after table
    doc.add_paragraph()


def add_footnote_text(doc, marker, text):
    """Add footnote as formatted text at the bottom (when not using Word's
    native footnote mechanism for simplicity in API-generated docs)."""
    para = doc.add_paragraph()
    # Superscript marker
    marker_run = para.add_run(marker)
    marker_run.font.name = FONT_NAME
    marker_run.font.size = Pt(10)
    marker_run.font.superscript = True
    # Footnote text
    text_run = para.add_run(f" {text}")
    text_run.font.name = FONT_NAME
    text_run.font.size = FONT_SIZE


def build_docx(data: dict, output_path: str):
    """Build a Word document from the structured JSON data."""
    doc = Document()

    # -----------------------------------------------------------------------
    # Set default font for the document
    # -----------------------------------------------------------------------
    style = doc.styles["Normal"]
    font = style.font
    font.name = FONT_NAME
    font.size = FONT_SIZE

    # Set font for all heading styles used
    for level in range(1, 7):
        style_name = f"Heading {level}"
        if style_name in doc.styles:
            heading_style = doc.styles[style_name]
            heading_style.font.name = FONT_NAME
            heading_style.font.size = FONT_SIZE
            heading_style.font.color.rgb = RGBColor(0, 0, 0)  # Black text

    # Set list styles
    for style_name in ["List Bullet", "List Number"]:
        if style_name in doc.styles:
            list_style = doc.styles[style_name]
            list_style.font.name = FONT_NAME
            list_style.font.size = FONT_SIZE

    # -----------------------------------------------------------------------
    # Document title (Heading 1)
    # -----------------------------------------------------------------------
    title = data.get("title", "Untitled Document")
    title_para = doc.add_heading(title, level=1)
    set_heading_font(title_para)
    for run in title_para.runs:
        run.font.bold = True

    # -----------------------------------------------------------------------
    # Process pages
    # -----------------------------------------------------------------------
    pages = data.get("pages", [])
    total_pages = len(pages)

    for page_idx, page in enumerate(pages):
        page_label = page.get("page_label", str(page_idx + 1))

        # Add Heading 6 page marker
        page_marker = doc.add_heading(page_label, level=6)
        set_heading_font(page_marker)

        # Collect footnotes for this page
        page_footnotes = []

        # Process each element on the page
        elements = page.get("elements", [])
        for elem in elements:
            elem_type = elem.get("type", "paragraph")

            if elem_type == "heading":
                level = elem.get("level", 2)
                level = max(1, min(level, 4))  # Clamp to 1-4
                text = elem.get("text", "")
                heading_para = doc.add_heading(text, level=level)
                set_heading_font(heading_para)
                # Apply bold for visual hierarchy (all headings get bold)
                for run in heading_para.runs:
                    run.font.bold = True
                    if level >= 3:
                        run.font.italic = True

            elif elem_type == "paragraph":
                text = elem.get("text", "")
                if not text.strip():
                    doc.add_paragraph()
                    continue
                bold = elem.get("bold", False)
                italic = elem.get("italic", False)
                add_paragraph_with_font(doc, text, bold=bold, italic=italic)

            elif elem_type == "list_item":
                text = elem.get("text", "")
                ordered = elem.get("ordered", False)
                style_name = "List Number" if ordered else "List Bullet"
                add_paragraph_with_font(doc, text, style=style_name)

            elif elem_type == "table":
                headers = elem.get("headers", [])
                rows = elem.get("rows", [])
                caption = elem.get("caption", None)
                add_table(doc, headers, rows, caption)

            elif elem_type == "figure":
                fig_num = elem.get("figure_number", "")
                caption = elem.get("caption", "")
                alt_text = elem.get("alt_text", "")
                description = elem.get("description", "")

                # Add figure placeholder with caption
                if caption:
                    cap_text = caption
                else:
                    cap_text = f"Figure {fig_num}" if fig_num else "Figure"

                fig_para = add_paragraph_with_font(
                    doc,
                    f"[{cap_text}]",
                    italic=True,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER,
                )

                if alt_text:
                    alt_para = add_paragraph_with_font(
                        doc,
                        f"Alt Text: {alt_text}",
                        italic=True,
                    )
                    alt_para.paragraph_format.left_indent = Cm(1)

                if description:
                    desc_para = add_paragraph_with_font(
                        doc,
                        f"Image Description: {description}",
                        italic=True,
                    )
                    desc_para.paragraph_format.left_indent = Cm(1)

            elif elem_type == "footnote":
                marker = elem.get("marker", "")
                text = elem.get("text", "")
                page_footnotes.append((marker, text))

            elif elem_type == "blockquote":
                text = elem.get("text", "")
                bq_para = add_paragraph_with_font(doc, text, italic=True)
                bq_para.paragraph_format.left_indent = Cm(1.5)

            elif elem_type == "code_block":
                text = elem.get("text", "")
                code_para = add_paragraph_with_font(doc, text)
                code_para.paragraph_format.left_indent = Cm(1)

        # Add page footnotes at the bottom
        if page_footnotes:
            # Add a thin separator line
            sep_para = doc.add_paragraph()
            sep_run = sep_para.add_run("_" * 30)
            sep_run.font.name = FONT_NAME
            sep_run.font.size = Pt(8)

            for marker, fn_text in page_footnotes:
                add_footnote_text(doc, marker, fn_text)

        # Add page break (except after the last page)
        if page_idx < total_pages - 1:
            add_page_break(doc)

    # -----------------------------------------------------------------------
    # Endnotes (if any, as body content)
    # -----------------------------------------------------------------------
    endnotes = data.get("endnotes", [])
    if endnotes:
        add_page_break(doc)
        en_heading = doc.add_heading("Notes", level=2)
        set_heading_font(en_heading)

        for endnote in endnotes:
            marker = endnote.get("marker", "")
            text = endnote.get("text", "")
            en_para = doc.add_paragraph()
            marker_run = en_para.add_run(f"{marker}. ")
            marker_run.font.name = FONT_NAME
            marker_run.font.size = FONT_SIZE
            marker_run.font.superscript = True
            text_run = en_para.add_run(text)
            text_run.font.name = FONT_NAME
            text_run.font.size = FONT_SIZE

    # -----------------------------------------------------------------------
    # Final pass: ensure ALL paragraphs use Times New Roman 12pt
    # -----------------------------------------------------------------------
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = FONT_NAME
            run.font.size = FONT_SIZE
            rpr = run._element.get_or_add_rPr()
            rfonts = rpr.get_or_add_rFonts()
            rfonts.set(qn("w:eastAsia"), FONT_NAME)

    # Also fix table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = FONT_NAME
                        run.font.size = FONT_SIZE

    # -----------------------------------------------------------------------
    # Save
    # -----------------------------------------------------------------------
    doc.save(output_path)
    print(f"[SUCCESS] Word document saved to: {output_path}")


# ---------------------------------------------------------------------------
# Verification
# ---------------------------------------------------------------------------

def verify_document(doc_path: str):
    """Verify the generated document's structure."""
    doc = Document(doc_path)

    h6_labels = []
    heading_counts = {1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0}
    table_count = 0
    total_paragraphs = 0

    for para in doc.paragraphs:
        total_paragraphs += 1
        if para.style.name.startswith("Heading"):
            try:
                level = int(para.style.name.split()[-1])
                heading_counts[level] = heading_counts.get(level, 0) + 1
                if level == 6:
                    h6_labels.append(para.text)
            except ValueError:
                pass

    table_count = len(doc.tables)

    print("\n" + "=" * 60)
    print("DOCUMENT VERIFICATION REPORT")
    print("=" * 60)
    print(f"Total paragraphs: {total_paragraphs}")
    print(f"Tables: {table_count}")
    print(f"\nHeading breakdown:")
    for level, count in sorted(heading_counts.items()):
        if count > 0:
            print(f"  Heading {level}: {count}")

    print(f"\nHeading 6 (page markers): {len(h6_labels)}")
    if h6_labels:
        print(f"  First: {h6_labels[0]}")
        print(f"  Last:  {h6_labels[-1]}")
        print(f"  Sequence: {', '.join(h6_labels[:10])}", end="")
        if len(h6_labels) > 10:
            print(f" ... ({len(h6_labels)} total)")
        else:
            print()

    # Check font consistency
    non_tnr_count = 0
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.name and run.font.name != FONT_NAME:
                non_tnr_count += 1

    if non_tnr_count == 0:
        print("\n✓ All text uses Times New Roman")
    else:
        print(f"\n✗ {non_tnr_count} runs use a font other than Times New Roman")

    print("=" * 60)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to accessible Word document using Claude Opus API"
    )
    parser.add_argument("pdf_path", help="Path to the input PDF file")
    parser.add_argument(
        "output_path",
        nargs="?",
        default=None,
        help="Path for the output .docx file (default: same name as PDF with .docx extension)",
    )
    parser.add_argument(
        "--api-key",
        default=None,
        help="Anthropic API key (or set ANTHROPIC_API_KEY env var)",
    )
    parser.add_argument(
        "--json-only",
        action="store_true",
        help="Only output the JSON structure (don't build .docx)",
    )
    parser.add_argument(
        "--from-json",
        default=None,
        help="Build .docx from a previously saved JSON file (skip API call)",
    )
    parser.add_argument(
        "--save-json",
        action="store_true",
        help="Save the intermediate JSON to a file alongside the .docx",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        default=True,
        help="Verify the generated document structure (default: True)",
    )

    args = parser.parse_args()

    # Determine output path
    if args.output_path is None:
        args.output_path = str(Path(args.pdf_path).with_suffix(".docx"))

    # Get API key
    api_key = args.api_key or os.environ.get("ANTHROPIC_API_KEY")

    if args.from_json:
        # Build from existing JSON
        print(f"[INFO] Loading JSON from: {args.from_json}")
        with open(args.from_json, "r", encoding="utf-8") as f:
            data = json.load(f)
    else:
        if not api_key:
            print("ERROR: No API key provided.")
            print("Set ANTHROPIC_API_KEY environment variable or use --api-key flag.")
            sys.exit(1)

        # Call Claude API
        data = upload_pdf_and_get_response(args.pdf_path, api_key)

    if args.json_only:
        print(json.dumps(data, indent=2, ensure_ascii=False))
        return

    # Save intermediate JSON if requested
    if args.save_json:
        json_path = str(Path(args.output_path).with_suffix(".json"))
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"[INFO] JSON saved to: {json_path}")

    # Build the Word document
    print(f"[INFO] Building Word document...")
    build_docx(data, args.output_path)

    # Verify
    if args.verify:
        verify_document(args.output_path)


if __name__ == "__main__":
    main()
