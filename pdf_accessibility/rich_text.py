"""Rich paragraph text: hyperlink detection and inline-equation rendering.

Any text passed through `add_formatted_text` is scanned for markdown links,
raw URLs, and `$...$` LaTeX equations, and each is converted into the
matching native Word construct (hyperlink run or OMML equation) in place.
"""

import html
import re

from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from .constants import FONT_NAME, FONT_SIZE
from .equations import insert_inline_equation


def add_hyperlink(paragraph, url, text, font_name=FONT_NAME, font_size_pt=12,
                  color="0002D0", underline=True, bold=False, italic=False):
    """Add an active, clickable Word hyperlink to a paragraph."""
    target_url = url
    if not target_url.startswith(("http://", "https://", "mailto:", "ftp://")):
        target_url = "https://" + target_url

    part = paragraph.part
    r_id = part.relate_to(target_url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} r:id="{r_id}" {nsdecls("r")}/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    rPr.append(parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>'))
    sz_val = int(font_size_pt * 2)
    rPr.append(parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_val}"/>'))

    if color:
        rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>'))
    if underline:
        rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>'))
    if bold:
        rPr.append(parse_xml(f'<w:b {nsdecls("w")}/>'))
    if italic:
        rPr.append(parse_xml(f'<w:i {nsdecls("w")}/>'))

    new_run.append(rPr)
    text_elem = parse_xml(f'<w:t {nsdecls("w")}>{html.escape(text)}</w:t>')
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_formatted_text(paragraph, text, bold=False, italic=False):
    """Add text to a paragraph, automatically detecting markdown links [text](url),
    raw URLs (http://, https://, www., mailto:), and inline LaTeX equations
    ($...$), inserting native Word hyperlinks and equations respectively."""
    if not text:
        return

    # Pattern matches markdown links `[display text](url)`, raw URLs
    # `https://...`, `http://...`, `www....`, `mailto:...`, or inline LaTeX
    # equations delimited by single dollar signs `$...$`.
    pattern = (
        r'(\[(?P<md_text>[^\]]+)\]\((?P<md_url>https?://[^\s)]+|www\.[^\s)]+|mailto:[^\s)]+)\))'
        r'|(?P<raw_url>https?://[^\s)]+|www\.[^\s)]+|mailto:[^\s)]+)'
        r'|\$(?P<equation>[^\s$](?:[^$]*[^\s$])?)\$'
    )

    last_idx = 0
    for match in re.finditer(pattern, text):
        start, end = match.span()
        # Add preceding normal text
        if start > last_idx:
            normal_text = text[last_idx:start]
            run = paragraph.add_run(normal_text)
            run.font.name = FONT_NAME
            run.font.size = FONT_SIZE
            run.font.bold = bold
            run.font.italic = italic
            run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT_NAME)

        equation = match.group("equation")
        if equation:
            insert_inline_equation(paragraph, equation)
            last_idx = end
            continue

        # Handle hyperlink match
        md_text = match.group("md_text")
        md_url = match.group("md_url")
        raw_url = match.group("raw_url")

        if md_text and md_url:
            link_text = md_text
            link_url = md_url
        else:
            link_text = raw_url
            link_url = raw_url

        add_hyperlink(paragraph, link_url, link_text, bold=bold, italic=italic)
        last_idx = end

    # Add remaining text after last match
    if last_idx < len(text):
        remaining_text = text[last_idx:]
        run = paragraph.add_run(remaining_text)
        run.font.name = FONT_NAME
        run.font.size = FONT_SIZE
        run.font.bold = bold
        run.font.italic = italic
        run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT_NAME)
