---
name: pdf-accessibility-remediation
description: Convert a source PDF into an accessible, WCAG 2.1 AA / Section 508 / PDF-UA aligned Word (.docx) document while preserving all content verbatim. Use whenever the user asks to "remediate" a PDF, convert a PDF into an "accessible Word document," fix accessibility issues in a scanned or exported document, or gives a page-by-page conversion spec involving Heading 6 page markers, page breaks per PDF page, figure captions, alt text, or footnote/endnote placement. Trigger this for course materials, reports, textbooks, or any scanned/typeset document that needs to become screen-reader friendly. Always confirm the source is one the requester has the rights or a legal basis to reproduce before running the full conversion — this skill is for producing an accessible copy of a document already lawfully held, not for generating full copies of arbitrary copyrighted works.
---

# PDF → Accessible Word Remediation

A repeatable spec for turning a PDF (textbook chapter, course document, journal
article, report) into a Word file that meets WCAG 2.1 AA / Section 508 /
PDF-UA expectations, without altering the substance of the source.

This skill defines the *rules*. Actual `.docx` file mechanics (tables, fonts,
footnotes, page breaks, verifying rendered output) should still go through the
`docx` skill — read `/mnt/skills/public/docx/SKILL.md` alongside this one.
For a document long enough to need iterative drafting, build it with `docx`
(npm) directly via a Node script rather than assembling paragraph-by-paragraph
in separate tool calls — it's far more reliable at this length.

## 0. Before you start

Confirm the source PDF is something the requester has the right to reproduce
in a new format — their own authored content, something their organization
licenses or owns, public-domain material, openly licensed material, or a
document they hold under a specific accessible-format exception (many
copyright regimes carve these out for organizations serving people with
disabilities — e.g. the Chafee Amendment in the US, the Marrakesh Treaty,
Section 52(1)(zb) of the Indian Copyright Act). Don't use this skill to
produce a full verbatim copy of an arbitrary copyrighted work (e.g. a
commercially published book or journal article) on the assumption that
"accessibility" alone licenses full reproduction — if that's unclear, ask
before converting the whole document.

## 1. Map the source structure first

Before writing anything, read the whole PDF and note:

- Total physical page count, and which pages carry a *printed* page number
  vs. which are unnumbered (covers, blank pages, dividers).
- The heading hierarchy actually present (don't invent levels that aren't
  there — a document with no subsections doesn't need Heading 4).
- Where footnotes live: bottom-of-page vs. a collected end-of-document notes
  list (see §6).
- Any tables, figures/images with captions, and any repeating running
  headers/footers.

Every PDF is structured differently — don't assume the next document follows
the same pattern as the last one.

## 2. Page-by-page assembly (the core pattern)

For each physical PDF page, in order:

1. A **Heading 6** paragraph containing that page's label — the printed page
   number if the page shows one (e.g. `261`), otherwise `Page N` using the
   PDF's own sequential position in the file (e.g. `Page 1` for an unnumbered
   cover). Never skip a page just because it's unnumbered — the point is a
   sequential, gapless landmark trail in the Navigation Pane.
2. That page's content (see §3–§7 for formatting rules).
3. A page break — *except* after the very last page.

After building the document, verify the Heading 6 trail: extract every
Heading 6 paragraph's text in document order and confirm it's sequential and
gapless relative to the source PDF's own pagination (e.g. via python-docx, or
by opening the rendered doc's Navigation Pane). This is worth checking
programmatically rather than eyeballing — it's easy to drop or duplicate a
page when transcribing a long document.

**Mid-sentence page breaks:** when a sentence is split across two PDF pages
by print pagination (often with a line-wrap hyphen, e.g. `"...forces be-"` /
`"come embodied..."`), don't preserve the hyphen as content — it's a print
artifact, not part of the text. Resolve it to the whole word (`become`),
and place the page break at that word boundary: the fragment before it closes
out the earlier page, the rest opens the next one, same paragraph either way.

**Running headers/footers are not content.** Repeated author names, article
titles, page numbers already captured by the Heading 6 marker, and download/
watermark stamps (e.g. "This content downloaded from...") don't get
re-typed into the body — they're page furniture, and repeating them is
exactly the "no repeated header/footer content" failure mode to avoid.

## 3. Heading levels

- **Heading 1**: the document's own title (once, where the source actually
  presents it as the title — not on a citation/cover page if the real title
  page comes later).
- **Heading 2 / 3 / 4**: content section headings, mapped to match the
  source's own visual/structural hierarchy (all-caps or bold section heads →
  Heading 2; italic or indented subsection heads → Heading 3; further nesting
  → Heading 4). Use only as many levels as the source actually has.

## 4. Global formatting

- Font: Times New Roman, 12pt, for **all** text — body, headings, footnotes,
  captions, table content. Preserve visual hierarchy for headings through
  bold / italics / case rather than point size, since size is fixed.
- Standard paragraph flow (no forced line breaks mid-paragraph); modest
  space-after between paragraphs reads better for screen-reader/low-vision
  use than first-line indent with no spacing.
- Epigraphs, verse, and block quotations: indented from the body margin,
  italic where the source itself uses italics for them.

## 5. Tables and figures

- **Tables**: real Word tables (not tab-stopped text), header row marked,
  **no merged cells** — split any merged header/cell from the source into
  repeated plain cells instead.
- **Figures**: add sequential "Figure N" captions *only* to images that
  actually carry a caption in the source. Don't invent captions for
  decorative or branding images (mastheads, logos, seals) — those either get
  skipped or marked as decorative, not numbered as figures.
- **Alt text**: every meaningful image needs a concise, descriptive alt text
  string. Purely decorative images should be marked decorative (empty alt),
  not given a fabricated description.

## 6. Equations

Every mathematical equation or formula gets transcribed as LaTeX and rendered
as a real, editable Word equation object (via `Insert > Equation` / OMML) —
never left as plain text, Unicode math approximations (`x²`, `√`), or a
screenshot. Match the source's own layout:

- **Standalone/display equations** (their own line, often numbered) become
  their own equation object, centered; if the source shows an equation
  number (e.g. `(1)`, `(3.2)`), it's right-aligned on the same line.
- **Inline equations** (embedded mid-sentence, in a list item, table cell, or
  footnote) render inline as a Word equation within the surrounding text,
  not pulled out onto their own line.

Transcribe exactly what the source shows — same variables, exponents,
subscripts, and symbols. This is remediation, not derivation: don't simplify,
re-derive, or "fix" an equation that looks unusual in the source.

## 7. Footnotes and endnotes — match the source's own choice

Don't default to one or the other — replicate what the PDF itself does:

- If the source places notes at the **bottom of the page** they reference,
  use real Word footnotes (`FootnoteReferenceRun` + a `footnotes` record in
  `docx`'s `Document` properties). Word anchors these to the correct page
  automatically, so they don't conflict with the manual page-break scheme
  in §2.
- If the source **collects notes at the end** of the document as a numbered
  list (e.g. an "Endnotes" or "Notes" section spanning its own pages), treat
  that section as ordinary body content — its own Heading 2, its own
  Heading 6 page markers, its own page breaks — rather than forcing Word's
  native endnote feature. Word's built-in endnotes always render in a single
  block at the very end of the file and ignore manual page breaks, which
  breaks the page-by-page fidelity this spec depends on. In-text reference
  markers in this case are plain superscript numbers, not linked fields.

## 8. Verify before delivering

```bash
python scripts/office/soffice.py --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 100 output.pdf page
```

Look at the first page, a middle page, and the last page. Check: font is
uniformly Times New Roman 12pt, page breaks land where expected, footnote (if
any) sits at the bottom of its page, tables have no merged cells, equations
render as real math objects (not literal `$...$` text or LaTeX source left
unconverted), and the Heading 6 sequence is intact. Fix and re-render before
presenting the file.

## 9. Never fabricate

If a page is illegible, a figure's caption is unclear, or the page-number
sequence in the source itself looks irregular, say so rather than guessing —
this is remediation, not paraphrase, so the transcription has to be exactly
right.
