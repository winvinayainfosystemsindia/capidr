# PDF to Accessible Word (.docx) Converter

This tool uses the **Claude Opus API (`claude-opus-5`)** to convert any PDF document into an accessible, WCAG 2.1 AA / Section 508 / PDF-UA compliant Word (`.docx`) document.

---

## Table of Contents

1. [Prerequisites & Installation](#1-prerequisites--installation)
2. [API Key Configuration](#2-api-key-configuration)
3. [Input & Output File Locations](#3-input--output-file-locations)
4. [How to Run](#4-how-to-run)
5. [Advanced CLI Options](#5-advanced-cli-options)
6. [Remediation Standards Applied](#6-remediation-standards-applied)

---

## 1. Prerequisites & Installation

### Step 1: Ensure Python 3.10+ is installed
Make sure you have Python installed on your system.

### Step 2: Install required packages
Run the following command in your terminal/command prompt to install all dependencies from `requirements.txt`:

```bash
pip install -r requirements.txt
```

---

## 2. API Key Configuration

You need an **Anthropic API key** to run the Claude Opus model.

### Option A: Set Environment Variable (Recommended)

**PowerShell (Windows):**
```powershell
$env:ANTHROPIC_API_KEY="sk-ant-your-api-key-here"
```

**Command Prompt (Windows):**
```cmd
set ANTHROPIC_API_KEY=sk-ant-your-api-key-here
```

**Bash / Linux / macOS:**
```bash
export ANTHROPIC_API_KEY="sk-ant-your-api-key-here"
```

### Option B: Pass via Command Line Argument
You can pass the API key directly using the `--api-key` flag when executing the script.

---

## 3. Input & Output File Locations

### Input File Location
- Place your PDF file anywhere on your system or inside the project folder (e.g. `c:\External-projects\WinVinaya\capidr\sample.pdf`).
- Pass the path to your PDF file as the first positional argument to `pdf_to_word.py`.

### Output File Location
- **Default Location**: By default, the output `.docx` file will be generated in the **same folder** as the input PDF with the same file name (changing `.pdf` to `.docx`).
  - *Example*: `sample.pdf` $\rightarrow$ `sample.docx`
- **Custom Location**: You can optionally specify a custom output path as the second positional argument.

---

## 4. How to Run

### Basic Conversion (Default Output Path)
```bash
python pdf_to_word.py path/to/your_document.pdf
```

### Specify Custom Output Path
```bash
python pdf_to_word.py path/to/your_document.pdf path/to/output_document.docx
```

### Pass API Key Directly
```bash
python pdf_to_word.py sample.pdf --api-key sk-ant-your-api-key-here
```

---

## 5. Advanced CLI Options

| Flag | Description | Example |
|------|-------------|---------|
| `--save-json` | Saves the intermediate structured JSON alongside the output `.docx` file. | `python pdf_to_word.py sample.pdf --save-json` |
| `--from-json` | Builds a `.docx` file directly from a saved JSON file without making a new API call. | `python pdf_to_word.py sample.pdf output.docx --from-json sample.json` |
| `--json-only` | Outputs only the extracted JSON structure to console without creating `.docx`. | `python pdf_to_word.py sample.pdf --json-only` |
| `--verify` | Automatically verifies font consistency and heading markers (enabled by default). | `python pdf_to_word.py sample.pdf` |

---

## 6. Remediation Standards Applied

The document is formatted strictly according to the `pdf-accessibility-remediation` specification:

- **Heading 6 Page Markers**: Every PDF page is marked with a Heading 6 paragraph indicating the page label (e.g. `Page 1` or printed page number).
- **Page Breaks**: Added after each physical page (except the final page).
- **Headings**: Formatted as Heading 2, 3, or 4 based on document hierarchy.
- **Font & Size**: **Times New Roman 12pt** enforced for all text, headings, table cells, and captions.
- **Tables**: Formatted as real Word tables with header rows marked and **no merged cells**.
- **Figure Captions & Alt Text**: Sequential figure numbering (`Figure 1`, `Figure 2`, etc.) with accessibility Alt Text.
- **Footnotes & Endnotes**: Preserved at the bottom of pages or converted to endnotes according to the source PDF.
