# WCAG 2.2 Remediated Document Validation Framework

An automated validation engine that checks remediated Word (`.docx`) documents against **WCAG 2.2 Level A / AA** accessibility standards and produces a comprehensive Excel report.

---

## Table of Contents

1. [Features](#features)
2. [Prerequisites & Installation](#prerequisites--installation)
3. [How to Run](#how-to-run)
4. [CLI Options & Flags](#cli-options--flags)
5. [Excel Report Structure](#excel-report-structure)
6. [Checklist Categories & Coverage](#checklist-categories--coverage)
7. [Architecture & Folder Structure](#architecture--folder-structure)
8. [Adding Custom Checkers](#adding-custom-checkers)

---

## Features

- **Automated Discovery**: 47 checkers across 12 distinct accessibility categories.
- **Smart N/A Handling**: Automatically flags non-applicable items as `N/A` (e.g. STEM equations in a non-STEM book, footnotes in a book without notes) rather than false failures.
- **Auto-Location Output**: By default, reports are saved directly in the same folder as the input Word document.
- **Two-Sheet Excel Report**:
  - **Summary Sheet**: Document metadata, overall pass rate %, status breakdown, and per-section statistics.
  - **Validation Details Sheet**: Granular row-by-row findings with exact locations, failure explanations, expected vs. actual values, conditional color formatting, and auto-filters.

---

## Prerequisites & Installation

Ensure you have Python 3.10+ installed and the required dependencies installed from the repository root:

```bash
pip install -r requirements.txt
```

*(Specifically requires `python-docx` and `openpyxl`)*

---

## How to Run

### 1. Basic Validation (Default Output Location)

The generated report is saved automatically in the **same directory** as the input document with the name `<document_name>_validation_report.xlsx`.

```bash
python -m validation "path/to/your_document.docx"
```

**Example:**
```bash
python -m validation "C:\Users\daran\Downloads\pdf\4.Effective TPD_DarlingHammond.docx"
```
*Output report:* `C:\Users\daran\Downloads\pdf\4.Effective TPD_DarlingHammond_validation_report.xlsx`

---

### 2. Custom Report Name

Provide a custom filename with `-o` or `--output`. Relative filenames will still automatically be saved in the input document's directory:

```bash
python -m validation "C:\Users\daran\Downloads\pdf\4.Effective TPD_DarlingHammond.docx" -o report.xlsx
```
*Output report:* `C:\Users\daran\Downloads\pdf\report.xlsx`

---

### 3. Verbose Mode (Detailed Console Progress)

Print real-time status of each checker as it executes:

```bash
python -m validation "path/to/document.docx" --verbose
```

---

### 4. Run Specific Category Checks

Filter execution to only run checks from specific sections:

```bash
python -m validation "path/to/document.docx" --checks "Tables" "Images & Graphics"
```

---

## CLI Options & Flags

| Flag | Short | Description | Example |
|---|---|---|---|
| `input` | *(Positional)* | Path to the remediated `.docx` file (Required). | `python -m validation sample.docx` |
| `--output` | `-o` | Output `.xlsx` filename or path (Defaults to `<input_dir>/<name>_validation_report.xlsx`). | `-o my_report.xlsx` |
| `--checks` | | Run only specific sections. | `--checks "Headings & Styles" "Tables"` |
| `--verbose` | | Print real-time execution details to console. | `--verbose` |
| `--help` | `-h` | Show help and available options. | `python -m validation --help` |

---

## Excel Report Structure

The generated Excel workbook contains two sheets:

### 1. Summary Sheet
- **Document Metadata**: File name, generation timestamp, total checks evaluated.
- **Pass Rate**: Color-coded percentage (calculated over applicable checks, excluding N/A).
- **Counts**: Total `PASS`, `FAIL`, and `N/A`.
- **Section Breakdown**: Subtotals for each of the 12 categories.

### 2. Validation Details Sheet
Contains 9 standard columns with frozen headers and auto-filters:

| Column | Description |
|---|---|
| **Section** | Category group (e.g. `Headings & Styles`, `Tables`) |
| **Checklist Item** | Specific item evaluated |
| **Description** | Human-readable requirement description |
| **WCAG 2.2 Criteria** | Relevant WCAG Success Criterion (e.g. `1.3.1 Info and Relationships (A)`) |
| **Status** | `PASS` (Green), `FAIL` (Red), or `N/A` (Gray) |
| **Failure Reason** | Detailed explanation of why the check failed |
| **Location** | Exact position in document (Page, Paragraph, Table, Row, Col, or Text snippet) |
| **Expected Value** | What the accessible standard requires |
| **Actual Value** | What was actually detected in the document |

---

## Checklist Categories & Coverage

| Category | Checkers Included | Primary WCAG Criteria |
|---|---|---|
| **Document Setup & Structure** | Language setting, page size & margins, orientation, blank pages, content in body, section & page breaks, logical reading order | 3.1.1, 1.3.4, 1.3.1, 1.3.2 |
| **Headings & Styles** | Heading styles applied, heading hierarchy (H1-H6, no skips), manual formatting detection, section bookmarks | 1.3.1, 2.4.6, 2.4.10, 2.4.1, 2.4.5 |
| **Lists & Numbering** | Built-in list styles (bullet/number), nesting indentation hierarchy, manual list markers detection | 1.3.1 |
| **Tables** | Real Word tables vs fake tabs, Repeat Header Row (`tblHeader`), column/row headers, merged/split cells, table captions, complex table summaries, STEM data tables | 1.1.1, 1.3.1, 1.3.2, 4.1.2, 2.4.6 |
| **Images & Graphics** | Informative alt text quality, decorative images, images-of-text avoidance, charts & data tables, diagram explanations | 1.1.1, 1.4.5 |
| **Equations (STEM)** | Equation Editor/MathType vs images, equation alt text, inline placement flow | 1.1.1, 4.1.2, 1.3.2 |
| **Symbols & Chemical Notations** | Unicode symbols vs Symbol fonts, chemical formula tagging, non-standard symbol definitions | 1.1.1, 1.3.1, 3.1.3, 4.1.2 |
| **Hyperlinks & Cross References** | Meaningful display text (no raw URLs), built-in cross references, destination link validation | 2.4.4, 2.4.5, 2.4.9 |
| **Footnotes & Endnotes** | Built-in footnote feature usage, marker linking, accessibility placement | 1.3.1, 1.3.2, 4.1.2, 2.4.3 |
| **Page Layout & Navigation** | Running page numbers, headers/footers, automatic Table of Contents, navigation bookmarks | 1.3.1, 2.4.1, 2.4.5 |
| **Metadata & Properties** | Title, Author, Subject metadata, personal info removal, document language property | 2.4.2, 3.1.1 |
| **Final QA & Accessibility Check** | Overall WCAG 2.2 AA compliance summary, reading order & navigation pane verification | All Level A & AA, 1.3.2, 2.4.3 |

---

## Architecture & Folder Structure

```
validation/
├── README.md                            # This file
├── __init__.py                          # Package root, exports run_validation()
├── __main__.py                          # Enables `python -m validation`
├── cli.py                               # CLI argument parser & entry point
├── models.py                            # CheckResult, ValidationReport dataclasses
├── base_check.py                        # BaseCheck abstract class
├── helpers.py                           # 20+ document parsing & XML helper utilities
├── runner.py                            # Dynamic checker discovery & execution orchestrator
├── report_generator.py                  # OpenPyXL-based styled report builder
│
└── checkers/                            # Category directories containing modular check files
    ├── document_setup/                  # 7 checks
    ├── headings_styles/                 # 4 checks
    ├── lists_numbering/                 # 3 checks
    ├── tables/                          # 7 checks
    ├── images_graphics/                 # 5 checks
    ├── equations/                       # 3 checks
    ├── symbols_notations/               # 3 checks
    ├── hyperlinks_crossrefs/            # 3 checks
    ├── footnotes_endnotes/              # 3 checks
    ├── page_layout_nav/                 # 4 checks
    ├── metadata_properties/             # 3 checks
    └── final_qa/                        # 2 checks
```

---

## Adding Custom Checkers

To add a new validation check:

1. Create a new `.py` file inside the appropriate `validation/checkers/<category>/` folder.
2. Inherit from `BaseCheck` and implement `run(self, doc, doc_path)`:

```python
from validation.base_check import BaseCheck
from validation.models import CheckResult

class CheckMyCustomRule(BaseCheck):
    section = "Headings & Styles"
    checklist_item = "Custom Rule Name"
    description = "Description of what this check verifies"
    wcag_criteria = "1.3.1 Info and Relationships (A)"

    def run(self, doc, doc_path: str) -> list[CheckResult]:
        # Perform validation logic
        if condition_met:
            return [self.pass_check(location="...", actual="...")]
        elif not_applicable:
            return [self.not_applicable("Reason for N/A")]
        else:
            return [self.fail_check(reason="...", location="...", expected="...", actual="...")]
```

*The runner will automatically discover and run the new checker without any manual registration.*
