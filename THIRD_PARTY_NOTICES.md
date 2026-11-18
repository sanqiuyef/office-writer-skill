# Third-party notices

This project is released under the MIT License (see `LICENSE`). The items below
are not part of the MIT grant or carry their own terms.

## 1. Vendored engine

`scripts/office_engine/*.py` are vendored from **mcp-office-writer 0.6.0**
(2026-11), Copyright (c) 2026 sanqiuyef, under the MIT License. Local
modifications are listed in the first three lines of each file. The upstream
MCP server entry points (`__init__.py`, `__main__.py`) are not vendored; this
skill has no MCP runtime dependency.

## 2. Runtime dependencies (installed separately, not vendored)

| Package | License |
|:---|:---|
| python-docx | MIT |
| openpyxl | MIT |
| python-pptx | MIT |
| pydantic | MIT |
| PyMuPDF (fitz) | AGPL-3.0 or commercial (dual-licensed by Artifex) |
| fpdf2 | LGPL-3.0 |
| pywin32 | PSF License |
| pytesseract | Apache-2.0 |
| Pillow | HPND |
| pypdf | BSD-3-Clause |
| xlsxwriter | BSD-2-Clause |

PyMuPDF is AGPL-3.0/commercial dual-licensed. Review the AGPL obligations
(including network-use and source-offer terms) before redistributing a product
that bundles it.

Optional external tools, installed separately under their own licenses:

- **dsoffice** / `@deepseek-ai/libreoffice-kit` — used by `convert_to_pdf`,
  `render_document_pages`, `recalculate_excel_formulas`.
- **Microsoft Office / WPS Office COM** — used only when installed locally.
- **Tesseract OCR** — used for scanned-PDF text extraction.

## 3. Templates and examples

`templates/*.docx`, `templates/*.xlsx`, `templates/*.pptx` are starter files
generated for this skill and contain no third-party text.
`assets/examples/*.json` are synthetic examples written for this skill.
