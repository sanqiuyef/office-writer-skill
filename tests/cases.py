"""Fixed CLI cases shared by tests/test_scripts.py and tests/update_expected.py.

The scripts are executed with cwd=scripts/, so fixture paths are written as
``../tests/fixtures/...``. Those relative strings appear verbatim in tool output.
"""

CASES = [
    {
        "name": "list_color_schemes",
        "script": "ppt_list_color_schemes.py",
        "args": [],
        "exit_code": 0,
    },
    {
        "name": "read_word",
        "script": "read_word_document.py",
        "args": ["--file-path", "../tests/fixtures/sample.docx"],
        "exit_code": 0,
    },
    {
        "name": "read_excel",
        "script": "read_excel_workbook.py",
        "args": ["--file-path", "../tests/fixtures/sample.xlsx"],
        "exit_code": 0,
    },
    {
        "name": "read_powerpoint",
        "script": "read_powerpoint.py",
        "args": ["--file-path", "../tests/fixtures/sample.pptx"],
        "exit_code": 0,
    },
    {
        "name": "read_pdf",
        "script": "read_pdf_document.py",
        "args": ["--file-path", "../tests/fixtures/sample.pdf"],
        "exit_code": 0,
    },
    {
        "name": "read_pdf_structured",
        "script": "read_pdf_structured.py",
        "args": ["--file-path", "../tests/fixtures/sample.pdf"],
        "exit_code": 0,
    },
    {
        "name": "read_txt",
        "script": "read_txt_document.py",
        "args": ["--file-path", "../tests/fixtures/sample.txt"],
        "exit_code": 0,
    },
    {
        "name": "inspect_word",
        "script": "inspect_office_document.py",
        "args": ["--file-path", "../tests/fixtures/sample.docx"],
        "exit_code": 0,
    },
    {
        "name": "validate_word",
        "script": "validate_word_document.py",
        "args": ["--file-path", "../tests/fixtures/sample.docx"],
        "exit_code": 0,
    },
    {
        "name": "analyze_powerpoint",
        "script": "ppt_analyze_template.py",
        "args": ["--file-path", "../tests/fixtures/sample.pptx"],
        "exit_code": 0,
    },
    {
        "name": "validate_missing",
        "script": "validate_word_document.py",
        "args": ["--file-path", "../tests/fixtures/missing.docx"],
        "exit_code": 1,
    },
]
