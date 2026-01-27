# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""genro-office: Office document generation for Genropy.

Two apps for office documents:

1. WordApp - Generate Word documents (.docx):
    ```python
    from genro_office import WordApp

    class MyReport(WordApp):
        def recipe(self, root):
            doc = root.document(title="Report")
            doc.heading(content="Introduction", level=1)
            doc.paragraph(content="Hello World!")

    report = MyReport()
    report.save("report.docx")
    ```

2. ExcelApp - Generate Excel spreadsheets (.xlsx):
    ```python
    from genro_office import ExcelApp

    class MySpreadsheet(ExcelApp):
        def recipe(self, root):
            wb = root.workbook()
            sheet = wb.sheet(name="Data")
            row = sheet.row()
            row.cell(content="Name")
            row.cell(content="Value")

    spreadsheet = MySpreadsheet()
    spreadsheet.save("data.xlsx")
    ```
"""

__version__ = "0.1.0"

__all__ = [
    "__version__",
]

# Optional: WordBuilder and WordApp (requires python-docx)
try:
    from genro_office.builders.word_builder import WordBuilder as WordBuilder
    from genro_office.word_app import WordApp as WordApp

    __all__.extend(["WordApp", "WordBuilder"])
except ImportError:
    pass

# Optional: ExcelBuilder and ExcelApp (requires openpyxl)
try:
    from genro_office.builders.excel_builder import ExcelBuilder as ExcelBuilder
    from genro_office.excel_app import ExcelApp as ExcelApp

    __all__.extend(["ExcelApp", "ExcelBuilder"])
except ImportError:
    pass
