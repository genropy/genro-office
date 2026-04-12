# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""genro-office: Reactive office document generation for Genropy.

Two apps for office documents with ^pointer data binding:

1. WordApp - Generate Word documents (.docx):
    ```python
    from genro_office import WordApp

    class MyReport(WordApp):
        def main(self, source):
            doc = source.document()
            doc.heading(content="^doc?title", level=1)
            doc.paragraph(content="^doc?body")

    report = MyReport()
    report.data.set_item("doc", "", title="Introduction", body="Hello World!")
    report.build()
    report.save("report.docx")
    ```

2. ExcelApp - Generate Excel spreadsheets (.xlsx):
    ```python
    from genro_office import ExcelApp

    class MySpreadsheet(ExcelApp):
        def main(self, source):
            wb = source.workbook()
            sheet = wb.sheet(name="Data")
            row = sheet.row()
            row.cell(content="^headers?col1")
            row.cell(content="^headers?col2")

    spreadsheet = MySpreadsheet()
    spreadsheet.data.set_item("headers", "", col1="Name", col2="Value")
    spreadsheet.build()
    spreadsheet.save("data.xlsx")
    ```
"""

__version__ = "0.6.0"

__all__ = [
    "__version__",
]

# Optional: WordBuilder, WordCompiler and WordApp (requires python-docx)
try:
    from genro_office.builders.word_builder import WordBuilder as WordBuilder
    from genro_office.compilers.word_compiler import WordCompiler as WordCompiler
    from genro_office.word_app import WordApp as WordApp

    __all__.extend(["WordApp", "WordBuilder", "WordCompiler"])
except ImportError:
    pass

# Optional: ExcelBuilder, ExcelCompiler and ExcelApp (requires openpyxl)
try:
    from genro_office.builders.excel_builder import ExcelBuilder as ExcelBuilder
    from genro_office.compilers.excel_compiler import ExcelCompiler as ExcelCompiler
    from genro_office.excel_app import ExcelApp as ExcelApp

    __all__.extend(["ExcelApp", "ExcelBuilder", "ExcelCompiler"])
except ImportError:
    pass
