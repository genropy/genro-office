#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Styled Excel spreadsheet example.

Demonstrates cell formatting: bold, italic, font sizes, row heights.
Uses ^path?attr for the report title.
"""

from genro_office import ExcelApp


class StyledSpreadsheet(ExcelApp):
    """An Excel spreadsheet with styled cells."""

    def main(self, source):
        wb = source.workbook()

        sheet = wb.sheet(name="Styled Report")

        # Title row with large bold font
        title_row = sheet.row(height=30.0)
        title_row.cell(
            content="^report?title", bold=True, font_size=18, width=25.0,
        )

        # Empty row for spacing
        sheet.row()

        # Header row
        header = sheet.row(height=20.0)
        header.cell(content="Category", bold=True, width=20.0)
        header.cell(content="Q1", bold=True, width=12.0)
        header.cell(content="Q2", bold=True, width=12.0)
        header.cell(content="Q3", bold=True, width=12.0)
        header.cell(content="Q4", bold=True, width=12.0)

        # Data rows
        categories = [
            ("Revenue", 100000, 120000, 115000, 140000),
            ("Expenses", 80000, 85000, 82000, 95000),
            ("Profit", 20000, 35000, 33000, 45000),
        ]

        for category, q1, q2, q3, q4 in categories:
            row = sheet.row()
            row.cell(content=category)
            row.cell(content=q1)
            row.cell(content=q2)
            row.cell(content=q3)
            row.cell(content=q4)

        sheet.row()

        notes_row = sheet.row()
        notes_row.cell(content="^report?notes", italic=True)

        # Second sheet
        sheet2 = wb.sheet(name="Team")

        header2 = sheet2.row(height=25.0)
        header2.cell(content="Team Members", bold=True, font_size=14, width=30.0)

        sheet2.row()

        members = [
            ("Alice", "Lead Developer", True),
            ("Bob", "Designer", False),
            ("Carol", "Project Manager", True),
            ("David", "Developer", False),
        ]

        for name, role, is_lead in members:
            row = sheet2.row()
            row.cell(content=name, bold=is_lead, width=15.0)
            row.cell(content=role, italic=not is_lead, width=20.0)


if __name__ == "__main__":
    spreadsheet = StyledSpreadsheet()
    spreadsheet.data.set_item(
        "report", "",
        title="Quarterly Report",
        notes="Note: All values in USD",
    )
    spreadsheet.build()
    spreadsheet.save("output.xlsx")
    print("Created: output.xlsx")
