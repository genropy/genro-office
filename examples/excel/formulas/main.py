#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Excel spreadsheet with formulas example.

Demonstrates formula support, cross-sheet references,
and ^path?attr data binding for labels.
"""

from genro_office import ExcelApp


class FormulaSpreadsheet(ExcelApp):
    """An Excel spreadsheet with formulas."""

    def main(self, source):
        wb = source.workbook()

        # Budget sheet with formulas
        sheet = wb.sheet(name="Budget")

        header = sheet.row(height=20.0)
        header.cell(content="Item", bold=True, width=25.0)
        header.cell(content="Quantity", bold=True, width=12.0)
        header.cell(content="Unit Price", bold=True, width=12.0)
        header.cell(content="Total", bold=True, width=15.0)

        items = [
            ("Laptops", 10, 1200),
            ("Monitors", 15, 350),
            ("Keyboards", 20, 75),
            ("Mice", 20, 25),
            ("Headsets", 10, 150),
        ]

        for i, (item, qty, price) in enumerate(items, start=2):
            row = sheet.row()
            row.cell(content=item)
            row.cell(content=qty)
            row.cell(content=price)
            row.cell(formula=f"=B{i}*C{i}")

        sheet.row()

        totals_row = sheet.row(height=25.0)
        totals_row.cell(content="TOTALS", bold=True)
        totals_row.cell(formula="=SUM(B2:B6)", bold=True)
        totals_row.cell(content="")
        totals_row.cell(formula="=SUM(D2:D6)", bold=True)

        # Statistics sheet
        stats = wb.sheet(name="Statistics")

        stats_header = stats.row(height=20.0)
        stats_header.cell(content="Metric", bold=True, width=20.0)
        stats_header.cell(content="Value", bold=True, width=15.0)

        metrics = [
            ("Total Items", "=SUM(Budget!B2:B6)"),
            ("Total Cost", "=SUM(Budget!D2:D6)"),
            ("Average Item Cost", "=AVERAGE(Budget!D2:D6)"),
            ("Max Item Cost", "=MAX(Budget!D2:D6)"),
            ("Min Item Cost", "=MIN(Budget!D2:D6)"),
        ]

        for metric, formula in metrics:
            row = stats.row()
            row.cell(content=metric)
            row.cell(formula=formula)

        stats.row()
        tax_row = stats.row()
        tax_row.cell(content="^tax?label", italic=True)
        tax_row.cell(formula="=SUM(Budget!D2:D6)*0.22")

        grand_total = stats.row(height=25.0)
        grand_total.cell(content="Grand Total", bold=True)
        grand_total.cell(formula="=SUM(Budget!D2:D6)*1.22", bold=True)


if __name__ == "__main__":
    spreadsheet = FormulaSpreadsheet()
    spreadsheet.data.set_item("tax", "", label="Tax (22%)")
    spreadsheet.build()
    spreadsheet.save("output.xlsx")
    print("Created: output.xlsx")
