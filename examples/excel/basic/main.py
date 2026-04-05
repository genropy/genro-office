#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Basic Excel spreadsheet example.

Demonstrates simple spreadsheet creation with header and data rows.
Uses ^path?attr for the sheet name.
"""

from genro_office import ExcelApp


class BasicSpreadsheet(ExcelApp):
    """A simple Excel spreadsheet with basic data."""

    def main(self, source):
        wb = source.workbook()

        sheet = wb.sheet(name="^config?sheet_name")

        # Header row
        header = sheet.row()
        header.cell(content="Product", width=20.0)
        header.cell(content="Category", width=15.0)
        header.cell(content="Price", width=10.0)
        header.cell(content="Stock", width=10.0)

        # Data rows
        products = [
            ("Laptop Pro", "Electronics", 1299.99, 50),
            ("Wireless Mouse", "Electronics", 29.99, 200),
            ("Office Chair", "Furniture", 249.99, 30),
            ("Desk Lamp", "Furniture", 45.99, 75),
            ("Notebook Set", "Office", 12.99, 500),
        ]

        for product, category, price, stock in products:
            row = sheet.row()
            row.cell(content=product)
            row.cell(content=category)
            row.cell(content=price)
            row.cell(content=stock)


if __name__ == "__main__":
    spreadsheet = BasicSpreadsheet()
    spreadsheet.data.set_item("config", "", sheet_name="Products")
    spreadsheet.build()
    spreadsheet.save("output.xlsx")
    print("Created: output.xlsx")
