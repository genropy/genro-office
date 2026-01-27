#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Basic Excel spreadsheet example."""

from genro_office import ExcelApp


class BasicSpreadsheet(ExcelApp):
    """A simple Excel spreadsheet with basic data."""

    def recipe(self, root):
        wb = root.workbook()

        sheet = wb.sheet(name="Products")

        # Header row
        header = sheet.row()
        header.cell(content="Product", width=20)
        header.cell(content="Category", width=15)
        header.cell(content="Price", width=10)
        header.cell(content="Stock", width=10)

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
    spreadsheet.save("output.xlsx")
    print("Created: output.xlsx")
