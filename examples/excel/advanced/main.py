#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Advanced Excel spreadsheet example.

Demonstrates all advanced features with ^path?attr data binding:
- Cell formatting (colors, borders, alignment)
- Merged cells
- Freeze panes, autofilter
- Charts (bar, line, pie)
- Style binding via ^styles?attr
"""

from genro_office import ExcelApp


class AdvancedSpreadsheet(ExcelApp):
    """An advanced Excel spreadsheet with all features."""

    def recipe(self, store):
        wb = store.workbook()

        self._create_sales_sheet(wb)
        self._create_charts_sheet(wb)
        self._create_merged_sheet(wb)

    def _create_sales_sheet(self, wb):
        """Create sales data sheet with formatting."""
        sheet = wb.sheet(
            name="Sales Data", freeze_panes="A2", autofilter="A1:E11",
        )

        header = sheet.row(height=25.0)
        header.cell(
            content="^report?title",
            bold=True,
            font_size=16,
            bg_color="^styles?header_bg",
            font_color="^styles?header_fg",
            align="center",
            width=15.0,
        )
        header.cell(content="", bg_color="^styles?header_bg", width=12.0)
        header.cell(content="", bg_color="^styles?header_bg", width=12.0)
        header.cell(content="", bg_color="^styles?header_bg", width=12.0)
        header.cell(content="", bg_color="^styles?header_bg", width=12.0)

        sheet.merge(range="A1:E1")

        col_header = sheet.row(height=20.0)
        for label in ("Month", "North", "South", "East", "West"):
            col_header.cell(
                content=label, bold=True,
                bg_color="D9E2F3", border="thin", align="center",
            )

        sales_data = [
            ("January", 12500, 9800, 15600, 11200),
            ("February", 13200, 10100, 16200, 12100),
            ("March", 14100, 10800, 17100, 13000),
            ("April", 13800, 11200, 16800, 12800),
            ("May", 15200, 12000, 18200, 14200),
            ("June", 16100, 12800, 19100, 15100),
            ("July", 15800, 12500, 18800, 14800),
            ("August", 16500, 13200, 19500, 15500),
        ]

        for month, north, south, east, west in sales_data:
            row = sheet.row()
            row.cell(content=month, border="thin")
            row.cell(
                content=north, number_format="#,##0",
                border="thin", align="right",
            )
            row.cell(
                content=south, number_format="#,##0",
                border="thin", align="right",
            )
            row.cell(
                content=east, number_format="#,##0",
                border="thin", align="right",
            )
            row.cell(
                content=west, number_format="#,##0",
                border="thin", align="right",
            )

        total_row = sheet.row(height=22.0)
        total_row.cell(
            content="TOTAL", bold=True,
            bg_color="FFC000", border="medium",
        )
        for col in ("B", "C", "D", "E"):
            total_row.cell(
                formula=f"=SUM({col}3:{col}10)",
                bold=True,
                bg_color="FFC000",
                border="medium",
                number_format="#,##0",
            )

    def _create_charts_sheet(self, wb):
        """Create sheet with different chart types."""
        sheet = wb.sheet(name="Charts")

        header = sheet.row()
        header.cell(content="Category", bold=True, width=15.0)
        header.cell(content="Value", bold=True, width=12.0)

        data = [
            ("Product A", 150),
            ("Product B", 220),
            ("Product C", 180),
            ("Product D", 290),
            ("Product E", 210),
        ]

        for category, value in data:
            row = sheet.row()
            row.cell(content=category)
            row.cell(content=value)

        sheet.chart(
            type="bar",
            title="Sales by Product (Bar)",
            data_range="B1:B6",
            categories_range="A2:A6",
            position="D2",
            width=12.0,
            height=8.0,
        )

        sheet.chart(
            type="pie",
            title="Sales Distribution (Pie)",
            data_range="B1:B6",
            categories_range="A2:A6",
            position="D14",
            width=12.0,
            height=8.0,
        )

        for _ in range(15):
            sheet.row()

        trend_header = sheet.row()
        trend_header.cell(content="Month", bold=True)
        trend_header.cell(content="Sales", bold=True)
        trend_header.cell(content="Costs", bold=True)

        trend_data = [
            ("Jan", 100, 80),
            ("Feb", 120, 85),
            ("Mar", 115, 82),
            ("Apr", 130, 90),
            ("May", 145, 95),
            ("Jun", 160, 100),
        ]

        start_row = 23

        for month, sales, costs in trend_data:
            row = sheet.row()
            row.cell(content=month)
            row.cell(content=sales)
            row.cell(content=costs)

        sheet.chart(
            type="line",
            title="Monthly Trend (Line)",
            data_range=f"B{start_row}:C{start_row + 6}",
            categories_range=f"A{start_row + 1}:A{start_row + 6}",
            position="E23",
            width=14.0,
            height=8.0,
        )

    def _create_merged_sheet(self, wb):
        """Create sheet demonstrating merged cells and formatting."""
        sheet = wb.sheet(name="Merged Cells")

        header = sheet.row(height=40.0)
        header.cell(
            content="^report?company_title",
            bold=True,
            font_size=24,
            bg_color="1F4E79",
            font_color="FFFFFF",
            align="center",
            valign="center",
            width=20.0,
        )
        header.cell(content="", bg_color="1F4E79", width=15.0)
        header.cell(content="", bg_color="1F4E79", width=15.0)
        header.cell(content="", bg_color="1F4E79", width=15.0)
        sheet.merge(range="A1:D1")

        subtitle = sheet.row(height=25.0)
        subtitle.cell(
            content="^report?subtitle",
            italic=True,
            font_size=14,
            bg_color="2E75B6",
            font_color="FFFFFF",
            align="center",
        )
        subtitle.cell(content="", bg_color="2E75B6")
        subtitle.cell(content="", bg_color="2E75B6")
        subtitle.cell(content="", bg_color="2E75B6")
        sheet.merge(range="A2:D2")

        sheet.row()

        self._create_quarter_block(sheet, "Q1 Results", 1250000, 980000, "A4:A6")
        sheet.row()
        self._create_quarter_block(sheet, "Q2 Results", 1450000, 1100000, "A8:A10")

    def _create_quarter_block(self, sheet, label, revenue, expenses, merge_range):
        """Create a quarter data block with merge."""
        profit = revenue - expenses

        row1 = sheet.row()
        row1.cell(content=label, bold=True, bg_color="D6DCE5", border="thin")
        row1.cell(content="Revenue", border="thin")
        row1.cell(content=revenue, border="thin", number_format="$#,##0")
        row1.cell(content="", border="thin")

        row2 = sheet.row()
        row2.cell(content="", bg_color="D6DCE5", border="thin")
        row2.cell(content="Expenses", border="thin")
        row2.cell(content=expenses, border="thin", number_format="$#,##0")
        row2.cell(content="", border="thin")

        row3 = sheet.row()
        row3.cell(content="", bg_color="D6DCE5", border="thin")
        row3.cell(content="Profit", bold=True, border="thin")
        row3.cell(
            content=profit, bold=True, border="thin",
            number_format="$#,##0", bg_color="C6EFCE",
        )
        row3.cell(content="", border="thin")

        sheet.merge(range=merge_range)


if __name__ == "__main__":
    spreadsheet = AdvancedSpreadsheet()

    spreadsheet.data.set_item(
        "report", "",
        title="Monthly Sales Report 2024",
        company_title="Company Report 2024",
        subtitle="Quarterly Performance Summary",
    )

    spreadsheet.data.set_item(
        "styles", "",
        header_bg="4472C4",
        header_fg="FFFFFF",
    )

    spreadsheet.setup()
    spreadsheet.save("output.xlsx")
    print("Created: output.xlsx")
