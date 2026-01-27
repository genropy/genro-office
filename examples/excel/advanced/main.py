#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Advanced Excel spreadsheet example.

Demonstrates all advanced features:
- Cell formatting (colors, borders, alignment)
- Merged cells
- Freeze panes
- Autofilter
- Charts (bar, line, pie)
"""

from genro_office import ExcelApp


class AdvancedSpreadsheet(ExcelApp):
    """An advanced Excel spreadsheet with all features."""

    def recipe(self, root):
        wb = root.workbook()

        # Sheet 1: Formatted data with freeze panes and autofilter
        self._create_sales_sheet(wb)

        # Sheet 2: Charts
        self._create_charts_sheet(wb)

        # Sheet 3: Merged cells example
        self._create_merged_sheet(wb)

    def _create_sales_sheet(self, wb):
        """Create sales data sheet with formatting."""
        sheet = wb.sheet(name="Sales Data", freeze_panes="A2", autofilter="A1:E11")

        # Header row with formatting
        header = sheet.row(height=25)
        header.cell(
            content="Monthly Sales Report 2024",
            bold=True,
            font_size=16,
            bg_color="4472C4",
            font_color="FFFFFF",
            align="center",
            width=15,
        )
        header.cell(content="", bg_color="4472C4", width=12)
        header.cell(content="", bg_color="4472C4", width=12)
        header.cell(content="", bg_color="4472C4", width=12)
        header.cell(content="", bg_color="4472C4", width=12)

        # Merge the title (after the row is created)
        sheet.merge(range="A1:E1")

        # Column headers
        col_header = sheet.row(height=20)
        col_header.cell(content="Month", bold=True, bg_color="D9E2F3", border="thin", align="center")
        col_header.cell(content="North", bold=True, bg_color="D9E2F3", border="thin", align="center")
        col_header.cell(content="South", bold=True, bg_color="D9E2F3", border="thin", align="center")
        col_header.cell(content="East", bold=True, bg_color="D9E2F3", border="thin", align="center")
        col_header.cell(content="West", bold=True, bg_color="D9E2F3", border="thin", align="center")

        # Data rows
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
            row.cell(content=north, number_format="#,##0", border="thin", align="right")
            row.cell(content=south, number_format="#,##0", border="thin", align="right")
            row.cell(content=east, number_format="#,##0", border="thin", align="right")
            row.cell(content=west, number_format="#,##0", border="thin", align="right")

        # Totals row
        total_row = sheet.row(height=22)
        total_row.cell(content="TOTAL", bold=True, bg_color="FFC000", border="medium")
        total_row.cell(
            formula="=SUM(B3:B10)",
            bold=True,
            bg_color="FFC000",
            border="medium",
            number_format="#,##0",
        )
        total_row.cell(
            formula="=SUM(C3:C10)",
            bold=True,
            bg_color="FFC000",
            border="medium",
            number_format="#,##0",
        )
        total_row.cell(
            formula="=SUM(D3:D10)",
            bold=True,
            bg_color="FFC000",
            border="medium",
            number_format="#,##0",
        )
        total_row.cell(
            formula="=SUM(E3:E10)",
            bold=True,
            bg_color="FFC000",
            border="medium",
            number_format="#,##0",
        )

    def _create_charts_sheet(self, wb):
        """Create sheet with different chart types."""
        sheet = wb.sheet(name="Charts")

        # Data for charts
        header = sheet.row()
        header.cell(content="Category", bold=True, width=15)
        header.cell(content="Value", bold=True, width=12)

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

        # Bar chart
        sheet.chart(
            type="bar",
            title="Sales by Product (Bar)",
            data_range="B1:B6",
            categories_range="A2:A6",
            position="D2",
            width=12,
            height=8,
        )

        # Pie chart
        sheet.chart(
            type="pie",
            title="Sales Distribution (Pie)",
            data_range="B1:B6",
            categories_range="A2:A6",
            position="D14",
            width=12,
            height=8,
        )

        # Line chart data (separate section)
        for _ in range(15):
            sheet.row()

        # Monthly trend data
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

        # Line chart for trend
        sheet.chart(
            type="line",
            title="Monthly Trend (Line)",
            data_range=f"B{start_row}:C{start_row + 6}",
            categories_range=f"A{start_row + 1}:A{start_row + 6}",
            position="E23",
            width=14,
            height=8,
        )

    def _create_merged_sheet(self, wb):
        """Create sheet demonstrating merged cells and formatting."""
        sheet = wb.sheet(name="Merged Cells")

        # Row 1: Large merged header
        header = sheet.row(height=40)
        header.cell(
            content="Company Report 2024",
            bold=True,
            font_size=24,
            bg_color="1F4E79",
            font_color="FFFFFF",
            align="center",
            valign="center",
            width=20,
        )
        header.cell(content="", bg_color="1F4E79", width=15)
        header.cell(content="", bg_color="1F4E79", width=15)
        header.cell(content="", bg_color="1F4E79", width=15)
        sheet.merge(range="A1:D1")

        # Row 2: Subtitle
        subtitle = sheet.row(height=25)
        subtitle.cell(
            content="Quarterly Performance Summary",
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

        # Row 3: Empty
        sheet.row()

        # Rows 4-6: Q1 Data
        row4 = sheet.row()
        row4.cell(content="Q1 Results", bold=True, bg_color="D6DCE5", border="thin")
        row4.cell(content="Revenue", border="thin")
        row4.cell(content=1250000, border="thin", number_format="$#,##0")
        row4.cell(content="", border="thin")

        row5 = sheet.row()
        row5.cell(content="", bg_color="D6DCE5", border="thin")
        row5.cell(content="Expenses", border="thin")
        row5.cell(content=980000, border="thin", number_format="$#,##0")
        row5.cell(content="", border="thin")

        row6 = sheet.row()
        row6.cell(content="", bg_color="D6DCE5", border="thin")
        row6.cell(content="Profit", bold=True, border="thin")
        row6.cell(content=270000, bold=True, border="thin", number_format="$#,##0", bg_color="C6EFCE")
        row6.cell(content="", border="thin")

        # Merge Q1 label vertically
        sheet.merge(range="A4:A6")

        # Row 7: Empty
        sheet.row()

        # Rows 8-10: Q2 Data
        row8 = sheet.row()
        row8.cell(content="Q2 Results", bold=True, bg_color="D6DCE5", border="thin")
        row8.cell(content="Revenue", border="thin")
        row8.cell(content=1450000, border="thin", number_format="$#,##0")
        row8.cell(content="", border="thin")

        row9 = sheet.row()
        row9.cell(content="", bg_color="D6DCE5", border="thin")
        row9.cell(content="Expenses", border="thin")
        row9.cell(content=1100000, border="thin", number_format="$#,##0")
        row9.cell(content="", border="thin")

        row10 = sheet.row()
        row10.cell(content="", bg_color="D6DCE5", border="thin")
        row10.cell(content="Profit", bold=True, border="thin")
        row10.cell(content=350000, bold=True, border="thin", number_format="$#,##0", bg_color="C6EFCE")
        row10.cell(content="", border="thin")

        # Merge Q2 label vertically
        sheet.merge(range="A8:A10")


if __name__ == "__main__":
    spreadsheet = AdvancedSpreadsheet()
    spreadsheet.save("output.xlsx")
    print("Created: output.xlsx")
