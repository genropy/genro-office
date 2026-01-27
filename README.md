# genro-office

Office document generation for Genropy - Word and Excel builders using the App pattern.

## Installation

```bash
pip install genro-office

# With Word support
pip install genro-office[word]

# With Excel support
pip install genro-office[excel]

# With all features
pip install genro-office[all]
```

## Quick Start

### Word Documents

```python
from genro_office import WordApp


class MyReport(WordApp):
    def recipe(self, root):
        doc = root.document(title="Quarterly Report")

        doc.heading(content="Introduction", level=1)
        doc.paragraph(content="This is the quarterly report for Q4 2024.")

        doc.heading(content="Summary", level=2)
        doc.paragraph(
            content="Revenue increased by 15%.",
            bold=True,
            color="008000",
        )

        # Table
        table = doc.table(style="Table Grid")
        header = table.row()
        header.cell(content="Product", bold=True)
        header.cell(content="Sales", bold=True)

        table.row().cell(content="Widget A").cell(content="$10,000")
        table.row().cell(content="Widget B").cell(content="$15,000")


report = MyReport()
report.save("report.docx")
```

### Excel Spreadsheets

```python
from genro_office import ExcelApp


class SalesReport(ExcelApp):
    def recipe(self, root):
        wb = root.workbook()

        sheet = wb.sheet(name="Sales", freeze_panes="A2")

        # Header
        header = sheet.row(height=20)
        header.cell(content="Product", bold=True, bg_color="4472C4", width=20)
        header.cell(content="Q1", bold=True, bg_color="4472C4", width=12)
        header.cell(content="Q2", bold=True, bg_color="4472C4", width=12)

        # Data
        sheet.row().cell(content="Widget A").cell(content=1000).cell(content=1200)
        sheet.row().cell(content="Widget B").cell(content=800).cell(content=950)

        # Total with formula
        total = sheet.row()
        total.cell(content="Total", bold=True)
        total.cell(formula="=SUM(B2:B3)", bold=True)
        total.cell(formula="=SUM(C2:C3)", bold=True)


report = SalesReport()
report.save("sales.xlsx")
```

## Features

### WordApp / WordBuilder

- **Document settings**: orientation, margins
- **Headings**: levels 1-9 with formatting
- **Paragraphs**: bold, italic, underline, colors, fonts, alignment, spacing
- **Inline runs**: mixed formatting within paragraphs
- **Lists**: bulleted and numbered (`itemlist`/`item`)
- **Tables**: styled with cell formatting, colors, alignment
- **Images**: with width/height control
- **Headers and footers**
- **Page breaks**

### ExcelApp / ExcelBuilder

- **Multiple worksheets**
- **Cell formatting**: bold, italic, colors, borders, alignment, number formats
- **Column widths and row heights**
- **Merged cells**
- **Freeze panes**
- **Autofilter**
- **Charts**: bar, line, pie
- **Formulas**: including cross-sheet references

## Examples

See the `examples/` directory for complete examples:

```text
examples/
├── excel/
│   ├── advanced/   # All features: charts, merge, freeze panes
│   ├── basic/      # Simple data entry
│   ├── formulas/   # Excel formulas
│   └── styled/     # Formatting
└── word/
    ├── advanced/   # All features
    ├── basic/      # Headings, paragraphs, tables
    ├── letter/     # Business letter
    └── report/     # Quarterly report
```

Run any example:

```bash
cd examples/word/basic
python main.py
# Creates output.docx
```

## API Reference

### WordApp

```python
class MyDocument(WordApp):
    def recipe(self, root):
        doc = root.document(
            title="Document Title",
            orientation="portrait",  # or "landscape"
            margin_top=2.5,          # cm
            margin_bottom=2.5,
            margin_left=2.0,
            margin_right=2.0,
        )

        # Headings
        doc.heading(content="Title", level=1, bold=True, color="FF0000")

        # Paragraphs
        doc.paragraph(
            content="Text",
            bold=False,
            italic=False,
            underline=False,
            font_size=12,
            font_name="Arial",
            color="000000",
            align="left",  # left, center, right, justify
            space_before=12,  # points
            space_after=12,
            line_spacing=1.5,
        )

        # Inline runs
        para = doc.paragraph(content="Normal ")
        para.run(content="bold", bold=True)
        para.run(content=" and ")
        para.run(content="red", color="FF0000")

        # Lists
        bullet_list = doc.itemlist(type="bullet")  # or "number"
        bullet_list.item(content="First item")
        bullet_list.item(content="Second item")

        # Tables
        table = doc.table(style="Table Grid", align="center")
        row = table.row(height=1.5)  # cm
        row.cell(
            content="Header",
            bold=True,
            bg_color="4472C4",
            align="center",
            valign="center",  # top, center, bottom
            width=5,  # cm
        )

        # Header/Footer
        header = doc.header()
        header.paragraph(content="Header text", align="right")

        footer = doc.footer()
        footer.paragraph(content="Page 1", align="center")

        # Images
        doc.image(path="logo.png", width=2, height=1, align="center")

        # Page break
        doc.pagebreak()
```

### ExcelApp

```python
class MySpreadsheet(ExcelApp):
    def recipe(self, root):
        wb = root.workbook()

        sheet = wb.sheet(
            name="Data",
            freeze_panes="A2",      # Freeze at cell
            autofilter="A1:D10",    # Filterable range
        )

        # Rows
        row = sheet.row(height=25, hidden=False)

        # Cells
        row.cell(
            content="Value",
            formula="=SUM(A1:A10)",  # Formula takes precedence
            width=15,                # Column width
            bold=True,
            italic=False,
            font_size=12,
            font_color="FF0000",
            bg_color="FFFF00",
            align="center",          # left, center, right
            valign="center",         # top, center, bottom
            wrap_text=True,
            border="thin",           # thin, medium, thick, double
            border_color="000000",
            number_format="#,##0.00",
        )

        # Merged cells
        sheet.merge(range="A1:D1")

        # Charts
        sheet.chart(
            type="bar",              # bar, line, pie
            title="Sales Chart",
            data_range="B2:B10",
            categories_range="A2:A10",
            position="E2",
            width=15,                # cm
            height=10,
        )
```

## Development Status

**Alpha** - Core features implemented, API stabilizing.

## License

Apache License 2.0 - See [LICENSE](LICENSE) for details.

Copyright 2025 Softwell S.r.l.
