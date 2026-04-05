# genro-office

Reactive office document generation for Genropy — Word and Excel with template-based data binding.

## Overview

genro-office uses the **genro-builders** architecture: the `main()` method defines a **document template** with `^pointer` placeholders, and the `data` Bag provides the actual content. Change the data → the document updates.

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
    def main(self, source):
        doc = source.document(title="^doc?title")
        doc.heading(content="^doc?section", level=1)
        doc.paragraph(content="^doc?body", bold=True, color="008000")

        table = doc.table(style="Table Grid")
        header = table.row()
        header.cell(content="Product", bold=True)
        header.cell(content="Sales", bold=True)

        row = table.row()
        row.cell(content="Widget A")
        row.cell(content="$10,000")


report = MyReport()
report.data.set_item(
    "doc", "",
    title="Quarterly Report",
    section="Introduction",
    body="Revenue increased by 15%.",
)
report.build()
report.save("report.docx")
```

### Excel Spreadsheets

```python
from genro_office import ExcelApp


class SalesReport(ExcelApp):
    def main(self, source):
        wb = source.workbook()
        sheet = wb.sheet(name="^config?sheet_name", freeze_panes="A2")

        header = sheet.row(height=20.0)
        header.cell(content="Product", bold=True, bg_color="4472C4", width=20.0)
        header.cell(content="Q1", bold=True, bg_color="4472C4", width=12.0)
        header.cell(content="Q2", bold=True, bg_color="4472C4", width=12.0)

        sheet.row().cell(content="Widget A").cell(content=1000).cell(content=1200)
        sheet.row().cell(content="Widget B").cell(content=800).cell(content=950)

        total = sheet.row()
        total.cell(content="Total", bold=True)
        total.cell(formula="=SUM(B2:B3)", bold=True)
        total.cell(formula="=SUM(C2:C3)", bold=True)


report = SalesReport()
report.data.set_item("config", "", sheet_name="Sales")
report.build()
report.save("sales.xlsx")
```

## Architecture

genro-office extends [genro-builders](https://github.com/genropy/genro-builders) `BuilderManager`:

```text
WordApp / ExcelApp (BuilderManager)
    │
    ├── main(source)      ← define template with ^pointers
    ├── data              ← bind content, styles, metadata
    ├── build()           ← run pipeline: store → main → build → render
    └── save(filepath)    ← write bytes to file
```

### Components

| Component                        | Role                                             |
| -------------------------------- | ------------------------------------------------ |
| **WordBuilder / ExcelBuilder**   | Define document schema (`@element` declarations) |
| **WordCompiler / ExcelCompiler** | Transform Bag → bytes (docx/xlsx)                |
| **WordApp / ExcelApp**           | User-facing app, extends `BuilderManager`        |

### Data Binding with `^pointer`

Two pointer forms are supported:

| Syntax       | Reads                                  | Example            |
| ------------ | -------------------------------------- | ------------------ |
| `^path`      | Value of the node at `path`            | `^invoice.title`   |
| `^path?attr` | Attribute `attr` of the node at `path` | `^sender?company`  |

#### Using `^path` (node values)

Each data key is a separate node:

```python
app.data["invoice.title"] = "Invoice #2024-001"
app.data["invoice.body"] = "Thank you for your business."
```

#### Using `^path?attr` with `set_item` (node attributes)

Group related data as attributes of a single node:

```python
class Invoice(WordApp):
    def main(self, source):
        doc = source.document()
        doc.heading(
            content="^invoice?title",
            font_name="^styles?heading_font",
            color="^styles?heading_color",
        )
        doc.paragraph(content="^invoice?body")

app = Invoice()
app.data.set_item(
    "invoice", "",
    title="Invoice #2024-001",
    body="Thank you for your business.",
)
app.data.set_item(
    "styles", "",
    heading_font="Arial",
    heading_color="1F4E79",
)
app.build()
app.save("invoice.docx")
```

The `set_item(path, value, **kwargs)` method creates a node at `path` with `value` and keyword arguments as attributes. The `^path?attr` syntax reads the attribute from that node.

### Live Update

When data changes after `build()`, the compiler tries a **live update** on the Document/Workbook object. If the changed attribute supports it (font, color, content, bold, etc.), only that element is updated. Otherwise, the full document is re-rendered.

### Reusable Components (`@component`)

Define reusable document blocks with `@component` on a custom builder subclass.
Components are parameterized structures that expand into elements at build time.

```python
from genro_builders.builder import component, element

class LetterBuilder(WordBuilder):
    @element(sub_tags="heading,paragraph,table,image,pagebreak,itemlist,header,footer,run,address_block")
    def document(self, title="", orientation=None, margin_top=None,
                 margin_bottom=None, margin_left=None, margin_right=None): ...

    @component(sub_tags="", parent_tags="document")
    def address_block(self, comp, prefix="sender", **kwargs):
        comp.paragraph(content=f"^{prefix}?name", bold=True)
        comp.paragraph(content=f"^{prefix}?street")
        comp.paragraph(content=f"^{prefix}?city")
        comp.paragraph(content=f"^{prefix}?country")
```

The same component is reused with different data prefixes — the pointer paths are relocatable:

```python
class BusinessLetter(WordApp):
    def __init__(self):
        self.builder = self.set_builder("main", LetterBuilder)

    def main(self, source):
        doc = source.document()
        doc.address_block(prefix="sender")     # ^sender?name, ^sender?street, ...
        doc.address_block(prefix="recipient")  # ^recipient?name, ^recipient?street, ...
```

See `examples/word/components/` and `examples/excel/components/` for complete examples.

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
- **Reusable components**: `@component` for parameterized document blocks
- **`^pointer` data binding** on all attributes

### ExcelApp / ExcelBuilder

- **Multiple worksheets**
- **Cell formatting**: bold, italic, colors, borders, alignment, number formats
- **Column widths and row heights**
- **Merged cells**
- **Freeze panes**
- **Autofilter**
- **Charts**: bar, line, pie
- **Formulas**: including cross-sheet references
- **Reusable components**: `@component` for parameterized sheet blocks
- **`^pointer` data binding** on all attributes

## API Reference

### WordApp

```python
class MyDocument(WordApp):
    def main(self, source):
        doc = source.document(
            title="^doc?title",
            orientation="portrait",  # or "landscape"
            margin_top=2.5,          # cm
            margin_bottom=2.5,
            margin_left=2.0,
            margin_right=2.0,
        )

        # Headings
        doc.heading(content="^doc?heading", level=1, bold=True, color="FF0000")

        # Paragraphs
        doc.paragraph(
            content="^content?body",
            bold=False,
            italic=False,
            underline=False,
            font_size=12,
            font_name="^styles?body_font",
            color="000000",
            align="left",        # left, center, right, justify
            space_before=12.0,   # points
            space_after=12.0,
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
            width=5.0,        # cm
        )

        # Header/Footer
        header = doc.header()
        header.paragraph(content="^doc?header", align="right")

        footer = doc.footer()
        footer.paragraph(content="^doc?footer", align="center")

        # Images
        doc.image(path="logo.png", width=2.0, height=1.0, align="center")

        # Page break
        doc.pagebreak()

# Usage
app = MyDocument()
app.data.set_item(
    "doc", "",
    title="My Document",
    heading="Chapter 1",
    header="Confidential",
    footer="Page 1",
)
app.data.set_item("content", "", body="Hello World")
app.data.set_item("styles", "", body_font="Georgia")
app.build()
app.save("output.docx")
```

### ExcelApp

```python
class MySpreadsheet(ExcelApp):
    def main(self, source):
        wb = source.workbook()

        sheet = wb.sheet(
            name="^config?sheet_name",
            freeze_panes="A2",
            autofilter="A1:D10",
        )

        row = sheet.row(height=25.0, hidden=False)

        row.cell(
            content="^data?value",
            formula="=SUM(A1:A10)",  # formula takes precedence
            width=15.0,
            bold=True,
            italic=False,
            font_size=12,
            font_color="FF0000",
            bg_color="FFFF00",
            align="center",
            valign="center",
            wrap_text=True,
            border="thin",
            border_color="000000",
            number_format="#,##0.00",
        )

        sheet.merge(range="A1:D1")

        sheet.chart(
            type="bar",
            title="^config?chart_title",
            data_range="B2:B10",
            categories_range="A2:A10",
            position="E2",
            width=15.0,
            height=10.0,
        )

# Usage
app = MySpreadsheet()
app.data.set_item("config", "", sheet_name="Data", chart_title="Sales Chart")
app.data.set_item("data", "", value=42)
app.build()
app.save("output.xlsx")
```

## Examples

See the `examples/` directory for complete examples:

```text
examples/
├── excel/
│   ├── advanced/     # All features: charts, merge, freeze panes, ^pointer
│   ├── basic/        # Simple data entry with ^pointer
│   ├── components/   # @component: annual report with monthly blocks
│   ├── formulas/     # Excel formulas
│   └── styled/       # Formatting with ^pointer
└── word/
    ├── advanced/     # All features with ^pointer styles
    ├── basic/        # Headings, paragraphs, tables with ^pointer
    ├── components/   # @component: business letter with reusable blocks
    ├── letter/       # Business letter template with ^pointer
    └── report/       # Quarterly report template with ^pointer
```

Run any example:

```bash
cd examples/word/basic
python main.py
# Creates output.docx
```

## Development Status

**Alpha** - Core features implemented, API stabilizing.

## License

Apache License 2.0 - See [LICENSE](LICENSE) for details.

Copyright 2025 Softwell S.r.l.
