#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Advanced Word document example.

Demonstrates all advanced features with ^path?attr data binding:
- Document settings (orientation, margins)
- Text formatting (bold, italic, underline, colors, fonts)
- Inline runs with mixed formatting
- Bulleted and numbered lists
- Formatted tables with colors and alignment
- Headers and footers
- Style parameters via data binding (^styles?attr)
"""

from genro_office import WordApp


class AdvancedDocument(WordApp):
    """An advanced Word document with all features."""

    def main(self, source):
        doc = source.document(
            title="^doc?title",
            orientation="portrait",
            margin_top=2.5,
            margin_bottom=2.5,
            margin_left=2.0,
            margin_right=2.0,
        )

        # Header
        header = doc.header()
        header.paragraph(
            content="^doc?header_text",
            italic=True,
            align="right",
            color="^styles?muted_color",
        )

        # Section 1: Text Formatting
        self._create_formatting_section(doc)

        # Section 2: Lists
        self._create_lists_section(doc)

        # Section 3: Tables
        self._create_tables_section(doc)

        # Footer
        footer = doc.footer()
        footer.paragraph(
            content="^doc?footer_text",
            align="center",
            font_size=10,
            color="^styles?muted_color",
        )

    def _create_formatting_section(self, doc):
        """Create section demonstrating text formatting."""
        doc.heading(content="1. Text Formatting", level=1)

        doc.paragraph(content="This is bold text.", bold=True)
        doc.paragraph(content="This is italic text.", italic=True)
        doc.paragraph(content="This is underlined text.", underline=True)

        doc.paragraph(
            content="This has multiple styles applied.",
            bold=True,
            italic=True,
            font_size=14,
            color="^styles?accent_color",
        )

        doc.heading(content="1.1 Alignment", level=2)
        doc.paragraph(content="Left aligned paragraph.", align="left")
        doc.paragraph(content="Center aligned paragraph.", align="center")
        doc.paragraph(content="Right aligned paragraph.", align="right")
        doc.paragraph(
            content="Justified text that demonstrates how text flows when "
            "the justify alignment option is used.",
            align="justify",
        )

        doc.heading(content="1.2 Mixed Inline Formatting", level=2)
        para = doc.paragraph(content="This paragraph has ")
        para.run(content="bold", bold=True)
        para.run(content=", ")
        para.run(content="italic", italic=True)
        para.run(content=", ")
        para.run(content="colored", color="FF0000")
        para.run(content=", and ")
        para.run(content="highlighted", highlight="yellow")
        para.run(content=" text inline.")

        doc.heading(content="1.3 Paragraph Spacing", level=2)
        doc.paragraph(
            content="This paragraph has extra space before it.",
            space_before=24.0,
        )
        doc.paragraph(
            content="This paragraph has extra space after it.",
            space_after=24.0,
        )
        doc.paragraph(
            content="This paragraph has 1.5 line spacing for better readability "
            "when you have multiple lines of text.",
            line_spacing=1.5,
        )

    def _create_lists_section(self, doc):
        """Create section demonstrating lists."""
        doc.heading(content="2. Lists", level=1)

        doc.heading(content="2.1 Bullet List", level=2)
        bullet_list = doc.itemlist(type="bullet")
        bullet_list.item(content="First bullet point")
        bullet_list.item(content="Second bullet point")
        bullet_list.item(content="Third bullet point with more text")

        doc.heading(content="2.2 Numbered List", level=2)
        num_list = doc.itemlist(type="number")
        num_list.item(content="First step")
        num_list.item(content="Second step")
        num_list.item(content="Third step")
        num_list.item(content="Fourth step")

    def _create_tables_section(self, doc):
        """Create section demonstrating tables."""
        doc.heading(content="3. Tables", level=1)

        doc.heading(content="3.1 Simple Table", level=2)
        table1 = doc.table(style="Table Grid")

        header_row = table1.row()
        header_row.cell(
            content="Product", bold=True, bg_color="4472C4", align="center",
        )
        header_row.cell(
            content="Price", bold=True, bg_color="4472C4", align="center",
        )
        header_row.cell(
            content="Quantity", bold=True, bg_color="4472C4", align="center",
        )

        data = [
            ("Widget A", "$10.00", "100"),
            ("Widget B", "$15.50", "75"),
            ("Widget C", "$8.25", "200"),
        ]

        for product, price, qty in data:
            row = table1.row()
            row.cell(content=product)
            row.cell(content=price, align="right")
            row.cell(content=qty, align="center")

        doc.heading(content="3.2 Formatted Table", level=2)
        table2 = doc.table(style="Table Grid", align="center")

        h_row = table2.row()
        h_row.cell(
            content="Category", bold=True, bg_color="1F4E79",
            align="center", valign="center", width=5.0,
        )
        h_row.cell(
            content="Q1", bold=True, bg_color="1F4E79",
            align="center", width=3.0,
        )
        h_row.cell(
            content="Q2", bold=True, bg_color="1F4E79",
            align="center", width=3.0,
        )
        h_row.cell(
            content="Total", bold=True, bg_color="1F4E79",
            align="center", width=3.0,
        )

        categories = [
            ("Revenue", "$50,000", "$55,000", "$105,000"),
            ("Expenses", "$35,000", "$38,000", "$73,000"),
            ("Profit", "$15,000", "$17,000", "$32,000"),
        ]

        for i, (cat, q1, q2, total) in enumerate(categories):
            row = table2.row()
            bg = "D6DCE5" if i % 2 == 0 else None
            row.cell(content=cat, bg_color=bg)
            row.cell(content=q1, align="right", bg_color=bg)
            row.cell(content=q2, align="right", bg_color=bg)
            row.cell(content=total, align="right", bold=True, bg_color=bg)


if __name__ == "__main__":
    document = AdvancedDocument()

    document.data.set_item(
        "doc", "",
        title="Advanced Document Features",
        header_text="Company Confidential",
        footer_text="Page 1 of 1",
    )

    document.data.set_item(
        "styles", "",
        muted_color="808080",
        accent_color="0000FF",
    )

    document.build()
    document.save("output.docx")
    print("Created: output.docx")
