#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Basic Word document example.

Demonstrates simple document creation with headings, paragraphs, and a table.
Uses ^path?attr data binding with set_item() for grouped attributes.
"""

from genro_office import WordApp


class BasicDocument(WordApp):
    """A simple Word document with headings, paragraphs, and a table."""

    def main(self, source):
        doc = source.document(title="^doc?title")

        doc.heading(content="^doc?section_intro", level=1)
        doc.paragraph(content="^content?intro")
        doc.paragraph(content="^content?features_intro")

        doc.heading(content="^doc?section_features", level=2)
        doc.paragraph(content="You can create headings at different levels.")
        doc.paragraph(content="You can add paragraphs with text content.")

        doc.heading(content="^doc?section_table", level=2)

        table = doc.table()
        row1 = table.row()
        row1.cell(content="Name")
        row1.cell(content="Value")

        row2 = table.row()
        row2.cell(content="Item 1")
        row2.cell(content="100")

        row3 = table.row()
        row3.cell(content="Item 2")
        row3.cell(content="200")


if __name__ == "__main__":
    document = BasicDocument()

    document.data.set_item(
        "doc", "",
        title="My First Document",
        section_intro="Introduction",
        section_features="Features",
        section_table="Sample Table",
    )

    document.data.set_item(
        "content", "",
        intro="This is a simple Word document created with genro-office.",
        features_intro="It demonstrates the basic features of the WordBuilder.",
    )

    document.build()
    document.save("output.docx")
    print("Created: output.docx")
