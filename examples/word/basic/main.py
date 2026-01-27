#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Basic Word document example."""

from genro_office import WordApp


class BasicDocument(WordApp):
    """A simple Word document with headings, paragraphs, and a table."""

    def recipe(self, root):
        doc = root.document(title="My First Document")

        doc.heading(content="Introduction", level=1)
        doc.paragraph(content="This is a simple Word document created with genro-office.")
        doc.paragraph(content="It demonstrates the basic features of the WordBuilder.")

        doc.heading(content="Features", level=2)
        doc.paragraph(content="You can create headings at different levels.")
        doc.paragraph(content="You can add paragraphs with text content.")

        doc.heading(content="Sample Table", level=2)

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
    document.save("output.docx")
    print("Created: output.docx")
