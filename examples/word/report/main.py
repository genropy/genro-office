#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Business report Word document example."""

from genro_office import WordApp


class ReportDocument(WordApp):
    """A business report with multiple sections and tables."""

    def recipe(self, root):
        doc = root.document(title="Quarterly Business Report")

        doc.heading(content="Executive Summary", level=1)
        doc.paragraph(
            content="This report provides an overview of the company's performance "
            "during Q4 2024. Key metrics show positive growth across all departments."
        )

        doc.heading(content="Financial Overview", level=1)
        doc.paragraph(
            content="Revenue increased by 15% compared to the previous quarter, "
            "while operating costs remained stable."
        )

        doc.heading(content="Revenue Breakdown", level=2)
        table = doc.table()

        header = table.row()
        header.cell(content="Department")
        header.cell(content="Q3 Revenue")
        header.cell(content="Q4 Revenue")
        header.cell(content="Growth")

        data = [
            ("Sales", "$1.2M", "$1.4M", "+16.7%"),
            ("Services", "$800K", "$920K", "+15.0%"),
            ("Licensing", "$400K", "$450K", "+12.5%"),
        ]

        for dept, q3, q4, growth in data:
            row = table.row()
            row.cell(content=dept)
            row.cell(content=q3)
            row.cell(content=q4)
            row.cell(content=growth)

        doc.heading(content="Key Achievements", level=1)
        doc.paragraph(content="1. Launched new product line in European market")
        doc.paragraph(content="2. Expanded customer support team by 25%")
        doc.paragraph(content="3. Achieved 99.9% service uptime")

        doc.heading(content="Next Steps", level=1)
        doc.paragraph(
            content="Focus areas for Q1 2025 include expanding into Asian markets "
            "and launching the next generation of our flagship product."
        )


if __name__ == "__main__":
    document = ReportDocument()
    document.save("output.docx")
    print("Created: output.docx")
