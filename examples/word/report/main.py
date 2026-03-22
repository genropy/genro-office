#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Business report Word document example.

Demonstrates a report template with ^path?attr data binding
for title, section headers, and paragraph content.
"""

from genro_office import WordApp


class ReportDocument(WordApp):
    """A business report with multiple sections and tables."""

    def recipe(self, store):
        doc = store.document(title="^report?title")

        doc.heading(content="^sections?summary", level=1)
        doc.paragraph(content="^content?summary")

        doc.heading(content="^sections?financial", level=1)
        doc.paragraph(content="^content?financial")

        doc.heading(content="^sections?revenue", level=2)
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

        doc.heading(content="^sections?achievements", level=1)
        doc.paragraph(content="^achievements?item1")
        doc.paragraph(content="^achievements?item2")
        doc.paragraph(content="^achievements?item3")

        doc.heading(content="^sections?next_steps", level=1)
        doc.paragraph(content="^content?next_steps")


if __name__ == "__main__":
    report = ReportDocument()

    report.data.set_item(
        "report", "",
        title="Quarterly Business Report",
    )

    report.data.set_item(
        "sections", "",
        summary="Executive Summary",
        financial="Financial Overview",
        revenue="Revenue Breakdown",
        achievements="Key Achievements",
        next_steps="Next Steps",
    )

    report.data.set_item(
        "content", "",
        summary=(
            "This report provides an overview of the company's performance "
            "during Q4 2024. Key metrics show positive growth across all departments."
        ),
        financial=(
            "Revenue increased by 15% compared to the previous quarter, "
            "while operating costs remained stable."
        ),
        next_steps=(
            "Focus areas for Q1 2025 include expanding into Asian markets "
            "and launching the next generation of our flagship product."
        ),
    )

    report.data.set_item(
        "achievements", "",
        item1="1. Launched new product line in European market",
        item2="2. Expanded customer support team by 25%",
        item3="3. Achieved 99.9% service uptime",
    )

    report.setup()
    report.save("output.docx")
    print("Created: output.docx")
