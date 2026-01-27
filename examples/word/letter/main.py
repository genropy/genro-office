#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Business letter Word document example."""

from genro_office import WordApp


class LetterDocument(WordApp):
    """A formal business letter."""

    def recipe(self, root):
        doc = root.document()

        # Sender info
        doc.paragraph(content="Acme Corporation")
        doc.paragraph(content="123 Business Street")
        doc.paragraph(content="New York, NY 10001")
        doc.paragraph(content="")
        doc.paragraph(content="January 15, 2025")
        doc.paragraph(content="")

        # Recipient
        doc.paragraph(content="Mr. John Smith")
        doc.paragraph(content="ABC Company")
        doc.paragraph(content="456 Commerce Ave")
        doc.paragraph(content="Los Angeles, CA 90001")
        doc.paragraph(content="")

        # Salutation
        doc.paragraph(content="Dear Mr. Smith,")
        doc.paragraph(content="")

        # Body
        doc.paragraph(
            content="Thank you for your interest in our products. We are pleased to "
            "provide you with the requested information about our enterprise solutions."
        )
        doc.paragraph(content="")
        doc.paragraph(
            content="Our team has reviewed your requirements and we believe our "
            "Premium Package would be the best fit for your organization. This package "
            "includes all the features you mentioned, plus dedicated support."
        )
        doc.paragraph(content="")
        doc.paragraph(
            content="I would be happy to schedule a call to discuss this further. "
            "Please let me know your availability for next week."
        )
        doc.paragraph(content="")

        # Closing
        doc.paragraph(content="Sincerely,")
        doc.paragraph(content="")
        doc.paragraph(content="")
        doc.paragraph(content="Jane Doe")
        doc.paragraph(content="Sales Director")
        doc.paragraph(content="Acme Corporation")


if __name__ == "__main__":
    document = LetterDocument()
    document.save("output.docx")
    print("Created: output.docx")
