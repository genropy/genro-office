#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Business letter Word document example.

Demonstrates a formal letter template with ^pointer data binding.
Uses ^path?attr syntax to read attributes from data nodes,
and set_item() to group related data as node attributes.
"""

from genro_office import WordApp


class LetterDocument(WordApp):
    """A formal business letter template."""

    def recipe(self, store):
        doc = store.document()

        # Sender info (attributes of 'sender' node)
        doc.paragraph(content="^sender?company")
        doc.paragraph(content="^sender?address")
        doc.paragraph(content="^sender?city")
        doc.paragraph(content="")
        doc.paragraph(content="^letter?date")
        doc.paragraph(content="")

        # Recipient info (attributes of 'recipient' node)
        doc.paragraph(content="^recipient?name")
        doc.paragraph(content="^recipient?company")
        doc.paragraph(content="^recipient?address")
        doc.paragraph(content="^recipient?city")
        doc.paragraph(content="")

        # Salutation
        doc.paragraph(content="^letter?salutation")
        doc.paragraph(content="")

        # Body paragraphs (attributes of 'body' node)
        doc.paragraph(content="^body?para1")
        doc.paragraph(content="")
        doc.paragraph(content="^body?para2")
        doc.paragraph(content="")
        doc.paragraph(content="^body?para3")
        doc.paragraph(content="")

        # Closing
        doc.paragraph(content="^letter?closing")
        doc.paragraph(content="")
        doc.paragraph(content="")
        doc.paragraph(content="^sender?signatory")
        doc.paragraph(content="^sender?role")
        doc.paragraph(content="^sender?company")


if __name__ == "__main__":
    letter = LetterDocument()

    letter.data.set_item(
        "sender", "",
        company="Acme Corporation",
        address="123 Business Street",
        city="New York, NY 10001",
        signatory="Jane Doe",
        role="Sales Director",
    )

    letter.data.set_item(
        "recipient", "",
        name="Mr. John Smith",
        company="ABC Company",
        address="456 Commerce Ave",
        city="Los Angeles, CA 90001",
    )

    letter.data.set_item(
        "letter", "",
        date="January 15, 2025",
        salutation="Dear Mr. Smith,",
        closing="Sincerely,",
    )

    letter.data.set_item(
        "body", "",
        para1=(
            "Thank you for your interest in our products. We are pleased to "
            "provide you with the requested information about our enterprise solutions."
        ),
        para2=(
            "Our team has reviewed your requirements and we believe our "
            "Premium Package would be the best fit for your organization. This package "
            "includes all the features you mentioned, plus dedicated support."
        ),
        para3=(
            "I would be happy to schedule a call to discuss this further. "
            "Please let me know your availability for next week."
        ),
    )

    letter.setup()
    letter.save("output.docx")
    print("Created: output.docx")
