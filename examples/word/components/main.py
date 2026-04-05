#!/usr/bin/env python3
# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Business letter with @component — reusable document blocks.

Demonstrates how to define reusable components that expand into
document structures. The same ``address_block`` component is used
twice with different data prefixes (sender vs recipient), proving
that pointer paths are relocatable.

Architecture:
    LetterBuilder  — extends WordBuilder, adds @component definitions
    LetterCompiler — extends WordCompiler, adds component transparency
    BusinessLetter — extends WordApp, uses components in main()
"""

from __future__ import annotations

from typing import Any

from genro_bag import Bag
from genro_builders.builder import component, element

from genro_office import WordApp
from genro_office.builders.word_builder import WordBuilder
from genro_office.compilers.word_compiler import WordCompiler

# ---------------------------------------------------------------------------
# Builder: vocabulary with components
# ---------------------------------------------------------------------------

class LetterBuilder(WordBuilder):
    """WordBuilder extended with letter-specific components."""

    # Redefine document to accept component tags as children
    @element(
        sub_tags=(
            "heading,paragraph,table,image,pagebreak,"
            "itemlist,header,footer,run,"
            "address_block,signature_block"
        ),
    )
    def document(
        self,
        title: str = "",
        orientation: str | None = None,
        margin_top: float | None = None,
        margin_bottom: float | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
    ) -> None:
        ...

    @component(sub_tags="", parent_tags="document")
    def address_block(self, comp: Any, prefix: str = "sender", **kwargs: Any) -> None:  # noqa: ARG002
        """Reusable address block: name, street, city, country.

        Args:
            comp: The component's internal Bag.
            prefix: Data path prefix for ^pointer resolution.
        """
        comp.paragraph(content=f"^{prefix}?name", bold=True)
        comp.paragraph(content=f"^{prefix}?street")
        comp.paragraph(content=f"^{prefix}?city")
        comp.paragraph(content=f"^{prefix}?country")

    @component(sub_tags="", parent_tags="document")
    def signature_block(self, comp: Any, prefix: str = "sender", **kwargs: Any) -> None:  # noqa: ARG002
        """Reusable signature block: spacing + signatory + role + company.

        Args:
            comp: The component's internal Bag.
            prefix: Data path prefix for ^pointer resolution.
        """
        comp.paragraph(content="")
        comp.paragraph(content="")
        comp.paragraph(content=f"^{prefix}?signatory")
        comp.paragraph(content=f"^{prefix}?role", italic=True)
        comp.paragraph(content=f"^{prefix}?company")


# ---------------------------------------------------------------------------
# Compiler: component transparency
# ---------------------------------------------------------------------------

class LetterCompiler(WordCompiler):
    """WordCompiler with component-transparent node walking."""

    def _build_node(self, node: Any, doc: Any) -> None:
        """Build a node — falls back to walking children for components."""
        tag = node.node_tag or ""

        build_method = self._custom_handlers.get(tag)
        if build_method is None:
            build_method = getattr(self, f"_build_{tag}", None)

        if build_method:
            build_method(node, doc)
        elif isinstance(node.value, Bag):
            # Component transparency: walk expanded children
            for child in node.value:
                self._build_node(child, doc)


# Link compiler to builder
LetterBuilder._compiler_class = LetterCompiler


# ---------------------------------------------------------------------------
# App: the letter template
# ---------------------------------------------------------------------------

class BusinessLetter(WordApp):
    """A business letter using reusable address and signature components."""

    def __init__(self) -> None:
        self.builder = self.set_builder("main", LetterBuilder)

    def main(self, source: Any) -> None:
        doc = source.document()

        # Sender address — uses address_block with prefix="sender"
        doc.address_block(prefix="sender")

        doc.paragraph(content="")
        doc.paragraph(content="^letter?date")
        doc.paragraph(content="")

        # Recipient address — SAME component, different data
        doc.address_block(prefix="recipient")

        doc.paragraph(content="")
        doc.paragraph(content="^letter?salutation")
        doc.paragraph(content="")

        # Body paragraphs
        doc.paragraph(content="^body?opening")
        doc.paragraph(content="")
        doc.paragraph(content="^body?main_text")
        doc.paragraph(content="")
        doc.paragraph(content="^body?closing_text")
        doc.paragraph(content="")

        # Closing and signature
        doc.paragraph(content="^letter?closing")
        doc.signature_block(prefix="sender")


if __name__ == "__main__":
    letter = BusinessLetter()

    letter.data.set_item(
        "sender", "",
        name="Jane Doe",
        street="123 Innovation Drive",
        city="San Francisco, CA 94105",
        country="United States",
        signatory="Jane Doe",
        role="VP of Engineering",
        company="Acme Technologies Inc.",
    )

    letter.data.set_item(
        "recipient", "",
        name="Mr. Robert Chen",
        street="456 Commerce Boulevard",
        city="London EC2A 1NT",
        country="United Kingdom",
    )

    letter.data.set_item(
        "letter", "",
        date="April 5, 2025",
        salutation="Dear Mr. Chen,",
        closing="Kind regards,",
    )

    letter.data.set_item(
        "body", "",
        opening=(
            "Thank you for your interest in our platform. Following our "
            "recent meeting, I am pleased to outline our partnership proposal."
        ),
        main_text=(
            "Our engineering team has reviewed your technical requirements "
            "and confirmed full compatibility with your existing infrastructure. "
            "The integration timeline we discussed remains on track for Q3 delivery."
        ),
        closing_text=(
            "I look forward to discussing the next steps at your convenience. "
            "Please do not hesitate to reach out if you have any questions."
        ),
    )

    letter.build()
    letter.save("output.docx")
    print("Created: output.docx")
