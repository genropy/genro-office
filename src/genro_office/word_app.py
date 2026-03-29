# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""WordApp — reactive app for Word documents (.docx).

Subclass WordApp and override ``recipe()`` to define the document
template. Data binding with ``^pointer`` is fully supported.

Example::

    from genro_office import WordApp

    class MyReport(WordApp):
        def recipe(self, source):
            doc = source.document(title="^doc?title")
            doc.heading(content="^doc?heading", level=1)
            doc.paragraph(content="^doc?body")

    report = MyReport()
    report.data.set_item("doc", "", title="Report", heading="Intro", body="Hello.")
    report.build()
    report.save("report.docx")
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, cast

from genro_builders import BuilderManager

from genro_office.builders.word_builder import WordBuilder

if TYPE_CHECKING:
    from genro_office.compilers.word_compiler import WordCompiler


class WordApp(BuilderManager):
    """Reactive app for Word document generation.

    Subclass and override ``recipe(source)`` to define the template.
    Call ``build()`` to populate and render, then ``save()`` to write.
    """

    def __init__(self) -> None:
        self.builder = self.set_builder("main", WordBuilder)

    @property
    def output(self) -> bytes | None:
        """Last rendered output as bytes (.docx format)."""
        return self.builder._output  # type: ignore[no-any-return]

    @property
    def _word_compiler(self) -> WordCompiler:
        """Return compiler cast to WordCompiler."""
        return cast("WordCompiler", self.builder._compiler_instance)

    def recipe(self, source: Any) -> None:
        """Define the document template. Override in subclass.

        Args:
            source: The source BuilderBag to populate with elements.
        """

    def build(self) -> None:
        """Run the recipe and build the document."""
        self.recipe(self.builder.source)
        self.builder.build()

    def render(self, built_bag: Any) -> bytes:
        """Render the built Bag to Word document bytes.

        Args:
            built_bag: The built Bag to render.

        Returns:
            Word document as bytes (.docx format).
        """
        return self._word_compiler.render(built_bag)

    def save(self, filepath: str) -> None:
        """Save the document to a file.

        Args:
            filepath: The path to save the document to.
        """
        output = self.output
        if output is None:
            self.build()
            output = self.output
        with open(filepath, "wb") as f:
            f.write(output)  # type: ignore[arg-type]
