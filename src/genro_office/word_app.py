# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""WordApp - Reactive app for Word documents (.docx).

Uses BagAppBase pipeline with ^pointer data binding.
The recipe defines the document template, data provides content.

Example::

    from genro_office import WordApp

    class MyReport(WordApp):
        def recipe(self, store):
            doc = store.document(title="^doc?title")
            doc.heading(content="^doc?heading", level=1)
            doc.paragraph(content="^doc?body")

    report = MyReport()
    report.data.set_item("doc", "", title="Report", heading="Intro", body="Hello.")
    report.setup()
    report.save("report.docx")
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, cast

from genro_builders import BagAppBase

if TYPE_CHECKING:
    from genro_bag import Bag

from genro_office.builders.word_builder import WordBuilder
from genro_office.compilers.word_compiler import WordCompiler


class WordApp(BagAppBase):
    """Reactive app for Word document generation.

    Extends BagAppBase with bytes output and live update support.
    """

    builder_class = WordBuilder
    compiler_class = WordCompiler
    _output: bytes | None = None  # type: ignore[assignment]

    @property
    def _word_compiler(self) -> WordCompiler:
        """Return compiler cast to WordCompiler."""
        return cast("WordCompiler", self._compiler)

    def render(self, compiled_bag: Bag) -> bytes:  # type: ignore[override]
        """Render a CompiledBag to Word document bytes.

        Args:
            compiled_bag: The compiled Bag (components expanded, pointers resolved).

        Returns:
            Word document as bytes (.docx format).
        """
        return self._word_compiler.render(compiled_bag)

    def save(self, filepath: str) -> None:
        """Save the document to a file.

        Args:
            filepath: The path to save the document to.
        """
        if self._output is None:
            self.compile()
        with open(filepath, "wb") as f:
            f.write(self._output)  # type: ignore[arg-type]

    def _on_node_updated(self, node: Any) -> None:
        """Called when a bound node is updated via data change.

        Tries live update first, falls back to full re-render.
        """
        if not self._auto_compile:
            return

        compiler = self._word_compiler
        if compiler.update_node(node):
            self._output = compiler.serialize()
            return

        self._rerender()
