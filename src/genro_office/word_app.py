# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""WordApp - Reactive app for Word documents (.docx).

Uses BagAppBase pipeline with ^pointer data binding.
The recipe defines the document template, data provides content.

Example::

    from genro_office import WordApp

    class MyReport(WordApp):
        def recipe(self, store):
            doc = store.document(title="^doc.title")
            doc.heading(content="^sections.intro.title", level=1)
            doc.paragraph(content="^sections.intro.body")

    report = MyReport()
    report.data["doc.title"] = "Annual Report"
    report.data["sections.intro.title"] = "Introduction"
    report.data["sections.intro.body"] = "Welcome to the report."
    report.setup()
    report.save("report.docx")
"""

from __future__ import annotations

from typing import Any, cast

from genro_builders import BagAppBase

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

    def compile(self) -> bytes:  # type: ignore[override]
        """Full pipeline: materialize -> bind -> render to bytes.

        Returns:
            Word document as bytes (.docx format).
        """
        if self._compiler is None:
            msg = (
                f"{type(self).__name__} has no compiler. "
                f"Set compiler_class on the app or builder."
            )
            raise RuntimeError(msg)

        compiler = self._word_compiler
        self._static_bag = compiler.preprocess(self._store)
        self._binding.bind(self._static_bag, self._data)
        self._output = compiler.compile_bound(self._static_bag)
        return self._output

    def render(self) -> bytes:
        """Render the document to bytes. Alias for compile()."""
        return self.compile()

    def save(self, filepath: str) -> None:
        """Save the document to a file.

        Args:
            filepath: The path to save the document to.
        """
        content = self.render()
        with open(filepath, "wb") as f:
            f.write(content)

    def _on_node_updated(self, node: Any) -> None:
        """Called when a bound node is updated via data change.

        Tries live update first, falls back to full recompile.
        """
        if not self._auto_compile:
            return

        compiler = self._word_compiler
        if compiler.update_node(node):
            self._output = compiler.serialize()
            return

        self._recompile()

    def _recompile(self) -> None:
        """Re-render the document without re-materializing."""
        if self._compiler is not None and self._static_bag is not None:
            self._output = self._word_compiler.compile_bound(self._static_bag)
