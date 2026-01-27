# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""WordApp - App per generare documenti Word (.docx)."""

from __future__ import annotations

from typing import cast

from genro_bag import Bag

from genro_office.builders.word_builder import WordBuilder


class WordApp:
    """App per generare documenti Word (.docx).

    Esempio:
        ```python
        from genro_office import WordApp

        class MyReport(WordApp):
            def recipe(self, root):
                doc = root.document(title="Report")
                doc.heading(content="Introduzione", level=1)
                doc.paragraph(content="Testo del paragrafo...")

        report = MyReport()
        report.save("report.docx")
        ```
    """

    def __init__(self) -> None:
        self._page = Bag(builder=WordBuilder)
        self._data = Bag()
        self.recipe(self._page)

    @property
    def page(self) -> Bag:
        """The page Bag (document structure)."""
        return self._page

    @property
    def data(self) -> Bag:
        """The data Bag (for data binding)."""
        return self._data

    def recipe(self, root: Bag) -> None:
        """Override this method to build your document.

        Args:
            root: The root Bag to add elements to.
        """

    def render(self) -> bytes:
        """Render the document to bytes.

        Returns:
            The Word document as bytes (.docx format).
        """
        builder = cast("WordBuilder", self._page.builder)
        return builder.compile(self._page)

    def save(self, filepath: str) -> None:
        """Save the document to a file.

        Args:
            filepath: The path to save the document to.
        """
        content = self.render()
        with open(filepath, "wb") as f:
            f.write(content)
