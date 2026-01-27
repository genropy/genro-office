# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""ExcelApp - App per generare spreadsheet Excel (.xlsx)."""

from __future__ import annotations

from typing import cast

from genro_bag import Bag

from genro_office.builders.excel_builder import ExcelBuilder


class ExcelApp:
    """App per generare spreadsheet Excel (.xlsx).

    Esempio:
        ```python
        from genro_office import ExcelApp

        class MySpreadsheet(ExcelApp):
            def recipe(self, root):
                wb = root.workbook()
                sheet = wb.sheet(name="Dati")
                row = sheet.row()
                row.cell(content="Nome")
                row.cell(content="Valore")

        spreadsheet = MySpreadsheet()
        spreadsheet.save("dati.xlsx")
        ```
    """

    def __init__(self) -> None:
        self._page = Bag(builder=ExcelBuilder)
        self._data = Bag()
        self.recipe(self._page)

    @property
    def page(self) -> Bag:
        """The page Bag (workbook structure)."""
        return self._page

    @property
    def data(self) -> Bag:
        """The data Bag (for data binding)."""
        return self._data

    def recipe(self, root: Bag) -> None:
        """Override this method to build your spreadsheet.

        Args:
            root: The root Bag to add elements to.
        """

    def render(self) -> bytes:
        """Render the spreadsheet to bytes.

        Returns:
            The Excel workbook as bytes (.xlsx format).
        """
        builder = cast("ExcelBuilder", self._page.builder)
        return builder.compile(self._page)

    def save(self, filepath: str) -> None:
        """Save the spreadsheet to a file.

        Args:
            filepath: The path to save the spreadsheet to.
        """
        content = self.render()
        with open(filepath, "wb") as f:
            f.write(content)
