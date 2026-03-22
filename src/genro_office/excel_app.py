# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""ExcelApp - Reactive app for Excel spreadsheets (.xlsx).

Uses BagAppBase pipeline with ^pointer data binding.
The recipe defines the spreadsheet template, data provides content.

Example::

    from genro_office import ExcelApp

    class MySpreadsheet(ExcelApp):
        def recipe(self, store):
            wb = store.workbook()
            sheet = wb.sheet(name="Data")
            row = sheet.row()
            row.cell(content="^headers?col1")
            row.cell(content="^headers?col2")

    spreadsheet = MySpreadsheet()
    spreadsheet.data.set_item("headers", "", col1="Name", col2="Value")
    spreadsheet.setup()
    spreadsheet.save("data.xlsx")
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, cast

from genro_builders import BagAppBase

if TYPE_CHECKING:
    from genro_bag import Bag

from genro_office.builders.excel_builder import ExcelBuilder
from genro_office.compilers.excel_compiler import ExcelCompiler


class ExcelApp(BagAppBase):
    """Reactive app for Excel spreadsheet generation.

    Extends BagAppBase with bytes output and live update support.
    """

    builder_class = ExcelBuilder
    compiler_class = ExcelCompiler
    _output: bytes | None = None  # type: ignore[assignment]

    @property
    def _excel_compiler(self) -> ExcelCompiler:
        """Return compiler cast to ExcelCompiler."""
        return cast("ExcelCompiler", self._compiler)

    def render(self, compiled_bag: Bag) -> bytes:  # type: ignore[override]
        """Render a CompiledBag to Excel workbook bytes.

        Args:
            compiled_bag: The compiled Bag (components expanded, pointers resolved).

        Returns:
            Excel workbook as bytes (.xlsx format).
        """
        return self._excel_compiler.render(compiled_bag)

    def save(self, filepath: str) -> None:
        """Save the spreadsheet to a file.

        Args:
            filepath: The path to save the spreadsheet to.
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

        compiler = self._excel_compiler
        if compiler.update_node(node):
            self._output = compiler.serialize()
            return

        self._rerender()
