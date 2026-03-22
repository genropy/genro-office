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
            row.cell(content="^headers.col1")
            row.cell(content="^headers.col2")

    spreadsheet = MySpreadsheet()
    spreadsheet.data["headers.col1"] = "Name"
    spreadsheet.data["headers.col2"] = "Value"
    spreadsheet.setup()
    spreadsheet.save("data.xlsx")
"""

from __future__ import annotations

from typing import Any, cast

from genro_builders import BagAppBase

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

    def compile(self) -> bytes:  # type: ignore[override]
        """Full pipeline: materialize -> bind -> render to bytes.

        Returns:
            Excel workbook as bytes (.xlsx format).
        """
        if self._compiler is None:
            msg = (
                f"{type(self).__name__} has no compiler. "
                f"Set compiler_class on the app or builder."
            )
            raise RuntimeError(msg)

        compiler = self._excel_compiler
        self._static_bag = compiler.preprocess(self._store)
        self._binding.bind(self._static_bag, self._data)
        self._output = compiler.compile_bound(self._static_bag)
        return self._output

    def render(self) -> bytes:
        """Render the spreadsheet to bytes. Alias for compile()."""
        return self.compile()

    def save(self, filepath: str) -> None:
        """Save the spreadsheet to a file.

        Args:
            filepath: The path to save the spreadsheet to.
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

        compiler = self._excel_compiler
        if compiler.update_node(node):
            self._output = compiler.serialize()
            return

        self._recompile()

    def _recompile(self) -> None:
        """Re-render the spreadsheet without re-materializing."""
        if self._compiler is not None and self._static_bag is not None:
            self._output = self._excel_compiler.compile_bound(self._static_bag)
