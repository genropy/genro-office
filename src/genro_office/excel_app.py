# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""ExcelApp — reactive app for Excel spreadsheets (.xlsx).

Subclass ExcelApp and override ``recipe()`` to define the spreadsheet
template. Data binding with ``^pointer`` is fully supported.

Example::

    from genro_office import ExcelApp

    class MySpreadsheet(ExcelApp):
        def recipe(self, source):
            wb = source.workbook()
            sheet = wb.sheet(name="Data")
            row = sheet.row()
            row.cell(content="^headers?col1")
            row.cell(content="^headers?col2")

    spreadsheet = MySpreadsheet()
    spreadsheet.data.set_item("headers", "", col1="Name", col2="Value")
    spreadsheet.build()
    spreadsheet.save("data.xlsx")
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, cast

from genro_builders import BuilderManager

from genro_office.builders.excel_builder import ExcelBuilder

if TYPE_CHECKING:
    from genro_office.compilers.excel_compiler import ExcelCompiler


class ExcelApp(BuilderManager):
    """Reactive app for Excel spreadsheet generation.

    Subclass and override ``recipe(source)`` to define the template.
    Call ``build()`` to populate and render, then ``save()`` to write.
    """

    def __init__(self) -> None:
        self.builder = self.set_builder("main", ExcelBuilder)

    @property
    def output(self) -> bytes | None:
        """Last rendered output as bytes (.xlsx format)."""
        return self.builder._output  # type: ignore[no-any-return]

    @property
    def _excel_compiler(self) -> ExcelCompiler:
        """Return compiler cast to ExcelCompiler."""
        return cast("ExcelCompiler", self.builder._compiler_instance)

    def recipe(self, source: Any) -> None:
        """Define the spreadsheet template. Override in subclass.

        Args:
            source: The source BuilderBag to populate with elements.
        """

    def build(self) -> None:
        """Run the recipe and build the spreadsheet."""
        self.recipe(self.builder.source)
        self.builder.build()

    def render(self, built_bag: Any) -> bytes:
        """Render the built Bag to Excel workbook bytes.

        Args:
            built_bag: The built Bag to render.

        Returns:
            Excel workbook as bytes (.xlsx format).
        """
        return self._excel_compiler.render(built_bag)

    def save(self, filepath: str) -> None:
        """Save the spreadsheet to a file.

        Args:
            filepath: The path to save the spreadsheet to.
        """
        output = self.output
        if output is None:
            self.build()
            output = self.output
        with open(filepath, "wb") as f:
            f.write(output)  # type: ignore[arg-type]
