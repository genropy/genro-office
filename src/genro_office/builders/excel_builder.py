# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""ExcelBuilder - Builder for Excel spreadsheets (.xlsx).

Defines the schema elements for Excel spreadsheet generation.
Compilation is handled by ExcelCompiler.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from genro_builders import BagBuilderBase
from genro_builders.builder import element

if TYPE_CHECKING:
    from genro_bag import Bag

try:
    from openpyxl import Workbook as _Workbook  # noqa: F401

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


class ExcelBuilder(BagBuilderBase):
    """Builder for Excel spreadsheets (.xlsx) using openpyxl."""

    _compiler_class: type | None = None

    def __init__(self, bag: Bag) -> None:
        if not OPENPYXL_AVAILABLE:
            msg = "openpyxl required: pip install openpyxl"
            raise ImportError(msg)
        super().__init__(bag)

    # -------------------------------------------------------------------------
    # Element definitions
    # -------------------------------------------------------------------------

    @element(sub_tags="sheet")
    def workbook(self) -> None:
        """Root workbook element."""
        ...

    @element(sub_tags="row,merge,chart", parent_tags="workbook")
    def sheet(
        self,
        name: str = "Sheet1",
        freeze_panes: str | None = None,
        autofilter: str | None = None,
    ) -> None:
        """Worksheet element.

        Args:
            name: Sheet name.
            freeze_panes: Cell reference to freeze panes at (e.g., "A2", "B3").
            autofilter: Range for autofilter (e.g., "A1:D10").
        """
        ...

    @element(sub_tags="cell", parent_tags="sheet")
    def row(self, height: float | None = None, hidden: bool = False) -> None:
        """Row element.

        Args:
            height: Row height in points.
            hidden: Whether the row is hidden.
        """
        ...

    @element(sub_tags="", parent_tags="row")
    def cell(
        self,
        content: Any = "",
        width: float | None = None,
        formula: str | None = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: int | None = None,
        font_color: str | None = None,
        bg_color: str | None = None,
        align: str | None = None,
        valign: str | None = None,
        wrap_text: bool = False,
        border: str | None = None,
        border_color: str | None = None,
        number_format: str | None = None,
    ) -> None:
        """Cell element with full formatting support.

        Args:
            content: Cell value.
            width: Column width (applies to entire column).
            formula: Excel formula (e.g., "=SUM(A1:A10)").
            bold: Bold font.
            italic: Italic font.
            underline: Underline font.
            font_size: Font size in points.
            font_color: Font color as hex (e.g., "FF0000" for red).
            bg_color: Background color as hex (e.g., "FFFF00" for yellow).
            align: Horizontal alignment ("left", "center", "right").
            valign: Vertical alignment ("top", "center", "bottom").
            wrap_text: Whether to wrap text in cell.
            border: Border style ("thin", "medium", "thick", "double").
            border_color: Border color as hex.
            number_format: Excel number format (e.g., "#,##0.00", "0%", "yyyy-mm-dd").
        """
        ...

    @element(sub_tags="", parent_tags="sheet")
    def merge(self, range: str = "") -> None:
        """Merge cells element.

        Args:
            range: Cell range to merge (e.g., "A1:D1", "B2:B5").
        """
        ...

    @element(sub_tags="", parent_tags="sheet")
    def chart(
        self,
        type: str = "bar",
        title: str | None = None,
        data_range: str = "",
        categories_range: str | None = None,
        position: str = "E1",
        width: float = 15,
        height: float = 10,
    ) -> None:
        """Chart element.

        Args:
            type: Chart type ("bar", "line", "pie").
            title: Chart title.
            data_range: Data range (e.g., "B2:B10" or "B2:D10" for multiple series).
            categories_range: Categories/labels range (e.g., "A2:A10").
            position: Cell position for chart (e.g., "E1").
            width: Chart width in cm.
            height: Chart height in cm.
        """
        ...
