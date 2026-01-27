# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""ExcelBuilder - Builder per spreadsheet Excel (.xlsx)."""

from __future__ import annotations

from io import BytesIO
from typing import TYPE_CHECKING, Any

from genro_bag import Bag
from genro_bag.builder import BagBuilderBase, element

if TYPE_CHECKING:
    from genro_bag.bagnode import BagNode

try:
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, LineChart, PieChart, Reference
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import column_index_from_string, get_column_letter

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


class ExcelBuilder(BagBuilderBase):
    """Builder per spreadsheet Excel (.xlsx) usando openpyxl."""

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
        # Font
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: int | None = None,
        font_color: str | None = None,
        # Fill
        bg_color: str | None = None,
        # Alignment
        align: str | None = None,
        valign: str | None = None,
        wrap_text: bool = False,
        # Border
        border: str | None = None,
        border_color: str | None = None,
        # Number format
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

    # -------------------------------------------------------------------------
    # Compile: Bag → bytes (xlsx)
    # -------------------------------------------------------------------------

    def compile(self, bag: Bag) -> bytes:
        """Compile Bag to Excel workbook bytes."""
        wb = Workbook()
        # Remove default sheet, we'll create named ones
        if wb.active:
            wb.remove(wb.active)

        for node in bag:
            self._compile_node(node, wb)

        # If no sheets were created, ensure at least one exists
        if not wb.sheetnames:
            wb.create_sheet("Sheet1")

        buffer = BytesIO()
        wb.save(buffer)
        return buffer.getvalue()

    def _compile_node(self, node: BagNode, wb: Any) -> None:
        """Compile a single node."""
        tag = node.tag or ""

        compile_method = getattr(self, f"_compile_{tag}", None)
        if compile_method:
            compile_method(node, wb)

    def _compile_workbook(self, node: BagNode, wb: Any) -> None:
        """Compile workbook node."""
        if isinstance(node.value, Bag):
            for child in node.value:
                self._compile_node(child, wb)

    def _compile_sheet(self, node: BagNode, wb: Any) -> None:
        """Compile sheet node."""
        name = node.attr.get("name", "Sheet1")
        ws = wb.create_sheet(title=name)

        freeze_panes = node.attr.get("freeze_panes")
        if freeze_panes:
            ws.freeze_panes = freeze_panes

        autofilter = node.attr.get("autofilter")
        if autofilter:
            ws.auto_filter.ref = autofilter

        if isinstance(node.value, Bag):
            # Collect merge and chart nodes to process after rows
            merge_nodes: list[BagNode] = []
            chart_nodes: list[BagNode] = []

            # First pass: compile all rows
            row_idx = 1
            for child_node in node.value:
                if child_node.tag == "row":
                    self._compile_row(child_node, ws, row_idx)
                    row_idx += 1
                elif child_node.tag == "merge":
                    merge_nodes.append(child_node)
                elif child_node.tag == "chart":
                    chart_nodes.append(child_node)

            # Second pass: apply merges (after all cells exist)
            for merge_node in merge_nodes:
                self._compile_merge(merge_node, ws)

            # Third pass: add charts
            for chart_node in chart_nodes:
                self._compile_chart(chart_node, ws)

    def _compile_row(self, node: BagNode, ws: Any, row_idx: int) -> None:
        """Compile row node."""
        height = node.attr.get("height")
        if height:
            ws.row_dimensions[row_idx].height = height

        hidden = node.attr.get("hidden", False)
        if hidden:
            ws.row_dimensions[row_idx].hidden = True

        if isinstance(node.value, Bag):
            col_idx = 1
            for cell_node in node.value:
                if cell_node.tag == "cell":
                    self._compile_cell(cell_node, ws, row_idx, col_idx)
                    col_idx += 1

    def _compile_cell(self, node: BagNode, ws: Any, row_idx: int, col_idx: int) -> None:
        """Compile cell node."""
        content = node.attr.get("content", "")
        formula = node.attr.get("formula")
        width = node.attr.get("width")

        cell = ws.cell(row=row_idx, column=col_idx)

        # Formula has priority over content
        if formula:
            cell.value = formula
        else:
            cell.value = content

        # Column width (applies to entire column)
        if width:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # Font styling
        self._apply_font(node, cell)

        # Fill (background color)
        bg_color = node.attr.get("bg_color")
        if bg_color:
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")

        # Alignment
        self._apply_alignment(node, cell)

        # Border
        self._apply_border(node, cell)

        # Number format
        number_format = node.attr.get("number_format")
        if number_format:
            cell.number_format = number_format

    def _apply_font(self, node: BagNode, cell: Any) -> None:
        """Apply font styling to cell."""
        bold = node.attr.get("bold", False)
        italic = node.attr.get("italic", False)
        underline = node.attr.get("underline", False)
        font_size = node.attr.get("font_size")
        font_color = node.attr.get("font_color")

        if bold or italic or underline or font_size or font_color:
            cell.font = Font(
                bold=bold,
                italic=italic,
                underline="single" if underline else None,
                size=font_size,
                color=font_color,
            )

    def _apply_alignment(self, node: BagNode, cell: Any) -> None:
        """Apply alignment to cell."""
        align = node.attr.get("align")
        valign = node.attr.get("valign")
        wrap_text = node.attr.get("wrap_text", False)

        if align or valign or wrap_text:
            cell.alignment = Alignment(
                horizontal=align,
                vertical=valign,
                wrap_text=wrap_text,
            )

    def _apply_border(self, node: BagNode, cell: Any) -> None:
        """Apply border to cell."""
        border_style = node.attr.get("border")
        border_color = node.attr.get("border_color", "000000")

        if border_style:
            side = Side(style=border_style, color=border_color)
            cell.border = Border(left=side, right=side, top=side, bottom=side)

    def _compile_merge(self, node: BagNode, ws: Any) -> None:
        """Compile merge node."""
        cell_range = node.attr.get("range", "")
        if cell_range:
            ws.merge_cells(cell_range)

    def _compile_chart(self, node: BagNode, ws: Any) -> None:
        """Compile chart node."""
        chart_type = node.attr.get("type", "bar")
        title = node.attr.get("title")
        data_range = node.attr.get("data_range", "")
        categories_range = node.attr.get("categories_range")
        position = node.attr.get("position", "E1")
        width = node.attr.get("width", 15)
        height = node.attr.get("height", 10)

        if not data_range:
            return

        # Create chart based on type
        if chart_type == "line":
            chart = LineChart()
        elif chart_type == "pie":
            chart = PieChart()
        else:
            chart = BarChart()

        if title:
            chart.title = title

        chart.width = width
        chart.height = height

        # Parse data range
        data_ref = self._parse_range_reference(ws, data_range)
        if data_ref:
            chart.add_data(data_ref, titles_from_data=True)

        # Parse categories range
        if categories_range:
            cat_ref = self._parse_range_reference(ws, categories_range)
            if cat_ref:
                chart.set_categories(cat_ref)

        ws.add_chart(chart, position)

    def _parse_range_reference(self, ws: Any, range_str: str) -> Any:
        """Parse a range string into a Reference object."""
        # Simple parsing for ranges like "A1:A10" or "B2:D10"
        if ":" not in range_str:
            return None

        start, end = range_str.split(":")

        # Extract column letter and row number
        start_col = ""
        start_row = ""
        for char in start:
            if char.isalpha():
                start_col += char
            else:
                start_row += char

        end_col = ""
        end_row = ""
        for char in end:
            if char.isalpha():
                end_col += char
            else:
                end_row += char

        # Convert column letters to numbers
        min_col = column_index_from_string(start_col)
        max_col = column_index_from_string(end_col)
        min_row = int(start_row)
        max_row = int(end_row)

        return Reference(ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)
