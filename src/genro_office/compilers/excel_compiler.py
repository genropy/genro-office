# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""ExcelCompiler - Compiler for Excel spreadsheets (.xlsx).

Transforms a Bag (built with ExcelBuilder) into an Excel workbook.
Maintains a live Workbook object for incremental updates.
"""

from __future__ import annotations

from io import BytesIO
from typing import TYPE_CHECKING, Any

from genro_bag import Bag
from genro_builders import BagCompilerBase
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

if TYPE_CHECKING:
    from genro_bag import BagNode


class ExcelCompiler(BagCompilerBase):
    """Compiler Excel: Bag -> bytes (.xlsx).

    Maintains a live Workbook and a _live_map {node_id: live_object}
    for incremental updates when data changes.
    """

    def __init__(self, builder: Any) -> None:
        super().__init__(builder)
        self._wb: Any | None = None
        self._live_map: dict[int, Any] = {}

    # -------------------------------------------------------------------------
    # Main compile entry points (override for bytes output)
    # -------------------------------------------------------------------------

    def compile(self, bag: Bag | None = None, data: Bag | None = None) -> bytes:  # type: ignore[override]
        """Compile bag to Excel workbook bytes.

        Args:
            bag: The Bag to compile. If None, uses builder.bag.
            data: Optional data Bag for pointer resolution.

        Returns:
            Excel workbook as bytes (.xlsx format).
        """
        if bag is None:
            bag = self.builder.bag

        processed_bag = self.preprocess(bag)

        if data is not None:
            self._resolve_pointers_inline(processed_bag, data)

        return self._render_to_bytes(processed_bag)

    def compile_bound(self, bound_bag: Bag) -> bytes:  # type: ignore[override]
        """Compile a pre-bound bag (app mode).

        Args:
            bound_bag: Bag with components expanded and pointers resolved.

        Returns:
            Excel workbook as bytes.
        """
        return self._render_to_bytes(bound_bag)

    def serialize(self) -> bytes:
        """Serialize the live Workbook to bytes without rebuilding."""
        if self._wb is None:
            return b""
        buffer = BytesIO()
        self._wb.save(buffer)
        return buffer.getvalue()

    # -------------------------------------------------------------------------
    # Live update
    # -------------------------------------------------------------------------

    def update_node(self, node: BagNode) -> bool:
        """Try to update the live object for a node.

        Returns True if the update was applied, False if a full
        recompile is needed.
        """
        live_obj = self._live_map.get(id(node))
        if live_obj is None:
            return False
        return self._apply_live_update(node, live_obj)

    def _apply_live_update(self, node: BagNode, live_cell: Any) -> bool:
        """Apply attribute changes to a live openpyxl cell."""
        tag = node.tag or ""

        if tag == "cell":
            formula = node.attr.get("formula")
            if formula:
                live_cell.value = formula
            else:
                live_cell.value = node.attr.get("content", "")

            self._apply_font(node, live_cell)
            self._apply_alignment(node, live_cell)
            self._apply_border(node, live_cell)

            bg_color = node.attr.get("bg_color")
            if bg_color:
                live_cell.fill = PatternFill(
                    start_color=str(bg_color), end_color=str(bg_color), fill_type="solid"
                )

            number_format = node.attr.get("number_format")
            if number_format:
                live_cell.number_format = str(number_format)

            return True

        return False

    # -------------------------------------------------------------------------
    # Workbook building (imperative walk)
    # -------------------------------------------------------------------------

    def _render_to_bytes(self, bag: Bag) -> bytes:
        """Build an Excel Workbook from a Bag and return bytes."""
        self._wb = Workbook()
        self._live_map.clear()

        if self._wb.active:
            self._wb.remove(self._wb.active)

        for node in bag:
            self._build_node(node, self._wb)

        if not self._wb.sheetnames:
            self._wb.create_sheet("Sheet1")

        buffer = BytesIO()
        self._wb.save(buffer)
        return buffer.getvalue()

    def _build_node(self, node: BagNode, wb: Any) -> None:
        """Build a single node into the Workbook."""
        tag = node.tag or ""

        build_method = getattr(self, f"_build_{tag}", None)
        if build_method:
            build_method(node, wb)

    def _build_workbook(self, node: BagNode, wb: Any) -> None:
        """Build workbook node."""
        if isinstance(node.value, Bag):
            for child in node.value:
                self._build_node(child, wb)

    def _build_sheet(self, node: BagNode, wb: Any) -> None:
        """Build sheet node."""
        name = node.attr.get("name", "Sheet1")
        ws = wb.create_sheet(title=str(name))

        freeze_panes = node.attr.get("freeze_panes")
        if freeze_panes:
            ws.freeze_panes = str(freeze_panes)

        autofilter = node.attr.get("autofilter")
        if autofilter:
            ws.auto_filter.ref = str(autofilter)

        if isinstance(node.value, Bag):
            merge_nodes: list[BagNode] = []
            chart_nodes: list[BagNode] = []

            row_idx = 1
            for child_node in node.value:
                if child_node.tag == "row":
                    self._build_row(child_node, ws, row_idx)
                    row_idx += 1
                elif child_node.tag == "merge":
                    merge_nodes.append(child_node)
                elif child_node.tag == "chart":
                    chart_nodes.append(child_node)

            for merge_node in merge_nodes:
                self._build_merge(merge_node, ws)

            for chart_node in chart_nodes:
                self._build_chart(chart_node, ws)

    def _build_row(self, node: BagNode, ws: Any, row_idx: int) -> None:
        """Build row node."""
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
                    self._build_cell(cell_node, ws, row_idx, col_idx)
                    col_idx += 1

    def _build_cell(self, node: BagNode, ws: Any, row_idx: int, col_idx: int) -> None:
        """Build cell node."""
        content = node.attr.get("content", "")
        formula = node.attr.get("formula")
        width = node.attr.get("width")

        cell = ws.cell(row=row_idx, column=col_idx)
        self._live_map[id(node)] = cell

        if formula:
            cell.value = formula
        else:
            cell.value = content

        if width:
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        self._apply_font(node, cell)

        bg_color = node.attr.get("bg_color")
        if bg_color:
            cell.fill = PatternFill(
                start_color=str(bg_color), end_color=str(bg_color), fill_type="solid"
            )

        self._apply_alignment(node, cell)
        self._apply_border(node, cell)

        number_format = node.attr.get("number_format")
        if number_format:
            cell.number_format = str(number_format)

    def _build_merge(self, node: BagNode, ws: Any) -> None:
        """Build merge node."""
        cell_range = node.attr.get("range", "")
        if cell_range:
            ws.merge_cells(str(cell_range))

    def _build_chart(self, node: BagNode, ws: Any) -> None:
        """Build chart node."""
        chart_type = node.attr.get("type", "bar")
        title = node.attr.get("title")
        data_range = node.attr.get("data_range", "")
        categories_range = node.attr.get("categories_range")
        position = node.attr.get("position", "E1")
        width = node.attr.get("width", 15)
        height = node.attr.get("height", 10)

        if not data_range:
            return

        if chart_type == "line":
            chart = LineChart()
        elif chart_type == "pie":
            chart = PieChart()
        else:
            chart = BarChart()

        if title:
            chart.title = str(title)

        chart.width = width
        chart.height = height

        data_ref = self._parse_range_reference(ws, str(data_range))
        if data_ref:
            chart.add_data(data_ref, titles_from_data=True)

        if categories_range:
            cat_ref = self._parse_range_reference(ws, str(categories_range))
            if cat_ref:
                chart.set_categories(cat_ref)

        ws.add_chart(chart, str(position))

    def _parse_range_reference(self, ws: Any, range_str: str) -> Any:
        """Parse a range string into a Reference object."""
        if ":" not in range_str:
            return None

        start, end = range_str.split(":")

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

        min_col = column_index_from_string(start_col)
        max_col = column_index_from_string(end_col)
        min_row = int(start_row)
        max_row = int(end_row)

        return Reference(ws, min_col=min_col, min_row=min_row, max_col=max_col, max_row=max_row)

    # -------------------------------------------------------------------------
    # Formatting helpers
    # -------------------------------------------------------------------------

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
                color=str(font_color) if font_color else None,
            )

    def _apply_alignment(self, node: BagNode, cell: Any) -> None:
        """Apply alignment to cell."""
        align = node.attr.get("align")
        valign = node.attr.get("valign")
        wrap_text = node.attr.get("wrap_text", False)

        if align or valign or wrap_text:
            cell.alignment = Alignment(
                horizontal=str(align) if align else None,
                vertical=str(valign) if valign else None,
                wrap_text=wrap_text,
            )

    def _apply_border(self, node: BagNode, cell: Any) -> None:
        """Apply border to cell."""
        border_style = node.attr.get("border")
        border_color = node.attr.get("border_color", "000000")

        if border_style:
            side = Side(style=str(border_style), color=str(border_color))
            cell.border = Border(left=side, right=side, top=side, bottom=side)


# Set compiler_class after ExcelCompiler is defined
from genro_office.builders.excel_builder import ExcelBuilder  # noqa: E402

ExcelBuilder.compiler_class = ExcelCompiler
