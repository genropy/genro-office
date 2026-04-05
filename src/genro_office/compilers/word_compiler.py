# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""WordCompiler - Compiler for Word documents (.docx).

Transforms a Bag (built with WordBuilder) into a Word document.
Maintains a live Document object for incremental updates.
"""

from __future__ import annotations

from io import BytesIO
from typing import TYPE_CHECKING, Any

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from genro_bag import Bag
from genro_builders import BagCompilerBase

if TYPE_CHECKING:
    from collections.abc import Iterator

    from genro_bag import BagNode


class WordCompiler(BagCompilerBase):
    """Compiler Word: Bag -> bytes (.docx).

    Maintains a live Document and a _live_map {node_id: live_object}
    for incremental updates when data changes.
    """

    def __init__(self, builder: Any) -> None:
        super().__init__(builder)
        self._doc: Any | None = None
        self._live_map: dict[int, Any] = {}
        self._custom_handlers: dict[str, Any] = {}

    def register_handler(self, tag: str, handler: Any) -> None:
        """Register a custom build handler for a tag.

        The handler receives (node, doc) and is called during document
        building. Custom handlers take precedence over built-in ones.

        Args:
            tag: The element tag name (e.g., "qrcode").
            handler: Callable(node, doc) that builds the element.
        """
        self._custom_handlers[tag] = handler

    # -------------------------------------------------------------------------
    # Rendering (CompiledBag → bytes)
    # -------------------------------------------------------------------------

    def render(self, compiled_bag: Bag) -> bytes:
        """Render a CompiledBag to Word document bytes.

        Args:
            compiled_bag: The compiled Bag (components expanded, pointers resolved).

        Returns:
            Word document as bytes (.docx format).
        """
        return self._render_to_bytes(compiled_bag)

    def serialize(self) -> bytes:
        """Serialize the live Document to bytes without rebuilding."""
        if self._doc is None:
            return b""
        buffer = BytesIO()
        self._doc.save(buffer)
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

    def _apply_live_update(self, node: BagNode, live_obj: Any) -> bool:
        """Apply attribute changes to a live docx object.

        Handles runs, paragraphs, headings, table cells, and list items.
        """
        tag = node.node_tag or ""

        if tag in ("run", "cell"):
            self._apply_run_formatting(node, live_obj)
            content = node.attr.get("content", "")
            if content:
                live_obj.text = str(content)
            return True

        if tag in ("paragraph", "heading", "item"):
            if hasattr(live_obj, "runs") and live_obj.runs:
                content = node.attr.get("content", "")
                if content and live_obj.runs:
                    live_obj.runs[0].text = str(content)
                self._apply_run_formatting(node, live_obj.runs[0])
            align = node.attr.get("align")
            if align:
                live_obj.alignment = self._get_alignment(align)
            return True

        return False

    # -------------------------------------------------------------------------
    # Document building (imperative walk)
    # -------------------------------------------------------------------------

    def _render_to_bytes(self, bag: Bag) -> bytes:
        """Build a Word Document from a Bag and return bytes."""
        self._doc = Document()
        self._live_map.clear()

        for node in bag:
            self._build_node(node, self._doc)

        buffer = BytesIO()
        self._doc.save(buffer)
        return buffer.getvalue()

    def _build_node(self, node: BagNode, doc: Any) -> None:
        """Build a single node into the Document."""
        tag = node.node_tag or ""

        build_method = self._custom_handlers.get(tag)
        if build_method is None:
            build_method = getattr(self, f"_build_{tag}", None)
        if build_method:
            build_method(node, doc)

    def _build_document(self, node: BagNode, doc: Any) -> None:
        """Build document node."""
        section = doc.sections[0]

        orientation = node.attr.get("orientation")
        if orientation == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            new_width = section.page_height
            new_height = section.page_width
            section.page_width = new_width
            section.page_height = new_height

        margin_top = node.attr.get("margin_top")
        if margin_top is not None:
            section.top_margin = Cm(margin_top)

        margin_bottom = node.attr.get("margin_bottom")
        if margin_bottom is not None:
            section.bottom_margin = Cm(margin_bottom)

        margin_left = node.attr.get("margin_left")
        if margin_left is not None:
            section.left_margin = Cm(margin_left)

        margin_right = node.attr.get("margin_right")
        if margin_right is not None:
            section.right_margin = Cm(margin_right)

        title = node.attr.get("title", "")
        if title:
            doc.add_heading(title, level=0)

        if isinstance(node.value, Bag):
            for child in node.value:
                self._build_node(child, doc)

    def _build_heading(self, node: BagNode, doc: Any) -> None:
        """Build heading node."""
        content = node.attr.get("content", "")
        level = node.attr.get("level", 1)

        heading = doc.add_heading(str(content), level=level)
        self._live_map[id(node)] = heading

        if heading.runs:
            run = heading.runs[0]
            self._apply_run_formatting(node, run)

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.node_tag == "run":
                    self._build_run_to_paragraph(child, heading)

    def _build_paragraph(self, node: BagNode, doc: Any) -> None:
        """Build paragraph node."""
        content = node.attr.get("content", "")
        style = node.attr.get("style")
        para = doc.add_paragraph(style=style) if style else doc.add_paragraph()
        self._live_map[id(node)] = para

        if content:
            run = para.add_run(str(content))
            self._apply_run_formatting(node, run)

        align = node.attr.get("align")
        if align:
            para.alignment = self._get_alignment(align)

        space_before = node.attr.get("space_before")
        if space_before is not None:
            para.paragraph_format.space_before = Pt(space_before)

        space_after = node.attr.get("space_after")
        if space_after is not None:
            para.paragraph_format.space_after = Pt(space_after)

        line_spacing = node.attr.get("line_spacing")
        if line_spacing is not None:
            para.paragraph_format.line_spacing = line_spacing

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.node_tag == "run":
                    self._build_run_to_paragraph(child, para)

    def _build_run_to_paragraph(self, node: BagNode, para: Any) -> None:
        """Build run node and add to paragraph."""
        content = node.attr.get("content", "")
        run = para.add_run(str(content))
        self._apply_run_formatting(node, run)
        self._live_map[id(node)] = run

    def _build_itemlist(self, node: BagNode, doc: Any) -> None:
        """Build itemlist node."""
        list_type = node.attr.get("type", "bullet")

        if isinstance(node.value, Bag):
            for item_node in node.value:
                if item_node.node_tag == "item":
                    self._build_item(item_node, doc, list_type)

    def _build_item(self, node: BagNode, doc: Any, list_type: str) -> None:
        """Build item node."""
        content = node.attr.get("content", "")
        style = "List Number" if list_type == "number" else "List Bullet"
        para = doc.add_paragraph(str(content), style=style)
        self._live_map[id(node)] = para

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.node_tag == "run":
                    self._build_run_to_paragraph(child, para)

    def _iter_table_rows(
        self, node: BagNode
    ) -> Iterator[tuple[BagNode, list[BagNode]]]:
        """Yield (row_node, cells) tuples from table node."""
        if not isinstance(node.value, Bag):
            return
        for row_node in node.value:
            if row_node.node_tag != "row":
                continue
            cells: list[BagNode] = []
            if isinstance(row_node.value, Bag):
                cells = [c for c in row_node.value if c.node_tag == "cell"]
            yield row_node, cells

    def _build_table(self, node: BagNode, doc: Any) -> None:
        """Build table node."""
        rows_data = list(self._iter_table_rows(node))
        if not rows_data:
            return

        num_cols = max(len(cells) for _, cells in rows_data)
        table = doc.add_table(rows=len(rows_data), cols=num_cols)

        style = node.attr.get("style")
        if style:
            table.style = style

        self._apply_table_alignment(node, table)

        for i, (row_node, cells) in enumerate(rows_data):
            self._build_table_row(table, i, row_node, cells)

    def _apply_table_alignment(self, node: BagNode, table: Any) -> None:
        """Apply alignment to table."""
        align = node.attr.get("align")
        if not align:
            return
        align_map = {
            "left": WD_TABLE_ALIGNMENT.LEFT,
            "center": WD_TABLE_ALIGNMENT.CENTER,
            "right": WD_TABLE_ALIGNMENT.RIGHT,
        }
        if align.lower() in align_map:
            table.alignment = align_map[align.lower()]

    def _build_table_row(
        self, table: Any, row_idx: int, row_node: BagNode, cells: list[BagNode]
    ) -> None:
        """Build a single table row."""
        table_row = table.rows[row_idx]

        row_height = row_node.attr.get("height")
        if row_height:
            table_row.height = Cm(row_height)

        for col_idx, cell_node in enumerate(cells):
            cell = table.cell(row_idx, col_idx)
            self._build_table_cell(cell_node, cell)

    def _build_table_cell(self, node: BagNode, cell: Any) -> None:
        """Build table cell content."""
        content = node.attr.get("content", "")
        width = node.attr.get("width")
        bold = node.attr.get("bold", False)
        bg_color = node.attr.get("bg_color")
        align = node.attr.get("align")
        valign = node.attr.get("valign")

        if width:
            cell.width = Cm(width)

        if bg_color:
            self._set_cell_shading(cell, bg_color)

        if valign:
            valign_map = {
                "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
                "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
            }
            if valign.lower() in valign_map:
                cell.vertical_alignment = valign_map[valign.lower()]

        if content:
            para = cell.paragraphs[0]
            para.clear()
            run = para.add_run(str(content))
            self._live_map[id(node)] = run

            if bold:
                run.bold = True

            self._apply_run_formatting(node, run)

            if align:
                para.alignment = self._get_alignment(align)

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.node_tag == "paragraph":
                    para = cell.add_paragraph()
                    child_content = child.attr.get("content", "")
                    if child_content:
                        run = para.add_run(str(child_content))
                        self._apply_run_formatting(child, run)
                        self._live_map[id(child)] = run
                elif child.node_tag == "run":
                    para = cell.paragraphs[-1] if cell.paragraphs else cell.add_paragraph()
                    self._build_run_to_paragraph(child, para)

    def _set_cell_shading(self, cell: Any, color: str) -> None:
        """Set cell background color."""
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), color)
        cell._tc.get_or_add_tcPr().append(shading)

    def _build_image(self, node: BagNode, doc: Any) -> None:
        """Build image node."""
        path = node.attr.get("path", "")
        width = node.attr.get("width")
        height = node.attr.get("height")
        align = node.attr.get("align")

        if not path:
            return

        para = doc.add_paragraph()

        if align:
            para.alignment = self._get_alignment(align)

        run = para.add_run()

        if width and height:
            run.add_picture(str(path), width=Inches(width), height=Inches(height))
        elif width:
            run.add_picture(str(path), width=Inches(width))
        elif height:
            run.add_picture(str(path), height=Inches(height))
        else:
            run.add_picture(str(path))

    def _build_pagebreak(self, _node: BagNode, doc: Any) -> None:
        """Build pagebreak node."""
        doc.add_page_break()

    def _build_header(self, node: BagNode, doc: Any) -> None:
        """Build header node."""
        section = doc.sections[0]
        header = section.header
        header.is_linked_to_previous = False

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.node_tag == "paragraph":
                    content = child.attr.get("content", "")
                    para = header.add_paragraph(str(content))
                    self._apply_paragraph_formatting(child, para)
                elif child.node_tag == "run":
                    para = (
                        header.paragraphs[0]
                        if header.paragraphs
                        else header.add_paragraph()
                    )
                    self._build_run_to_paragraph(child, para)

    def _build_footer(self, node: BagNode, doc: Any) -> None:
        """Build footer node."""
        section = doc.sections[0]
        footer = section.footer
        footer.is_linked_to_previous = False

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.node_tag == "paragraph":
                    content = child.attr.get("content", "")
                    para = footer.add_paragraph(str(content))
                    self._apply_paragraph_formatting(child, para)
                elif child.node_tag == "run":
                    para = (
                        footer.paragraphs[0]
                        if footer.paragraphs
                        else footer.add_paragraph()
                    )
                    self._build_run_to_paragraph(child, para)

    # -------------------------------------------------------------------------
    # Formatting helpers
    # -------------------------------------------------------------------------

    def _apply_run_formatting(self, node: BagNode, run: Any) -> None:
        """Apply formatting attributes to a run."""
        bold = node.attr.get("bold", False)
        if bold:
            run.bold = True

        italic = node.attr.get("italic", False)
        if italic:
            run.italic = True

        underline = node.attr.get("underline", False)
        if underline:
            run.underline = True

        strike = node.attr.get("strike", False)
        if strike:
            run.font.strike = True

        font_size = node.attr.get("font_size")
        if font_size:
            run.font.size = Pt(font_size)

        font_name = node.attr.get("font_name")
        if font_name:
            run.font.name = str(font_name)

        color = node.attr.get("color")
        if color:
            run.font.color.rgb = RGBColor.from_string(str(color))

        highlight = node.attr.get("highlight")
        if highlight:
            highlight_map = {
                "yellow": WD_COLOR_INDEX.YELLOW,
                "green": WD_COLOR_INDEX.BRIGHT_GREEN,
                "cyan": WD_COLOR_INDEX.TURQUOISE,
                "magenta": WD_COLOR_INDEX.PINK,
                "blue": WD_COLOR_INDEX.BLUE,
                "red": WD_COLOR_INDEX.RED,
                "gray": WD_COLOR_INDEX.GRAY_25,
            }
            if str(highlight).lower() in highlight_map:
                run.font.highlight_color = highlight_map[str(highlight).lower()]

    def _get_alignment(self, align: str) -> Any:
        """Convert alignment string to WD_ALIGN_PARAGRAPH."""
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return align_map.get(str(align).lower(), WD_ALIGN_PARAGRAPH.LEFT)

    def _apply_paragraph_formatting(self, node: BagNode, para: Any) -> None:
        """Apply paragraph-level formatting."""
        align = node.attr.get("align")
        if align:
            para.alignment = self._get_alignment(align)

        if para.runs:
            for run in para.runs:
                self._apply_run_formatting(node, run)


# Set _compiler_class after WordCompiler is defined
from genro_office.builders.word_builder import WordBuilder  # noqa: E402

WordBuilder._compiler_class = WordCompiler
