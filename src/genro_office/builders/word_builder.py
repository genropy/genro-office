# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""WordBuilder - Builder per documenti Word (.docx)."""

from __future__ import annotations

from io import BytesIO
from typing import TYPE_CHECKING, Any

from genro_bag import Bag
from genro_bag.builder import BagBuilderBase, element

if TYPE_CHECKING:
    from collections.abc import Iterator

    from genro_bag.bagnode import BagNode

try:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Inches, Pt, RGBColor

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


class WordBuilder(BagBuilderBase):
    """Builder per documenti Word (.docx) usando python-docx."""

    def __init__(self, bag: Bag) -> None:
        if not DOCX_AVAILABLE:
            msg = "python-docx required: pip install python-docx"
            raise ImportError(msg)
        super().__init__(bag)

    # -------------------------------------------------------------------------
    # Element definitions
    # -------------------------------------------------------------------------

    @element(sub_tags="heading,paragraph,table,image,pagebreak,itemlist,header,footer,run")
    def document(
        self,
        title: str = "",
        # Page setup
        orientation: str | None = None,
        margin_top: float | None = None,
        margin_bottom: float | None = None,
        margin_left: float | None = None,
        margin_right: float | None = None,
    ) -> None:
        """Root document element.

        Args:
            title: Document title (added as heading level 0).
            orientation: Page orientation ("portrait" or "landscape").
            margin_top: Top margin in cm.
            margin_bottom: Bottom margin in cm.
            margin_left: Left margin in cm.
            margin_right: Right margin in cm.
        """
        ...

    @element(sub_tags="run", parent_tags="document")
    def heading(
        self,
        content: str = "",
        level: int = 1,
        # Formatting
        bold: bool | None = None,
        italic: bool | None = None,
        color: str | None = None,
    ) -> None:
        """Heading element (H1-H9).

        Args:
            content: Heading text.
            level: Heading level (1-9).
            bold: Override bold formatting.
            italic: Override italic formatting.
            color: Text color as hex (e.g., "FF0000").
        """
        ...

    @element(sub_tags="run", parent_tags="document,cell,header,footer,listitem")
    def paragraph(
        self,
        content: str = "",
        style: str | None = None,
        # Formatting
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: int | None = None,
        font_name: str | None = None,
        color: str | None = None,
        # Alignment
        align: str | None = None,
        # Spacing
        space_before: float | None = None,
        space_after: float | None = None,
        line_spacing: float | None = None,
    ) -> None:
        """Paragraph element with formatting support.

        Args:
            content: Paragraph text.
            style: Word style name.
            bold: Bold text.
            italic: Italic text.
            underline: Underlined text.
            font_size: Font size in points.
            font_name: Font name (e.g., "Arial", "Times New Roman").
            color: Text color as hex (e.g., "FF0000" for red).
            align: Text alignment ("left", "center", "right", "justify").
            space_before: Space before paragraph in points.
            space_after: Space after paragraph in points.
            line_spacing: Line spacing multiplier (e.g., 1.5 for 1.5 lines).
        """
        ...

    @element(sub_tags="", parent_tags="paragraph,heading,cell")
    def run(
        self,
        content: str = "",
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        strike: bool = False,
        font_size: int | None = None,
        font_name: str | None = None,
        color: str | None = None,
        highlight: str | None = None,
    ) -> None:
        """Inline text run with formatting.

        Args:
            content: Text content.
            bold: Bold text.
            italic: Italic text.
            underline: Underlined text.
            strike: Strikethrough text.
            font_size: Font size in points.
            font_name: Font name.
            color: Text color as hex.
            highlight: Highlight color ("yellow", "green", "cyan", "magenta", etc.).
        """
        ...

    @element(sub_tags="item", parent_tags="document")
    def itemlist(self, type: str = "bullet") -> None:
        """List element (bulleted or numbered).

        Args:
            type: List type ("bullet" or "number").
        """
        ...

    @element(sub_tags="paragraph,run", parent_tags="itemlist")
    def item(self, content: str = "") -> None:
        """List item element.

        Args:
            content: Item text.
        """
        ...

    @element(sub_tags="row", parent_tags="document")
    def table(
        self,
        style: str | None = None,
        align: str | None = None,
        autofit: bool = True,
    ) -> None:
        """Table element.

        Args:
            style: Table style name (e.g., "Table Grid", "Light Shading").
            align: Table alignment ("left", "center", "right").
            autofit: Auto-fit table to content.
        """
        ...

    @element(sub_tags="cell", parent_tags="table")
    def row(self, height: float | None = None) -> None:
        """Table row element.

        Args:
            height: Row height in cm.
        """
        ...

    @element(sub_tags="paragraph,run", parent_tags="row")
    def cell(
        self,
        content: str = "",
        width: float | None = None,
        bold: bool = False,
        bg_color: str | None = None,
        align: str | None = None,
        valign: str | None = None,
    ) -> None:
        """Table cell element.

        Args:
            content: Cell text.
            width: Column width in cm.
            bold: Bold text.
            bg_color: Background color as hex.
            align: Horizontal alignment ("left", "center", "right").
            valign: Vertical alignment ("top", "center", "bottom").
        """
        ...

    @element(sub_tags="", parent_tags="document")
    def image(
        self,
        path: str = "",
        width: float | None = None,
        height: float | None = None,
        align: str | None = None,
    ) -> None:
        """Image element.

        Args:
            path: Image file path.
            width: Image width in inches.
            height: Image height in inches.
            align: Image alignment ("left", "center", "right").
        """
        ...

    @element(sub_tags="", parent_tags="document")
    def pagebreak(self) -> None:
        """Page break element."""
        ...

    @element(sub_tags="paragraph,run", parent_tags="document")
    def header(self) -> None:
        """Document header element."""
        ...

    @element(sub_tags="paragraph,run", parent_tags="document")
    def footer(self) -> None:
        """Document footer element."""
        ...

    # -------------------------------------------------------------------------
    # Compile: Bag → bytes (docx)
    # -------------------------------------------------------------------------

    def compile(self, bag: Bag) -> bytes:
        """Compile Bag to Word document bytes."""
        doc = Document()

        for node in bag:
            self._compile_node(node, doc)

        buffer = BytesIO()
        doc.save(buffer)
        return buffer.getvalue()

    def _compile_node(self, node: BagNode, doc: Any) -> None:
        """Compile a single node."""
        tag = node.tag or ""

        compile_method = getattr(self, f"_compile_{tag}", None)
        if compile_method:
            compile_method(node, doc)

    def _compile_document(self, node: BagNode, doc: Any) -> None:
        """Compile document node."""
        # Page setup
        section = doc.sections[0]

        orientation = node.attr.get("orientation")
        if orientation == "landscape":
            section.orientation = WD_ORIENT.LANDSCAPE
            # Swap width and height for landscape
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

        # Title
        title = node.attr.get("title", "")
        if title:
            doc.add_heading(title, level=0)

        # Children
        if isinstance(node.value, Bag):
            for child in node.value:
                self._compile_node(child, doc)

    def _compile_heading(self, node: BagNode, doc: Any) -> None:
        """Compile heading node."""
        content = node.attr.get("content", "")
        level = node.attr.get("level", 1)

        heading = doc.add_heading(content, level=level)

        # Apply formatting to the heading run
        if heading.runs:
            run = heading.runs[0]
            self._apply_run_formatting(node, run)

        # Process child runs
        if isinstance(node.value, Bag):
            for child in node.value:
                if child.tag == "run":
                    self._compile_run_to_paragraph(child, heading)

    def _compile_paragraph(self, node: BagNode, doc: Any) -> None:
        """Compile paragraph node."""
        content = node.attr.get("content", "")
        style = node.attr.get("style")
        para = doc.add_paragraph(style=style) if style else doc.add_paragraph()

        # Add content with formatting
        if content:
            run = para.add_run(content)
            self._apply_run_formatting(node, run)

        # Alignment
        align = node.attr.get("align")
        if align:
            para.alignment = self._get_alignment(align)

        # Spacing
        space_before = node.attr.get("space_before")
        if space_before is not None:
            para.paragraph_format.space_before = Pt(space_before)

        space_after = node.attr.get("space_after")
        if space_after is not None:
            para.paragraph_format.space_after = Pt(space_after)

        line_spacing = node.attr.get("line_spacing")
        if line_spacing is not None:
            para.paragraph_format.line_spacing = line_spacing

        # Process child runs
        if isinstance(node.value, Bag):
            for child in node.value:
                if child.tag == "run":
                    self._compile_run_to_paragraph(child, para)

    def _compile_run_to_paragraph(self, node: BagNode, para: Any) -> None:
        """Compile run node and add to paragraph."""
        content = node.attr.get("content", "")
        run = para.add_run(content)
        self._apply_run_formatting(node, run)

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
            run.font.name = font_name

        color = node.attr.get("color")
        if color:
            run.font.color.rgb = RGBColor.from_string(color)

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
            if highlight.lower() in highlight_map:
                run.font.highlight_color = highlight_map[highlight.lower()]

    def _get_alignment(self, align: str) -> Any:
        """Convert alignment string to WD_ALIGN_PARAGRAPH."""
        align_map = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return align_map.get(align.lower(), WD_ALIGN_PARAGRAPH.LEFT)

    def _compile_itemlist(self, node: BagNode, doc: Any) -> None:
        """Compile itemlist node."""
        list_type = node.attr.get("type", "bullet")

        if isinstance(node.value, Bag):
            for item_node in node.value:
                if item_node.tag == "item":
                    self._compile_item(item_node, doc, list_type)

    def _compile_item(self, node: BagNode, doc: Any, list_type: str) -> None:
        """Compile item node."""
        content = node.attr.get("content", "")
        style = "List Number" if list_type == "number" else "List Bullet"
        para = doc.add_paragraph(content, style=style)

        # Process child runs/paragraphs
        if isinstance(node.value, Bag):
            for child in node.value:
                if child.tag == "run":
                    self._compile_run_to_paragraph(child, para)

    def _iter_table_rows(
        self, node: BagNode
    ) -> Iterator[tuple[BagNode, list[BagNode]]]:
        """Yield (row_node, cells) tuples from table node."""
        if not isinstance(node.value, Bag):
            return
        for row_node in node.value:
            if row_node.tag != "row":
                continue
            cells: list[BagNode] = []
            if isinstance(row_node.value, Bag):
                cells = [c for c in row_node.value if c.tag == "cell"]
            yield row_node, cells

    def _compile_table(self, node: BagNode, doc: Any) -> None:
        """Compile table node."""
        rows_data = list(self._iter_table_rows(node))
        if not rows_data:
            return

        num_cols = max(len(cells) for _, cells in rows_data)
        table = doc.add_table(rows=len(rows_data), cols=num_cols)

        # Table style
        style = node.attr.get("style")
        if style:
            table.style = style

        # Table alignment
        self._apply_table_alignment(node, table)

        # Populate cells
        for i, (row_node, cells) in enumerate(rows_data):
            self._compile_table_row(table, i, row_node, cells)

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

    def _compile_table_row(
        self, table: Any, row_idx: int, row_node: BagNode, cells: list[BagNode]
    ) -> None:
        """Compile a single table row."""
        table_row = table.rows[row_idx]

        row_height = row_node.attr.get("height")
        if row_height:
            table_row.height = Cm(row_height)

        for col_idx, cell_node in enumerate(cells):
            cell = table.cell(row_idx, col_idx)
            self._compile_table_cell(cell_node, cell)

    def _compile_table_cell(self, node: BagNode, cell: Any) -> None:
        """Compile table cell content."""
        content = node.attr.get("content", "")
        width = node.attr.get("width")
        bold = node.attr.get("bold", False)
        bg_color = node.attr.get("bg_color")
        align = node.attr.get("align")
        valign = node.attr.get("valign")

        # Set cell width
        if width:
            cell.width = Cm(width)

        # Background color
        if bg_color:
            self._set_cell_shading(cell, bg_color)

        # Vertical alignment
        if valign:
            valign_map = {
                "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
                "center": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
                "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
            }
            if valign.lower() in valign_map:
                cell.vertical_alignment = valign_map[valign.lower()]

        # Add content
        if content:
            para = cell.paragraphs[0]
            para.clear()
            run = para.add_run(str(content))

            if bold:
                run.bold = True

            # Apply other formatting from node
            self._apply_run_formatting(node, run)

            # Horizontal alignment
            if align:
                para.alignment = self._get_alignment(align)

        # Process child elements
        if isinstance(node.value, Bag):
            for child in node.value:
                if child.tag == "paragraph":
                    para = cell.add_paragraph()
                    child_content = child.attr.get("content", "")
                    if child_content:
                        run = para.add_run(child_content)
                        self._apply_run_formatting(child, run)
                elif child.tag == "run":
                    para = cell.paragraphs[-1] if cell.paragraphs else cell.add_paragraph()
                    self._compile_run_to_paragraph(child, para)

    def _set_cell_shading(self, cell: Any, color: str) -> None:
        """Set cell background color."""
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), color)
        cell._tc.get_or_add_tcPr().append(shading)

    def _compile_image(self, node: BagNode, doc: Any) -> None:
        """Compile image node."""
        path = node.attr.get("path", "")
        width = node.attr.get("width")
        height = node.attr.get("height")
        align = node.attr.get("align")

        if not path:
            return

        # Create paragraph for alignment
        para = doc.add_paragraph()

        if align:
            para.alignment = self._get_alignment(align)

        run = para.add_run()

        # Add picture with dimensions
        if width and height:
            run.add_picture(path, width=Inches(width), height=Inches(height))
        elif width:
            run.add_picture(path, width=Inches(width))
        elif height:
            run.add_picture(path, height=Inches(height))
        else:
            run.add_picture(path)

    def _compile_pagebreak(self, _node: BagNode, doc: Any) -> None:
        """Compile pagebreak node."""
        doc.add_page_break()

    def _compile_header(self, node: BagNode, doc: Any) -> None:
        """Compile header node."""
        section = doc.sections[0]
        header = section.header
        header.is_linked_to_previous = False

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.tag == "paragraph":
                    content = child.attr.get("content", "")
                    para = header.add_paragraph(content)
                    self._apply_paragraph_formatting(child, para)
                elif child.tag == "run":
                    para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                    self._compile_run_to_paragraph(child, para)

    def _compile_footer(self, node: BagNode, doc: Any) -> None:
        """Compile footer node."""
        section = doc.sections[0]
        footer = section.footer
        footer.is_linked_to_previous = False

        if isinstance(node.value, Bag):
            for child in node.value:
                if child.tag == "paragraph":
                    content = child.attr.get("content", "")
                    para = footer.add_paragraph(content)
                    self._apply_paragraph_formatting(child, para)
                elif child.tag == "run":
                    para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                    self._compile_run_to_paragraph(child, para)

    def _apply_paragraph_formatting(self, node: BagNode, para: Any) -> None:
        """Apply paragraph-level formatting."""
        align = node.attr.get("align")
        if align:
            para.alignment = self._get_alignment(align)

        # Apply formatting to runs
        if para.runs:
            for run in para.runs:
                self._apply_run_formatting(node, run)
