# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""WordBuilder - Builder for Word documents (.docx).

Defines the schema elements for Word document generation.
Compilation is handled by WordCompiler.
"""

from __future__ import annotations

from typing import TYPE_CHECKING

from genro_builders import BagBuilderBase
from genro_builders.builder import element

if TYPE_CHECKING:
    from genro_bag import Bag

try:
    from docx import Document as _Document  # noqa: F401

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False


class WordBuilder(BagBuilderBase):
    """Builder for Word documents (.docx) using python-docx."""

    _compiler_class: type | None = None

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
        bold: bool = False,
        italic: bool = False,
        underline: bool = False,
        font_size: int | None = None,
        font_name: str | None = None,
        color: str | None = None,
        align: str | None = None,
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
