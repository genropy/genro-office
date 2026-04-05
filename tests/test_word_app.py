# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Tests for WordApp."""

from __future__ import annotations

import tempfile
from pathlib import Path

from genro_office import WordApp


class SimpleWordDoc(WordApp):
    """Simple Word document for testing."""

    def main(self, source):
        doc = source.document(title="Test Document")
        doc.heading(content="Chapter 1", level=1)
        doc.paragraph(content="This is a test paragraph.")


class WordDocWithTable(WordApp):
    """Word document with table for testing."""

    def main(self, source):
        doc = source.document(title="Table Test")
        doc.paragraph(content="Here is a table:")

        table = doc.table()
        row1 = table.row()
        row1.cell(content="Name")
        row1.cell(content="Value")

        row2 = table.row()
        row2.cell(content="Alpha")
        row2.cell(content="100")


class TestWordApp:
    """Tests for WordApp."""

    def test_create_empty_app(self):
        """Test creating an empty WordApp."""
        app = WordApp()
        assert app.builder.source is not None
        assert app.data is not None

    def test_render_simple_document(self):
        """Test rendering a simple document."""
        app = SimpleWordDoc()
        app.build()
        result = app.output

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_render_document_with_table(self):
        """Test rendering a document with a table."""
        app = WordDocWithTable()
        app.build()
        result = app.output

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_save_document(self):
        """Test saving a document to file."""
        app = SimpleWordDoc()
        app.build()

        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "test.docx"
            app.save(str(filepath))

            assert filepath.exists()
            assert filepath.stat().st_size > 0

            with open(filepath, "rb") as f:
                assert f.read(2) == b"PK"

    def test_source_property(self):
        """Test source property returns a BuilderBag."""
        app = SimpleWordDoc()
        assert app.builder.source is not None

    def test_data_property(self):
        """Test data property returns the shared data Bag."""
        app = SimpleWordDoc()
        assert app.data is app.builder.data

    def test_live_update_cell(self):
        """Test live update handles table cell tag."""
        app = WordDocWithTable()
        app.build()
        compiler = app._word_compiler
        # _live_map is populated for cell runs during build
        assert len(compiler._live_map) > 0

    def test_register_handler(self):
        """Test custom handler registration and dispatch."""
        called_with = []

        def custom_handler(node, _doc):
            called_with.append(node.node_tag)

        app = WordApp()
        compiler = app._word_compiler
        compiler.register_handler("custom_tag", custom_handler)
        assert "custom_tag" in compiler._custom_handlers

    def test_data_binding(self):
        """Test ^pointer data binding in main."""

        class BoundDoc(WordApp):
            def main(self, source):
                doc = source.document()
                doc.heading(content="^doc.title", level=1)
                doc.paragraph(content="^doc.body")

        app = BoundDoc()
        app.data["doc.title"] = "Bound Title"
        app.data["doc.body"] = "Bound content paragraph."
        app.build()

        result = app.output
        assert isinstance(result, bytes)
        assert result[:2] == b"PK"
