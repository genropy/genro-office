# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Tests for WordApp."""

from __future__ import annotations

import tempfile
from pathlib import Path

from genro_office import WordApp


class SimpleWordDoc(WordApp):
    """Simple Word document for testing."""

    def recipe(self, root):
        doc = root.document(title="Test Document")
        doc.heading(content="Chapter 1", level=1)
        doc.paragraph(content="This is a test paragraph.")


class WordDocWithTable(WordApp):
    """Word document with table for testing."""

    def recipe(self, root):
        doc = root.document(title="Table Test")
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
        assert app.page is not None
        assert app.data is not None

    def test_render_simple_document(self):
        """Test rendering a simple document."""
        app = SimpleWordDoc()
        result = app.render()

        assert isinstance(result, bytes)
        assert len(result) > 0
        # DOCX files start with PK (ZIP signature)
        assert result[:2] == b"PK"

    def test_render_document_with_table(self):
        """Test rendering a document with a table."""
        app = WordDocWithTable()
        result = app.render()

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_save_document(self):
        """Test saving a document to file."""
        app = SimpleWordDoc()

        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "test.docx"
            app.save(str(filepath))

            assert filepath.exists()
            assert filepath.stat().st_size > 0

            # Verify it's a valid DOCX (ZIP file)
            with open(filepath, "rb") as f:
                assert f.read(2) == b"PK"

    def test_page_property(self):
        """Test page property returns the Bag."""
        app = SimpleWordDoc()
        assert app.page is app._page

    def test_data_property(self):
        """Test data property returns the data Bag."""
        app = SimpleWordDoc()
        assert app.data is app._data
