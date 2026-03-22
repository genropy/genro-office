# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Tests for WordApp."""

from __future__ import annotations

import tempfile
from pathlib import Path

from genro_office import WordApp


class SimpleWordDoc(WordApp):
    """Simple Word document for testing."""

    def recipe(self, store):
        doc = store.document(title="Test Document")
        doc.heading(content="Chapter 1", level=1)
        doc.paragraph(content="This is a test paragraph.")


class WordDocWithTable(WordApp):
    """Word document with table for testing."""

    def recipe(self, store):
        doc = store.document(title="Table Test")
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
        assert app.store is not None
        assert app.data is not None

    def test_render_simple_document(self):
        """Test rendering a simple document."""
        app = SimpleWordDoc()
        app.setup()
        result = app.render()

        assert isinstance(result, bytes)
        assert len(result) > 0
        # DOCX files start with PK (ZIP signature)
        assert result[:2] == b"PK"

    def test_render_document_with_table(self):
        """Test rendering a document with a table."""
        app = WordDocWithTable()
        app.setup()
        result = app.render()

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_save_document(self):
        """Test saving a document to file."""
        app = SimpleWordDoc()
        app.setup()

        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "test.docx"
            app.save(str(filepath))

            assert filepath.exists()
            assert filepath.stat().st_size > 0

            # Verify it's a valid DOCX (ZIP file)
            with open(filepath, "rb") as f:
                assert f.read(2) == b"PK"

    def test_store_property(self):
        """Test store property returns the BuilderBag."""
        app = SimpleWordDoc()
        assert app.store is app._store

    def test_data_property(self):
        """Test data property returns the data Bag."""
        app = SimpleWordDoc()
        assert app.data is app._data

    def test_data_binding(self):
        """Test ^pointer data binding in recipe."""

        class BoundDoc(WordApp):
            def recipe(self, store):
                doc = store.document()
                doc.heading(content="^doc.title", level=1)
                doc.paragraph(content="^doc.body")

        app = BoundDoc()
        app.data["doc.title"] = "Bound Title"
        app.data["doc.body"] = "Bound content paragraph."
        app.setup()

        result = app.render()
        assert isinstance(result, bytes)
        assert result[:2] == b"PK"
