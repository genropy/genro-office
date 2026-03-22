# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Tests for ExcelApp."""

from __future__ import annotations

import tempfile
from pathlib import Path

from genro_office import ExcelApp


class SimpleExcelDoc(ExcelApp):
    """Simple Excel spreadsheet for testing."""

    def recipe(self, store):
        wb = store.workbook()
        sheet = wb.sheet(name="Data")

        row1 = sheet.row()
        row1.cell(content="Name")
        row1.cell(content="Value")

        row2 = sheet.row()
        row2.cell(content="Alpha")
        row2.cell(content=100)


class ExcelDocWithFormula(ExcelApp):
    """Excel spreadsheet with formula for testing."""

    def recipe(self, store):
        wb = store.workbook()
        sheet = wb.sheet(name="Calculations")

        row1 = sheet.row()
        row1.cell(content="A")
        row1.cell(content=10)

        row2 = sheet.row()
        row2.cell(content="B")
        row2.cell(content=20)

        row3 = sheet.row()
        row3.cell(content="Sum")
        row3.cell(formula="=B1+B2")


class ExcelDocWithStyling(ExcelApp):
    """Excel spreadsheet with styling for testing."""

    def recipe(self, store):
        wb = store.workbook()
        sheet = wb.sheet(name="Styled")

        row = sheet.row(height=30.0)
        row.cell(content="Bold", bold=True, width=20.0)
        row.cell(content="Italic", italic=True)
        row.cell(content="Large", font_size=16)


class TestExcelApp:
    """Tests for ExcelApp."""

    def test_create_empty_app(self):
        """Test creating an empty ExcelApp."""
        app = ExcelApp()
        assert app.store is not None
        assert app.data is not None

    def test_render_simple_spreadsheet(self):
        """Test rendering a simple spreadsheet."""
        app = SimpleExcelDoc()
        app.setup()
        result = app.output

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_render_spreadsheet_with_formula(self):
        """Test rendering a spreadsheet with formula."""
        app = ExcelDocWithFormula()
        app.setup()
        result = app.output

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_render_spreadsheet_with_styling(self):
        """Test rendering a spreadsheet with styling."""
        app = ExcelDocWithStyling()
        app.setup()
        result = app.output

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_save_spreadsheet(self):
        """Test saving a spreadsheet to file."""
        app = SimpleExcelDoc()
        app.setup()

        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "test.xlsx"
            app.save(str(filepath))

            assert filepath.exists()
            assert filepath.stat().st_size > 0

            with open(filepath, "rb") as f:
                assert f.read(2) == b"PK"

    def test_store_property(self):
        """Test store property returns the BuilderBag."""
        app = SimpleExcelDoc()
        assert app.store is app._store

    def test_data_property(self):
        """Test data property returns the data Bag."""
        app = SimpleExcelDoc()
        assert app.data is app._data

    def test_data_binding(self):
        """Test ^pointer data binding in recipe."""

        class BoundSheet(ExcelApp):
            def recipe(self, store):
                wb = store.workbook()
                sheet = wb.sheet(name="Bound")
                row = sheet.row()
                row.cell(content="^headers.col1")
                row.cell(content="^headers.col2")

        app = BoundSheet()
        app.data["headers.col1"] = "Name"
        app.data["headers.col2"] = "Value"
        app.setup()

        result = app.output
        assert isinstance(result, bytes)
        assert result[:2] == b"PK"
