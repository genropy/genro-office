# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Tests for ExcelApp."""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest

from genro_office import ExcelApp


class SimpleExcelDoc(ExcelApp):
    """Simple Excel spreadsheet for testing."""

    def recipe(self, root):
        wb = root.workbook()
        sheet = wb.sheet(name="Data")

        row1 = sheet.row()
        row1.cell(content="Name")
        row1.cell(content="Value")

        row2 = sheet.row()
        row2.cell(content="Alpha")
        row2.cell(content=100)


class ExcelDocWithFormula(ExcelApp):
    """Excel spreadsheet with formula for testing."""

    def recipe(self, root):
        wb = root.workbook()
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

    def recipe(self, root):
        wb = root.workbook()
        sheet = wb.sheet(name="Styled")

        row = sheet.row(height=30)
        row.cell(content="Bold", bold=True, width=20)
        row.cell(content="Italic", italic=True)
        row.cell(content="Large", font_size=16)


class TestExcelApp:
    """Tests for ExcelApp."""

    def test_create_empty_app(self):
        """Test creating an empty ExcelApp."""
        app = ExcelApp()
        assert app.page is not None
        assert app.data is not None

    def test_render_simple_spreadsheet(self):
        """Test rendering a simple spreadsheet."""
        app = SimpleExcelDoc()
        result = app.render()

        assert isinstance(result, bytes)
        assert len(result) > 0
        # XLSX files start with PK (ZIP signature)
        assert result[:2] == b"PK"

    def test_render_spreadsheet_with_formula(self):
        """Test rendering a spreadsheet with formula."""
        app = ExcelDocWithFormula()
        result = app.render()

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_render_spreadsheet_with_styling(self):
        """Test rendering a spreadsheet with styling."""
        app = ExcelDocWithStyling()
        result = app.render()

        assert isinstance(result, bytes)
        assert len(result) > 0
        assert result[:2] == b"PK"

    def test_save_spreadsheet(self):
        """Test saving a spreadsheet to file."""
        app = SimpleExcelDoc()

        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "test.xlsx"
            app.save(str(filepath))

            assert filepath.exists()
            assert filepath.stat().st_size > 0

            # Verify it's a valid XLSX (ZIP file)
            with open(filepath, "rb") as f:
                assert f.read(2) == b"PK"

    def test_page_property(self):
        """Test page property returns the Bag."""
        app = SimpleExcelDoc()
        assert app.page is app._page

    def test_data_property(self):
        """Test data property returns the data Bag."""
        app = SimpleExcelDoc()
        assert app.data is app._data
