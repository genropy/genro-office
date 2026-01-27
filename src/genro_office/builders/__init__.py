# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Office document builders for genro-office."""

__all__: list[str] = []

# WordBuilder is optional - only available if python-docx is installed
try:
    from genro_office.builders.word_builder import WordBuilder as WordBuilder

    __all__.append("WordBuilder")
except ImportError:
    pass

# ExcelBuilder is optional - only available if openpyxl is installed
try:
    from genro_office.builders.excel_builder import ExcelBuilder as ExcelBuilder

    __all__.append("ExcelBuilder")
except ImportError:
    pass
