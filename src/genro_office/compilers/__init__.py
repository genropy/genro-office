# Copyright 2025 Softwell S.r.l. - Licensed under Apache License 2.0
# See LICENSE file for details

"""Office document compilers for genro-office."""

__all__: list[str] = []

# WordCompiler is optional - only available if python-docx is installed
try:
    from genro_office.compilers.word_compiler import WordCompiler as WordCompiler

    __all__.append("WordCompiler")
except ImportError:
    pass

# ExcelCompiler is optional - only available if openpyxl is installed
try:
    from genro_office.compilers.excel_compiler import ExcelCompiler as ExcelCompiler

    __all__.append("ExcelCompiler")
except ImportError:
    pass
