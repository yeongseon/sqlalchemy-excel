"""sqlalchemy-excel: SQLAlchemy model-driven Excel toolkit.

Generate Excel templates, validate uploads, import to database,
and export query results — all driven by your SQLAlchemy ORM models.

Example:
    >>> from sqlalchemy_excel import ExcelMapping, ExcelTemplate
    >>> mapping = ExcelMapping.from_model(User)
    >>> template = ExcelTemplate([mapping])
    >>> template.save("users_template.xlsx")
"""

from __future__ import annotations

from sqlalchemy_excel._compat import ensure_defusedxml
from sqlalchemy_excel.exceptions import (
    ConstraintViolationError,
    DuplicateKeyError,
    ExportError,
    FileFormatError,
    HeaderMismatchError,
    ImportError_,
    MappingError,
    ReaderError,
    SheetNotFoundError,
    SqlalchemyExcelError,
    TemplateError,
    ValidationError,
)

# Ensure defusedxml is available at import time
ensure_defusedxml()

# Lazy imports for main API classes to avoid circular imports
# and to keep import time fast


def __getattr__(name: str) -> object:
    """Lazy import public API classes."""
    if name == "ExcelMapping":
        from sqlalchemy_excel.mapping import ExcelMapping

        return ExcelMapping
    if name == "ColumnMapping":
        from sqlalchemy_excel.mapping import ColumnMapping

        return ColumnMapping
    if name == "ExcelTemplate":
        from sqlalchemy_excel.template import ExcelTemplate

        return ExcelTemplate
    if name == "ExcelWorkbookSession":
        from sqlalchemy_excel.excelio import ExcelWorkbookSession

        return ExcelWorkbookSession
    if name == "ExcelValidator":
        from sqlalchemy_excel.validation import ExcelValidator

        return ExcelValidator
    if name == "ValidationReport":
        from sqlalchemy_excel.validation import ValidationReport

        return ValidationReport
    if name == "CellError":
        from sqlalchemy_excel.validation import CellError

        return CellError
    if name == "ExcelImporter":
        from sqlalchemy_excel.load import ExcelImporter

        return ExcelImporter
    if name == "ImportResult":
        from sqlalchemy_excel.load import ImportResult

        return ImportResult
    if name == "ExcelExporter":
        from sqlalchemy_excel.export import ExcelExporter

        return ExcelExporter
    raise AttributeError(f"module 'sqlalchemy_excel' has no attribute {name!r}")


__all__ = [
    "CellError",
    "ColumnMapping",
    "ConstraintViolationError",
    "DuplicateKeyError",
    "ExcelExporter",
    "ExcelImporter",
    "ExcelMapping",
    "ExcelTemplate",
    "ExcelValidator",
    "ExcelWorkbookSession",
    "ExportError",
    "FileFormatError",
    "HeaderMismatchError",
    "ImportError_",
    "ImportResult",
    "MappingError",
    "ReaderError",
    "SheetNotFoundError",
    "SqlalchemyExcelError",
    "TemplateError",
    "ValidationError",
    "ValidationReport",
]
