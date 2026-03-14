"""Custom exception hierarchy for sqlalchemy-excel."""

from __future__ import annotations


class SqlalchemyExcelError(Exception):
    """Base exception for all sqlalchemy-excel errors."""


class MappingError(SqlalchemyExcelError):
    """Raised when ORM model introspection fails.

    Examples:
        - Model has no mapped columns
        - Unsupported column type encountered
        - Ambiguous column mapping
    """


class TemplateError(SqlalchemyExcelError):
    """Raised when Excel template generation fails."""


class ReaderError(SqlalchemyExcelError):
    """Base exception for Excel file reading errors."""


class FileFormatError(ReaderError):
    """Raised when input file is not a valid .xlsx file."""


class SheetNotFoundError(ReaderError):
    """Raised when specified sheet name doesn't exist."""

    def __init__(self, sheet_name: str, available: list[str]) -> None:
        self.sheet_name = sheet_name
        self.available = available
        super().__init__(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {', '.join(available)}"
        )


class HeaderMismatchError(ReaderError):
    """Raised when Excel headers don't match expected columns."""

    def __init__(
        self,
        missing: list[str],
        extra: list[str],
    ) -> None:
        self.missing = missing
        self.extra = extra
        parts: list[str] = []
        if missing:
            parts.append(f"Missing columns: {', '.join(missing)}")
        if extra:
            parts.append(f"Unexpected columns: {', '.join(extra)}")
        super().__init__(". ".join(parts))


class ValidationError(SqlalchemyExcelError):
    """Raised when data validation fails.

    Contains a ValidationReport with detailed errors.
    """

    def __init__(self, report: object) -> None:
        self.report = report
        super().__init__(str(report))


class ImportError_(SqlalchemyExcelError):  # noqa: N801, N818
    """Raised when database import fails.

    Underscore suffix to avoid shadowing builtin ImportError.
    """


class DuplicateKeyError(ImportError_):
    """Raised when an insert violates a unique constraint."""


class ConstraintViolationError(ImportError_):
    """Raised when data violates a database constraint."""


class ExportError(SqlalchemyExcelError):
    """Raised when Excel export fails."""
