"""sqlalchemy-excel — SQLAlchemy dialect for Excel files."""

from __future__ import annotations

from .dialect import ExcelDialect, ExcelGraphDialect

__version__ = "0.4.0"

__all__ = [
    "ExcelDialect",
    "ExcelGraphDialect",
    "__version__",
]
