"""sqlalchemy-excel — SQLAlchemy dialect for Excel files."""

from __future__ import annotations

from .dialect import ExcelDialect, ExcelGraphDialect

__version__ = "0.2.2"

__all__ = [
    "ExcelDialect",
    "ExcelGraphDialect",
    "__version__",
]
