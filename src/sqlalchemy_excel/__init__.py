"""sqlalchemy-excel — SQLAlchemy dialect for Excel files."""

from __future__ import annotations

from .dialect import ExcelDialect

__version__ = "0.1.1"

__all__ = [
    "ExcelDialect",
    "__version__",
]
