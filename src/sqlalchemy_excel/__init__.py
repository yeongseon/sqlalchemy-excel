"""sqlalchemy-excel — SQLAlchemy dialect for Excel files."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version

from .dialect import ExcelDialect, ExcelGraphDialect
from .dml import Insert, insert

try:
    __version__ = version("sqlalchemy-excel")
except PackageNotFoundError:
    __version__ = "0.5.4"

__all__ = [
    "ExcelDialect",
    "ExcelGraphDialect",
    "Insert",
    "__version__",
    "insert",
]
