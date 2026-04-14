"""sqlalchemy-excel — SQLAlchemy dialect for Excel files."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version

from sqlalchemy.dialects import registry

from .dialect import ExcelDialect, ExcelGraphDialect
from .dml import Insert, insert

registry.register("excel", "sqlalchemy_excel.dialect", "ExcelDialect")
registry.register("excel.graph", "sqlalchemy_excel.dialect", "ExcelGraphDialect")

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
