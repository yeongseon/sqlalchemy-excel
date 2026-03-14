"""Internal type aliases for sqlalchemy-excel."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any, BinaryIO, Union

if TYPE_CHECKING:
    from os import PathLike

# Path-like types accepted by file operations
FilePath = Union[str, "PathLike[str]"]

# Source types accepted by readers
FileSource = Union[str, "PathLike[str]", BinaryIO]

# A single row of data from Excel
RowDict = dict[str, Any]

# Column name
ColumnName = str
