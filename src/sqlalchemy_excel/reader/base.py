"""Core reader protocol and shared reader utilities."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import TYPE_CHECKING, Protocol

if TYPE_CHECKING:
    from collections.abc import Iterable

    from sqlalchemy_excel._types import FileSource, RowDict


@dataclass(slots=True)
class ReaderResult:
    """Container for parsed Excel reader output.

    Attributes:
        headers: Normalized column headers extracted from the sheet.
        rows: Row iterator yielding mapping objects keyed by normalized headers.
        total_rows: Number of parsed data rows if known eagerly; ``None`` for streaming.
    """

    headers: list[str]
    rows: Iterable[RowDict]
    total_rows: int | None


class BaseReader(Protocol):
    """Protocol for Excel reader backends."""

    def read(
        self,
        source: FileSource,
        sheet_name: str | None = None,
        header_row: int | None = None,
    ) -> ReaderResult: ...


_NON_HEADER_SAFE_PATTERN = re.compile(r"[^0-9a-z_]")


def normalize_header(header: str) -> str:
    """Normalize an Excel header value to a predictable column identifier.

    Normalization rules:
    1. Strip leading/trailing whitespace
    2. Lowercase
    3. Replace spaces with underscores
    4. Remove non-alphanumeric characters except underscores

    Args:
        header: Raw header string from the worksheet.

    Returns:
        Normalized header string.
    """

    cleaned = header.strip().lower().replace(" ", "_")
    return _NON_HEADER_SAFE_PATTERN.sub("", cleaned)
