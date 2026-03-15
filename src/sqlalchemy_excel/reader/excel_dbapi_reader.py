"""Excel reader backed by excel-dbapi SQL interface."""

from __future__ import annotations

import importlib
import os
import tempfile
from contextlib import suppress
from pathlib import Path
from typing import TYPE_CHECKING, Any
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from sqlalchemy_excel.exceptions import FileFormatError, ReaderError, SheetNotFoundError
from sqlalchemy_excel.reader.base import ReaderResult, normalize_header

if TYPE_CHECKING:
    from sqlalchemy_excel._types import FileSource, RowDict

excel_dbapi: Any = importlib.import_module("excel_dbapi")


class ExcelDbapiReader:
    """Reader backend using excel-dbapi for SQL-based Excel data access.

    Args:
        read_only: API parity with ``OpenpyxlReader``.
        max_file_size: Maximum allowed file size in bytes.
    """

    def __init__(
        self,
        *,
        read_only: bool = False,
        max_file_size: int = 50 * 1024 * 1024,
    ) -> None:
        self.read_only: bool = read_only
        self.max_file_size: int = max_file_size

    def read(
        self,
        source: FileSource,
        sheet_name: str | None = None,
        header_row: int | None = None,
    ) -> ReaderResult:
        """Read rows from an Excel worksheet via SQL.

        Args:
            source: File path or binary stream.
            sheet_name: Worksheet name. Defaults to first sheet.
            header_row: Header-row setting kept for API parity.

        Returns:
            ReaderResult with normalized headers and row dicts.

        Raises:
            ReaderError: On file access and generic read errors.
            FileFormatError: If source is not a valid .xlsx file.
            SheetNotFoundError: If target sheet does not exist.
        """

        del header_row

        file_path, remove_after_read = self._resolve_source(source)
        self._validate_file_size_path(file_path)

        conn: Any | None = None
        try:
            conn = excel_dbapi.connect(
                str(file_path),
                engine="openpyxl",
                autocommit=True,
                data_only=True,
            )
            if conn is None:
                raise ReaderError("Unable to open Excel source")

            connection = conn

            workbook = connection.workbook
            available_sheets: list[str] = list(workbook.sheetnames)
            if not available_sheets:
                raise ReaderError("Workbook contains no sheets")

            resolved_sheet = sheet_name or available_sheets[0]
            if resolved_sheet not in available_sheets:
                raise SheetNotFoundError(resolved_sheet, available_sheets)

            cursor = connection.cursor()
            cursor.execute(f"SELECT * FROM {resolved_sheet}")

            description = cursor.description
            if description is None:
                return ReaderResult(headers=[], rows=[], total_rows=0)

            headers = self._normalize_headers([str(item[0]) for item in description])

            rows: list[RowDict] = []
            for raw_row in cursor.fetchall():
                values = list(raw_row)
                if all(self._is_empty_cell(value) for value in values):
                    continue
                row_dict: RowDict = dict(zip(headers, values, strict=True))
                rows.append(row_dict)

            return ReaderResult(headers=headers, rows=rows, total_rows=len(rows))
        except (InvalidFileException, BadZipFile) as exc:
            raise FileFormatError("Input is not a valid .xlsx file") from exc
        except (SheetNotFoundError, ReaderError, FileFormatError):
            raise
        except Exception as exc:
            raise ReaderError(f"Failed to read Excel data: {exc}") from exc
        finally:
            if conn is not None:
                conn.close()
            if remove_after_read:
                self._remove_temp_file(file_path)

    def _resolve_source(self, source: FileSource) -> tuple[str, bool]:
        """Convert file source to path, using a temp file for streams."""

        if isinstance(source, str | Path | os.PathLike):
            return str(source), False

        binary_source = source
        try:
            original_position = binary_source.tell()
        except (AttributeError, OSError) as exc:
            raise ReaderError("Unable to determine binary source position") from exc

        try:
            content = binary_source.read()
            _ = binary_source.seek(original_position)
        except (AttributeError, OSError) as exc:
            raise ReaderError("Unable to read binary Excel source") from exc

        if not isinstance(content, bytes):
            raise ReaderError("Binary source did not return bytes")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            _ = tmp.write(content)
            tmp.flush()
            temp_path = tmp.name

        return temp_path, True

    def _validate_file_size_path(self, path: str) -> None:
        """Validate file size against configured limit."""

        try:
            size = os.path.getsize(path)
        except OSError as exc:
            raise ReaderError(f"Unable to access Excel source: {exc}") from exc

        if size > self.max_file_size:
            raise ReaderError(
                f"File size ({size:,} bytes) exceeds maximum "
                f"({self.max_file_size:,} bytes)"
            )

    @staticmethod
    def _normalize_headers(raw_headers: list[str]) -> list[str]:
        headers: list[str] = []
        for index, header in enumerate(raw_headers, start=1):
            normalized = normalize_header(header)
            if not normalized:
                raise ReaderError(
                    f"Header value is empty after normalization at column {index}"
                )
            if normalized in headers:
                raise ReaderError(
                    f"Duplicate normalized header detected: '{normalized}'"
                )
            headers.append(normalized)
        return headers

    @staticmethod
    def _is_empty_cell(value: object) -> bool:
        if value is None:
            return True
        if isinstance(value, str):
            return not value.strip()
        return False

    @staticmethod
    def _remove_temp_file(path: str) -> None:
        with suppress(OSError):
            os.unlink(path)
