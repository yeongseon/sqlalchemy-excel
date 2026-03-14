"""ExcelWorkbookSession dual-channel wrapper for excel-dbapi."""

from __future__ import annotations

import importlib
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from pathlib import Path

excel_dbapi: Any = importlib.import_module("excel_dbapi")


class ExcelWorkbookSession:
    """Dual-channel session wrapping an excel-dbapi connection.

    Provides two access channels to an Excel file:
    - Data channel (``cursor``): DB-API cursor for SQL-based reads and writes.
    - Format channel (``workbook``): openpyxl workbook for style-oriented updates.

    Args:
        conn: An open excel-dbapi connection.
    """

    def __init__(self, conn: Any) -> None:
        """Initialize session from an open excel-dbapi connection.

        Args:
            conn: An open excel-dbapi connection object.
        """

        self._conn: Any = conn
        self._cursor: Any = conn.cursor()

    @classmethod
    def open(
        cls,
        path: str | Path,
        *,
        create: bool = False,
        data_only: bool = False,
    ) -> ExcelWorkbookSession:
        """Open an Excel file via excel-dbapi and return a session.

        Args:
            path: Path to the Excel file.
            create: If ``True``, create the file if it does not exist.
            data_only: If ``True``, read cached formula values.

        Returns:
            A new ``ExcelWorkbookSession`` instance.
        """

        conn = excel_dbapi.connect(
            str(path),
            engine="openpyxl",
            autocommit=False,
            create=create,
            data_only=data_only,
        )
        return cls(conn)

    @property
    def conn(self) -> Any:
        """Return the underlying excel-dbapi connection."""

        return self._conn

    @property
    def cursor(self) -> Any:
        """Return a DB-API cursor for SQL-based data operations."""

        return self._cursor

    @property
    def workbook(self) -> Any:
        """Return the openpyxl workbook managed by the connection."""

        return self._conn.workbook

    def commit(self) -> None:
        """Commit pending workbook changes to disk."""

        self._conn.commit()

    def rollback(self) -> None:
        """Rollback pending workbook changes."""

        self._conn.rollback()

    def close(self) -> None:
        """Close the underlying connection."""

        self._conn.close()

    def __enter__(self) -> ExcelWorkbookSession:
        """Enter context manager scope and return session instance."""

        return self

    def __exit__(self, exc_type: object, exc_val: object, exc_tb: object) -> None:
        """Exit context manager scope and close the session."""

        self.close()
