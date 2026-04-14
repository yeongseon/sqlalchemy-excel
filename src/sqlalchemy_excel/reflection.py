"""Reflection (introspection) support for the Excel dialect.

Provides the Inspector integration so that SQLAlchemy can discover
tables (worksheets), columns, and primary keys from an Excel file.
"""

from __future__ import annotations

import logging
from contextlib import suppress
from typing import Any, cast

from sqlalchemy import types as sa_types
from sqlalchemy.exc import NoSuchTableError

_TYPE_MAP: dict[str, type[sa_types.TypeEngine[Any]]] = {
    "TEXT": sa_types.String,
    "INTEGER": sa_types.Integer,
    "SMALLINT": sa_types.Integer,
    "BIGINT": sa_types.Integer,
    "FLOAT": sa_types.Float,
    "REAL": sa_types.Float,
    "DECIMAL": sa_types.Float,
    "NUMERIC": sa_types.Float,
    "DOUBLE": sa_types.Float,
    "DOUBLE PRECISION": sa_types.Float,
    "BOOLEAN": sa_types.Boolean,
    "DATE": sa_types.Date,
    "DATETIME": sa_types.DateTime,
    "TIMESTAMP": sa_types.DateTime,
}

_LOG = logging.getLogger(__name__)


def _sa_type_from_name(type_name: str) -> sa_types.TypeEngine[Any]:
    """Convert an excel-dbapi type name to a SQLAlchemy type instance."""
    cls = _TYPE_MAP.get(type_name.upper(), sa_types.String)
    return cls()


class ExcelInspectionMixin:
    """Mixin that provides reflection methods for ExcelDialect.

    These methods are called by SQLAlchemy's Inspector to discover
    the structure of an Excel workbook.
    """

    def get_table_names(
        self,
        connection: Any,
        schema: str | None = None,
        **kw: Any,
    ) -> list[str]:
        """Return all worksheet names, excluding the metadata sheet."""
        import excel_dbapi

        if schema is not None:
            return []

        raw_conn = connection.connection.dbapi_connection
        return cast("list[str]", excel_dbapi.list_tables(raw_conn, include_meta=False))

    def get_view_names(
        self,
        connection: Any,
        schema: str | None = None,
        **kw: Any,
    ) -> list[str]:
        """Excel does not support views."""
        return []

    def get_columns(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Return column information for the given table (worksheet).

        First tries to read from the metadata sheet (written by CREATE TABLE).
        Falls back to type inference from data sampling.
        """
        import excel_dbapi

        if schema is not None:
            raise NoSuchTableError(table_name)

        raw_conn = self._raw_connection(connection)
        table_exists = cast("bool", excel_dbapi.has_table(raw_conn, table_name))
        if not table_exists:
            self._clear_stale_metadata_if_present(raw_conn, table_name)
            raise NoSuchTableError(table_name)

        header_names = self._worksheet_header_names(raw_conn, table_name)
        meta = self._validated_metadata(raw_conn, table_name, header_names)

        if meta is not None:
            return [
                {
                    "name": col["name"],
                    "type": _sa_type_from_name(col["type_name"]),
                    "nullable": col.get("nullable", True),
                    "default": None,
                    "autoincrement": False,
                    "comment": None,
                }
                for col in meta
            ]

        # Fallback: infer from data
        inferred = excel_dbapi.get_columns(raw_conn, table_name)
        return [
            {
                "name": col["name"],
                "type": _sa_type_from_name(col["type"]),
                "nullable": col.get("nullable", True),
                "default": None,
                "autoincrement": False,
                "comment": None,
            }
            for col in inferred
        ]

    def get_pk_constraint(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> dict[str, Any]:
        """Return primary key constraint info from the metadata sheet."""
        import excel_dbapi

        if schema is not None:
            raise NoSuchTableError(table_name)

        raw_conn = self._raw_connection(connection)
        table_exists = cast("bool", excel_dbapi.has_table(raw_conn, table_name))
        if not table_exists:
            self._clear_stale_metadata_if_present(raw_conn, table_name)
            raise NoSuchTableError(table_name)

        header_names = self._worksheet_header_names(raw_conn, table_name)
        meta = self._validated_metadata(raw_conn, table_name, header_names)

        if meta is not None:
            pk_cols = [col["name"] for col in meta if col.get("primary_key", False)]
            if pk_cols:
                return {"constrained_columns": pk_cols, "name": None}

        return {"constrained_columns": [], "name": None}

    def get_foreign_keys(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Excel does not support foreign keys."""
        if schema is not None:
            raise NoSuchTableError(table_name)
        self._assert_table_exists(connection, table_name)
        return []

    def get_indexes(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Excel does not support indexes."""
        if schema is not None:
            raise NoSuchTableError(table_name)
        self._assert_table_exists(connection, table_name)
        return []

    def get_unique_constraints(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Excel does not support unique constraints."""
        if schema is not None:
            raise NoSuchTableError(table_name)
        self._assert_table_exists(connection, table_name)
        return []

    def get_check_constraints(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Excel does not support check constraints."""
        if schema is not None:
            raise NoSuchTableError(table_name)
        self._assert_table_exists(connection, table_name)
        return []

    def get_table_comment(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> dict[str, Any]:
        """Excel does not support table comments."""
        if schema is not None:
            raise NoSuchTableError(table_name)
        self._assert_table_exists(connection, table_name)
        return {"text": None}

    def get_schema_names(self, connection: Any, **kw: Any) -> list[str]:
        """Excel does not support schemas."""
        return []

    @staticmethod
    def _raw_connection(connection: Any) -> Any:
        return connection.connection.dbapi_connection

    @staticmethod
    def _metadata_header_matches(
        meta: list[dict[str, Any]],
        header_names: list[str] | None,
    ) -> bool:
        if header_names is None:
            return False
        # Check that metadata column names match live headers
        # AND metadata column count matches live column count.
        # This guards against the edge case where a sheet is deleted and
        # recreated with the same header names but different types/nullability/PKs.
        return [col["name"] for col in meta] == header_names and len(meta) == len(
            header_names
        )

    @staticmethod
    def _cursor_header_names(raw_conn: Any, table_name: str) -> list[str] | None:
        try:
            cursor = raw_conn.cursor()
        except Exception:
            return None

        try:
            cursor.execute(f"SELECT * FROM {table_name} LIMIT 0")
            description = getattr(cursor, "description", None)
            if not description:
                return None
            header_names = [
                str(column[0])
                for column in description
                if column is not None and column[0] is not None
            ]
            return header_names or None
        except Exception:
            return None
        finally:
            with suppress(Exception):
                cursor.close()

    def _validated_metadata(
        self,
        raw_conn: Any,
        table_name: str,
        header_names: list[str] | None,
    ) -> list[dict[str, Any]] | None:
        import excel_dbapi

        meta = excel_dbapi.read_table_metadata(raw_conn, table_name)
        if meta is None:
            return None

        if header_names is None:
            header_names = self._cursor_header_names(raw_conn, table_name)
            if header_names is None:
                _LOG.warning(
                    "Could not validate metadata headers for table '%s'; using stored metadata",
                    table_name,
                )
                return meta

        if not self._metadata_header_matches(meta, header_names):
            excel_dbapi.remove_table_metadata(raw_conn, table_name)
            return None

        return meta

    @staticmethod
    def _clear_stale_metadata_if_present(raw_conn: Any, table_name: str) -> None:
        import excel_dbapi

        meta = excel_dbapi.read_table_metadata(raw_conn, table_name)
        if meta is not None:
            excel_dbapi.remove_table_metadata(raw_conn, table_name)

    @staticmethod
    def _worksheet_header_names(raw_conn: Any, table_name: str) -> list[str] | None:
        from excel_dbapi.exceptions import NotSupportedError

        try:
            workbook = getattr(raw_conn, "workbook", None)
        except NotSupportedError:
            return None
        if workbook is None:
            return None

        try:
            worksheet = workbook[table_name]
        except (KeyError, NotSupportedError):
            return None

        headers: list[str] = []
        for index in range(1, worksheet.max_column + 1):
            value = worksheet.cell(row=1, column=index).value
            if value is not None:
                headers.append(str(value))
        return headers

    def _assert_table_exists(self, connection: Any, table_name: str) -> None:
        import excel_dbapi

        raw_conn = self._raw_connection(connection)
        if not cast("bool", excel_dbapi.has_table(raw_conn, table_name)):
            self._clear_stale_metadata_if_present(raw_conn, table_name)
            raise NoSuchTableError(table_name)
