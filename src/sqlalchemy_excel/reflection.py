"""Reflection (introspection) support for the Excel dialect.

Provides the Inspector integration so that SQLAlchemy can discover
tables (worksheets), columns, and primary keys from an Excel file.
"""

from __future__ import annotations

from typing import Any, cast

from sqlalchemy import types as sa_types

_TYPE_MAP: dict[str, type[sa_types.TypeEngine[Any]]] = {
    "TEXT": sa_types.String,
    "INTEGER": sa_types.Integer,
    "FLOAT": sa_types.Float,
    "BOOLEAN": sa_types.Boolean,
    "DATE": sa_types.Date,
    "DATETIME": sa_types.DateTime,
}


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

        raw_conn = connection.connection.dbapi_connection

        # Try metadata sheet first
        meta = excel_dbapi.read_table_metadata(raw_conn, table_name)
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

        raw_conn = connection.connection.dbapi_connection
        meta = excel_dbapi.read_table_metadata(raw_conn, table_name)
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
        return []

    def get_indexes(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Excel does not support indexes."""
        return []

    def get_unique_constraints(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Excel does not support unique constraints."""
        return []

    def get_check_constraints(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> list[dict[str, Any]]:
        """Excel does not support check constraints."""
        return []

    def get_table_comment(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> dict[str, Any]:
        """Excel does not support table comments."""
        return {"text": None}

    def get_schema_names(self, connection: Any, **kw: Any) -> list[str]:
        """Excel does not support schemas."""
        return []
