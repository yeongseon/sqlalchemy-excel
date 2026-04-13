"""ExcelDialect — SQLAlchemy dialect that uses excel-dbapi as the DB-API driver."""

from __future__ import annotations

import re
from typing import TYPE_CHECKING, Any, Literal, cast
from urllib.parse import unquote as _url_unquote

from sqlalchemy import event, pool
from sqlalchemy.engine import default
from sqlalchemy.schema import Table

from .compiler import ExcelCompiler, ExcelIdentifierPreparer
from .ddl import ExcelDDLCompiler
from .reflection import ExcelInspectionMixin
from .types import ExcelTypeCompiler

if TYPE_CHECKING:
    from sqlalchemy.engine import URL
    from sqlalchemy.engine.interfaces import ConnectArgsType
    from sqlalchemy.sql.compiler import IdentifierPreparer


def _after_create(
    target: Table,
    connection: Any,
    **kw: Any,
) -> None:
    """Write column metadata after CREATE TABLE (Excel dialect only)."""
    if connection.dialect.name != "excel":
        return
    import excel_dbapi

    raw_conn = connection.connection.dbapi_connection
    pk_cols = {col.name for col in target.primary_key.columns}

    columns = []
    for col in target.columns:
        type_compiler = connection.dialect.type_compiler
        type_name = type_compiler.process(col.type)
        columns.append(
            {
                "name": col.name,
                "type_name": type_name,
                "nullable": col.nullable if col.nullable is not None else True,
                "primary_key": col.name in pk_cols,
            }
        )

    excel_dbapi.write_table_metadata(raw_conn, target.name, columns)


def _after_drop(
    target: Table,
    connection: Any,
    **kw: Any,
) -> None:
    """Remove column metadata after DROP TABLE (Excel dialect only)."""
    if connection.dialect.name != "excel":
        return
    import excel_dbapi

    raw_conn = connection.connection.dbapi_connection
    excel_dbapi.remove_table_metadata(raw_conn, target.name)


# Register DDL events globally for all Table objects
event.listen(Table, "after_create", _after_create)
event.listen(Table, "after_drop", _after_drop)


class ExcelDialect(  # type: ignore[misc]  # pyright: ignore[reportIncompatibleMethodOverride]
    ExcelInspectionMixin,
    default.DefaultDialect,
):
    """SQLAlchemy dialect for Excel files via excel-dbapi.

    Connection URLs::

        # Relative path
        excel:///data.xlsx

        # Absolute path
        excel:////home/user/data.xlsx
        excel:///C:/Users/data.xlsx  (Windows)

    """

    name: str = "excel"
    driver: str = "dbapi"
    default_paramstyle: str = "qmark"

    # ── Feature flags ──────────────────────────────────────
    supports_alter: bool = False
    supports_sequences: bool = False
    supports_schemas: bool = False
    supports_views: bool = False
    supports_native_boolean: bool = True
    supports_native_decimal: bool = False
    supports_statement_cache: bool = False
    supports_default_values: bool = False
    supports_default_metavalue: bool = False
    supports_empty_insert: bool = False
    supports_multivalues_insert: bool = True
    postfetch_lastrowid: bool = False
    insertmanyvalues_implicit_sentinel: Any = None

    # ── Compiler classes ──────────────────────────────────
    statement_compiler = ExcelCompiler
    ddl_compiler = ExcelDDLCompiler
    type_compiler_cls = ExcelTypeCompiler
    preparer: type[IdentifierPreparer] = cast(
        "type[IdentifierPreparer]", ExcelIdentifierPreparer
    )

    @classmethod
    def import_dbapi(cls) -> Any:
        import excel_dbapi

        return excel_dbapi

    @classmethod
    def get_pool_class(cls, url: URL) -> type[pool.Pool]:
        return pool.StaticPool

    def create_connect_args(self, url: URL) -> ConnectArgsType:
        """Translate a SQLAlchemy URL to excel-dbapi connect() arguments.

        URL formats:
            excel:///relative/path.xlsx   →  file_path="relative/path.xlsx"
            excel:////absolute/path.xlsx  →  file_path="/absolute/path.xlsx"
        """
        # url.database contains the path after the third slash
        file_path = url.database
        if not file_path:
            raise ValueError("No file path in URL. Use excel:///path/to/file.xlsx")

        kwargs: dict[str, Any] = {
            "file_path": file_path,
            "engine": "openpyxl",
            "autocommit": False,
            "create": True,
        }

        # Forward query parameters
        query = dict(url.query)
        if "engine" in query:
            kwargs["engine"] = query.pop("engine")
        if "autocommit" in query:
            autocommit = query.pop("autocommit")
            if isinstance(autocommit, tuple):
                autocommit_text = autocommit[0] if autocommit else ""
            else:
                autocommit_text = autocommit
            kwargs["autocommit"] = autocommit_text.lower() in (
                "true",
                "1",
                "yes",
            )

        return ([], kwargs)

    def on_connect(self) -> None:
        """No-op: no special connection initialization needed."""

    def do_execute(
        self,
        cursor: Any,
        statement: str,
        parameters: Any,
        context: Any = None,
    ) -> None:
        """Execute a statement, normalizing whitespace for excel-dbapi."""
        normalized = re.sub(r"\s+", " ", statement).strip()
        cursor.execute(normalized, parameters)

    def do_execute_no_params(
        self,
        cursor: Any,
        statement: str,
        context: Any = None,
    ) -> None:
        """Execute a statement with no parameters."""
        normalized = re.sub(r"\s+", " ", statement).strip()
        cursor.execute(normalized, None)

    def do_ping(self, dbapi_connection: Any) -> bool:
        """Ping the connection by verifying it's not closed."""
        return not getattr(dbapi_connection, "closed", True)

    def is_disconnect(self, e: Exception, connection: Any, cursor: Any) -> bool:
        """Excel connections don't have network-level disconnects."""
        return False

    def get_default_isolation_level(self, dbapi_conn: Any) -> Literal["SERIALIZABLE"]:
        return "SERIALIZABLE"

    def _check_unicode_returns(
        self, connection: Any, additional_tests: Any = None
    ) -> bool:
        return True

    def _check_unicode_description(self, connection: Any) -> bool:
        return True

    def has_table(
        self,
        connection: Any,
        table_name: str,
        schema: str | None = None,
        **kw: Any,
    ) -> bool:
        """Check if a worksheet (table) exists."""
        import excel_dbapi

        raw_conn = connection.connection.dbapi_connection
        return cast("bool", excel_dbapi.has_table(raw_conn, table_name))

    def do_begin(self, dbapi_connection: Any) -> None:
        """No-op: excel-dbapi doesn't have explicit BEGIN."""

    def do_commit(self, dbapi_connection: Any) -> None:
        """Commit (save) the workbook."""
        dbapi_connection.commit()

    def do_rollback(self, dbapi_connection: Any) -> None:
        """Rollback via the DB-API driver's snapshot/restore implementation.

        The excel-dbapi openpyxl backend implements transactional rollback by
        restoring the latest committed snapshot. Some non-transactional
        backends (for example Graph/autocommit connections) raise
        ``NotSupportedError``; we treat that as a no-op so pool reset does not
        fail on connection return.
        """
        from excel_dbapi.exceptions import NotSupportedError

        try:
            dbapi_connection.rollback()
        except NotSupportedError:
            pass

    def do_close(self, dbapi_connection: Any) -> None:
        """Close the underlying excel-dbapi connection."""
        dbapi_connection.close()


class ExcelGraphDialect(ExcelDialect):  # type: ignore[misc,unused-ignore]
    """SQLAlchemy dialect for remote Excel files via Microsoft Graph API.

    Connection URLs::

        # With drive_id and item_id
        excel+graph:///drive_id/item_id

        # With query parameters
        excel+graph:///drive_id/item_id?readonly=false

    Credentials must be passed via ``connect_args``::

        engine = create_engine(
            "excel+graph:///drive_id/item_id",
            connect_args={"credential": DefaultAzureCredential()},
        )
    """

    driver: str = "graph"
    supports_statement_cache: bool = False

    def create_connect_args(self, url: URL) -> ConnectArgsType:
        """Translate an excel+graph:// URL to excel-dbapi connect() arguments.

        URL format: excel+graph:///drive_id/item_id
        Maps to DSN: msgraph://drives/{drive_id}/items/{item_id}
        """
        database = url.database
        if not database:
            raise ValueError(
                "No drive/item path in URL. Use excel+graph:///drive_id/item_id"
            )

        parts = database.strip("/").split("/")
        if len(parts) != 2:
            raise ValueError(
                f"Expected excel+graph:///drive_id/item_id (got {len(parts)} path segments: {database!r})"
            )

        drive_id = _url_unquote(parts[0])
        item_id = _url_unquote(parts[1])
        dsn = f"msgraph://drives/{drive_id}/items/{item_id}"

        kwargs = {
            "file_path": dsn,
            "engine": "graph",
            "autocommit": True,
            "create": False,
        }

        query = dict(url.query)
        if "readonly" in query:
            raw = query.pop("readonly")
            val = raw[0] if isinstance(raw, tuple) else raw
            kwargs["readonly"] = str(val).lower() in ("true", "1", "yes")

        return ([], kwargs)
