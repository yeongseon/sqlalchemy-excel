"""ExcelDialect — SQLAlchemy dialect that uses excel-dbapi as the DB-API driver."""

from __future__ import annotations

import warnings
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

    for col in target.columns:
        if col.server_default is not None:
            warnings.warn(
                "Excel dialect does not support server_default; value will be ignored",
                stacklevel=2,
            )
        if col.autoincrement is True:
            warnings.warn(
                "Excel dialect does not support autoincrement=True; value must be set explicitly",
                stacklevel=2,
            )

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


# Register DDL events on the Table class.
#
# This is the standard pattern for third-party SQLAlchemy dialects that need
# post-DDL hooks (cf. sqlalchemy-exasol, sqlalchemy-bigquery).  The listeners
# guard on connection.dialect.name so they are a no-op for every non-Excel
# engine in the process.  propagate=False prevents inheritance propagation.
event.listen(Table, "after_create", _after_create, propagate=False)
event.listen(Table, "after_drop", _after_drop, propagate=False)


def _normalize_statement_whitespace_quote_aware(statement: str) -> str:
    out: list[str] = []
    in_quote = False
    quote_char = ""
    i = 0
    length = len(statement)

    while i < length:
        ch = statement[i]

        if in_quote:
            out.append(ch)
            if ch == quote_char:
                if i + 1 < length and statement[i + 1] == quote_char:
                    out.append(statement[i + 1])
                    i += 1
                else:
                    in_quote = False
            i += 1
            continue

        if ch in ("'", '"'):
            in_quote = True
            quote_char = ch
            out.append(ch)
            i += 1
            continue

        if ch.isspace():
            if out and out[-1] != " ":
                out.append(" ")
            i += 1
            continue

        out.append(ch)
        i += 1

    return "".join(out).strip()


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
    supports_alter: bool = True
    supports_sequences: bool = False
    supports_schemas: bool = False
    supports_views: bool = False
    supports_native_boolean: bool = True
    supports_native_decimal: bool = False
    # NOTE(issue #59): keep statement cache disabled for now.
    #
    # ExcelCompiler mutates traversal-scoped state (`_has_join`, `_in_in_clause`,
    # `_subquery_depth`) while rendering SQL. Those flags directly affect emitted
    # text (for example whether columns are table-qualified). The current
    # implementation appears deterministic for equivalent statement trees, but we
    # keep cache opt-in conservative until we have dedicated cache-key
    # determinism coverage for these stateful paths.
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
        file_path = _url_unquote(url.database) if url.database else None
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
        normalized = _normalize_statement_whitespace_quote_aware(statement)
        cursor.execute(normalized, parameters)
        self._sync_alter_table_metadata(cursor, normalized)

    def do_execute_no_params(
        self,
        cursor: Any,
        statement: str,
        context: Any = None,
    ) -> None:
        """Execute a statement with no parameters."""
        normalized = _normalize_statement_whitespace_quote_aware(statement)
        cursor.execute(normalized, None)
        self._sync_alter_table_metadata(cursor, normalized)

    def _sync_alter_table_metadata(self, cursor: Any, statement: str) -> None:
        if not statement.upper().startswith("ALTER TABLE "):
            return

        import excel_dbapi

        tokens = statement.split()
        if len(tokens) < 6:
            return

        table_name = tokens[2].strip('"')
        operation = tokens[3].upper()

        raw_conn = cursor.connection
        current_meta = excel_dbapi.read_table_metadata(raw_conn, table_name) or []

        type_map = {col["name"]: col["type_name"] for col in current_meta}
        nullable_map = {col["name"]: col.get("nullable", True) for col in current_meta}
        pk_map = {col["name"]: col.get("primary_key", False) for col in current_meta}

        if operation == "ADD" and len(tokens) >= 7 and tokens[4].upper() == "COLUMN":
            col_name = tokens[5].strip('"')
            added_type = tokens[6].upper()
            if added_type == "FLOAT":
                added_type = "REAL"
            type_map[col_name] = added_type
            # Preserve nullable/PK hints from trailing constraints:
            # ALTER TABLE t ADD COLUMN c TYPE NOT NULL PRIMARY KEY
            tail = " ".join(t.upper() for t in tokens[7:])
            nullable_map[col_name] = "NOT NULL" not in tail
            pk_map[col_name] = "PRIMARY KEY" in tail or "PRIMARY_KEY" in tail

        if operation == "DROP" and len(tokens) == 6 and tokens[4].upper() == "COLUMN":
            removed = tokens[5].strip('"')
            type_map.pop(removed, None)
            nullable_map.pop(removed, None)
            pk_map.pop(removed, None)

        if (
            operation == "RENAME"
            and len(tokens) == 8
            and tokens[4].upper() == "COLUMN"
            and tokens[6].upper() == "TO"
        ):
            old_name = tokens[5].strip('"')
            new_name = tokens[7].strip('"')
            if old_name in type_map:
                type_map[new_name] = type_map.pop(old_name)
            if old_name in nullable_map:
                nullable_map[new_name] = nullable_map.pop(old_name)
            if old_name in pk_map:
                pk_map[new_name] = pk_map.pop(old_name)

        live_columns = excel_dbapi.get_columns(raw_conn, table_name)
        columns = [
            {
                "name": col["name"],
                "type_name": type_map.get(col["name"], col.get("type", "TEXT")),
                "nullable": nullable_map.get(col["name"], True),
                "primary_key": pk_map.get(col["name"], False),
            }
            for col in live_columns
        ]
        excel_dbapi.write_table_metadata(raw_conn, table_name, columns)

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
    # Shares the same compiler state model as ExcelDialect.
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
