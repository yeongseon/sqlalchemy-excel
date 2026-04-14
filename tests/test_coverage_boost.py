from __future__ import annotations

# pyright: reportPrivateUsage=false, reportAny=false, reportExplicitAny=false, reportUnannotatedClassAttribute=false, reportUnusedCallResult=false
from types import SimpleNamespace
from typing import Any, cast

import pytest
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
    bindparam,
    create_engine,
    select,
)
from sqlalchemy.exc import CompileError, SAWarning
from sqlalchemy.sql.base import ColumnCollection
from sqlalchemy.sql.ddl import ExecutableDDLElement

from sqlalchemy_excel.compiler import ExcelCompiler
from sqlalchemy_excel.ddl import ExcelDDLCompiler
from sqlalchemy_excel.dialect import ExcelDialect, _after_create, _after_drop
from sqlalchemy_excel.dml import (
    OnConflictClause,
    OnConflictDoNothing,
    OnConflictDoUpdate,
    insert,
)


@pytest.fixture
def users_table() -> Table:
    metadata = MetaData()
    return Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("age", Integer),
    )


class AddColumnNoColumn(ExecutableDDLElement):
    __visit_name__ = "add_column"

    def __init__(self, table_name: str) -> None:
        self.table_name = table_name
        self.schema = None


class DropColumnFromColumnObj(ExecutableDDLElement):
    __visit_name__ = "drop_column"

    def __init__(self, table_name: str, column: Any) -> None:
        self.table_name = table_name
        self.column = column
        self.schema = None


class RenameColumnAltNames(ExecutableDDLElement):
    __visit_name__ = "rename_column"

    def __init__(self, table_name: str, old_name: str, new_name: str) -> None:
        self.table_name = table_name
        self.old_name = old_name
        self.new_name = new_name
        self.schema = None


def test_compiler_is_true_onclause_none_path() -> None:
    assert ExcelCompiler._is_true_onclause(None) is False


def test_compiler_validate_join_tree_requires_on_clause() -> None:
    bad_join = SimpleNamespace(
        left=object(), right=object(), onclause=None, full=False, isouter=False
    )
    with pytest.raises(CompileError, match="requires an ON clause"):
        ExcelCompiler._validate_join_tree(cast("Any", bad_join))


def test_compiler_validate_join_tree_recurses_into_right_join() -> None:
    metadata = MetaData()
    users = Table("users", metadata, Column("id", Integer, primary_key=True))
    orders = Table("orders", metadata, Column("user_id", Integer))
    items = Table("items", metadata, Column("order_user_id", Integer))

    right_join = orders.join(items, orders.c.user_id == items.c.order_user_id)
    outer_join = users.join(right_join, users.c.id == users.c.id)
    with pytest.raises(CompileError, match="different join sources"):
        ExcelCompiler._validate_join_tree(cast("Any", outer_join))


def test_compiler_count_empty_args_rejected(tmp_xlsx: str, users_table: Table) -> None:
    from sqlalchemy.sql.functions import Function

    engine = create_engine(f"excel:///{tmp_xlsx}")
    stmt = select(cast("Any", Function("sum"))).select_from(users_table)
    with pytest.raises(CompileError, match="expression arguments"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()


def test_compiler_visit_subquery_rejects_when_join_context(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users = Table("users", metadata, Column("id", Integer, primary_key=True))
    orders = Table("orders", metadata, Column("user_id", Integer))
    sub = select(orders.c.user_id).subquery()

    compiler_inst = cast("Any", select(users).compile(dialect=engine.dialect))
    compiler_inst._in_in_clause = True
    compiler_inst._has_join = True
    with pytest.raises(CompileError, match="subqueries with JOIN"):
        compiler_inst.visit_subquery(sub)
    engine.dispose()


def test_compiler_quote_aware_helpers_cover_escaped_quotes() -> None:
    sql = "('a''b')"
    depth = ExcelCompiler._update_depth_quote_aware(sql, 0, len(sql), 0)
    assert depth == 0
    assert (
        ExcelCompiler._has_top_level_compound_op("SELECT 'UNION''ALL' FROM t") is False
    )


def test_compiler_strip_compound_parens_with_quoted_segments() -> None:
    sql = "'x''y' (SELECT 'a''b' AS s FROM t1 ORDER BY s) UNION SELECT id FROM t2"
    out = ExcelCompiler._strip_compound_branch_parens(sql)
    assert "ORDER BY s" in out
    assert "'x''y'" in out


def test_compiler_strip_order_and_limit_helpers_with_quoted_and_embedded_keywords() -> (
    None
):
    stripped = ExcelCompiler._strip_top_level_order_by(
        "SELECT id FROM t ORDER BY id, 'LIMIT' LIMITED 5 OFFSET 2"
    )
    assert stripped.endswith("OFFSET 2")

    assert (
        ExcelCompiler._has_top_level_limit_offset(
            "SELECT 'LIMIT' AS x FROM t WHERE name = 'OFFSET'"
        )
        is False
    )


def test_compiler_on_conflict_do_nothing_without_target_text(
    tmp_xlsx: str, users_table: Table
) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    compiler_inst = cast("Any", select(users_table).compile(dialect=engine.dialect))
    clause = OnConflictDoNothing(index_elements=None)
    assert (
        compiler_inst.visit_on_conflict_do_nothing(clause) == "ON CONFLICT DO NOTHING"
    )
    engine.dispose()


def test_compiler_upsert_bindparam_null_type_and_extra_key_warns(
    tmp_xlsx: str,
    users_table: Table,
) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    stmt = (
        insert(users_table)
        .values(id=1, name="a", age=1)
        .on_conflict_do_update(
            index_elements=["id"],
            set_={"age": bindparam("new_age"), "extra_col": 10},
        )
    )
    with pytest.warns(SAWarning, match="Additional column names"):
        sql = str(stmt.compile(dialect=engine.dialect))
    assert "extra_col" in sql
    engine.dispose()


def test_ddl_compiler_error_paths_and_alternative_name_sources(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    compiler = ExcelDDLCompiler(engine.dialect, cast("Any", None))

    with pytest.raises(CompileError, match="does not support schemas"):
        compiler._format_alter_table_name(
            SimpleNamespace(schema="s", table_name="users")
        )
    with pytest.raises(CompileError, match="requires a table name"):
        compiler._format_alter_table_name(
            SimpleNamespace(schema=None, table_name=None, table=None)
        )
    with pytest.raises(CompileError, match="requires a column name"):
        compiler._format_alter_column_name(None)
    with pytest.raises(CompileError, match="requires a column"):
        AddColumnNoColumn("users").compile(dialect=engine.dialect)

    dropped = str(
        DropColumnFromColumnObj("users", SimpleNamespace(name="age")).compile(
            dialect=engine.dialect
        )
    )
    assert dropped == "ALTER TABLE users DROP COLUMN age"

    renamed = str(
        RenameColumnAltNames("users", "old_col", "new_col").compile(
            dialect=engine.dialect
        )
    )
    assert renamed == "ALTER TABLE users RENAME COLUMN old_col TO new_col"
    engine.dispose()


def test_dml_clause_none_targets_and_invalid_set_inputs() -> None:
    clause = OnConflictClause(index_elements=None)
    assert clause.inferred_target_elements is None
    assert clause.inferred_target_whereclause is None

    with pytest.raises(ValueError, match="must not be empty"):
        OnConflictDoUpdate(index_elements=["id"], set_=ColumnCollection())


def test_dialect_event_hooks_return_early_for_non_excel(users_table: Table) -> None:
    fake_conn = SimpleNamespace(dialect=SimpleNamespace(name="sqlite"))
    _after_create(users_table, fake_conn)
    _after_drop(users_table, fake_conn)


def test_dialect_sync_alter_table_short_statement_returns_early() -> None:
    dialect = ExcelDialect()
    cursor = SimpleNamespace(connection=object())
    dialect._sync_alter_table_metadata(cursor, "ALTER TABLE users")


def test_dialect_sync_alter_table_add_float_maps_to_real(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    writes: list[list[dict[str, Any]]] = []

    def read_table_metadata(_conn: Any, _table: str) -> list[dict[str, Any]]:
        return [
            {
                "name": "id",
                "type_name": "INTEGER",
                "nullable": False,
                "primary_key": True,
            }
        ]

    def get_columns(_conn: Any, _table: str) -> list[dict[str, Any]]:
        return [{"name": "id", "type": "INTEGER"}, {"name": "amount", "type": "FLOAT"}]

    def write_table_metadata(
        _conn: Any, _table: str, columns: list[dict[str, Any]]
    ) -> None:
        writes.append(columns)

    import excel_dbapi

    monkeypatch.setattr(excel_dbapi, "read_table_metadata", read_table_metadata)
    monkeypatch.setattr(excel_dbapi, "get_columns", get_columns)
    monkeypatch.setattr(excel_dbapi, "write_table_metadata", write_table_metadata)

    dialect = ExcelDialect()
    cursor = SimpleNamespace(connection=object())
    dialect._sync_alter_table_metadata(
        cursor, "ALTER TABLE users ADD COLUMN amount FLOAT"
    )

    amount = next(col for col in writes[-1] if col["name"] == "amount")
    assert amount["type_name"] == "REAL"


def test_dialect_sync_alter_table_rename_unknown_column_keeps_maps_stable(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    writes: list[list[dict[str, Any]]] = []

    def read_table_metadata(_conn: Any, _table: str) -> list[dict[str, Any]]:
        return []

    def get_columns(_conn: Any, _table: str) -> list[dict[str, Any]]:
        return [{"name": "new_name", "type": "TEXT"}]

    def write_table_metadata(
        _conn: Any, _table: str, columns: list[dict[str, Any]]
    ) -> None:
        writes.append(columns)

    import excel_dbapi

    monkeypatch.setattr(excel_dbapi, "read_table_metadata", read_table_metadata)
    monkeypatch.setattr(excel_dbapi, "get_columns", get_columns)
    monkeypatch.setattr(excel_dbapi, "write_table_metadata", write_table_metadata)

    dialect = ExcelDialect()
    cursor = SimpleNamespace(connection=object())
    dialect._sync_alter_table_metadata(
        cursor,
        "ALTER TABLE users RENAME COLUMN old_name TO new_name",
    )

    assert writes[-1] == [
        {
            "name": "new_name",
            "type_name": "TEXT",
            "nullable": True,
            "primary_key": False,
        }
    ]
