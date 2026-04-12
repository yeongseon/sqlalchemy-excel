from __future__ import annotations

from typing import Any, cast

import pytest
from sqlalchemy import (
    Column,
    Index,
    Integer,
    MetaData,
    Sequence,
    String,
    Table,
    UniqueConstraint,
    create_engine,
    distinct,
    exc,
    func,
    insert,
    select,
)
from sqlalchemy.engine import make_url
from sqlalchemy.schema import (
    AddConstraint,
    CreateIndex,
    CreateSequence,
    DropConstraint,
    DropIndex,
    DropSequence,
)
from sqlalchemy.sql import elements


def _build_tables(metadata: MetaData) -> tuple[Table, Table]:
    users = Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
    )
    return users, orders


def test_compiler_join_guard(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    with pytest.raises(exc.CompileError, match="JOIN"):
        select(users).join(orders, users.c.id == orders.c.user_id).compile(
            dialect=engine.dialect
        )
    engine.dispose()


def test_compiler_cte_subquery_returning_for_update_guards(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    with pytest.raises(exc.CompileError, match="CTE"):
        cte = select(users.c.id).cte("user_ids")
        select(cte.c.id).compile(dialect=engine.dialect)
    with pytest.raises(exc.CompileError, match="subqueries"):
        subq = select(users.c.id).subquery()
        select(subq.c.id).compile(dialect=engine.dialect)
    with pytest.raises(exc.CompileError, match="RETURNING"):
        insert(users).values(id=1, name="Alice").returning(users.c.id).compile(
            dialect=engine.dialect
        )
    with pytest.raises(exc.CompileError, match="FOR UPDATE"):
        select(users).with_for_update().compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_count_distinct(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    stmt = select(func.count(distinct(users.c.name))).select_from(users)
    with pytest.raises(exc.CompileError, match="expression arguments"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_count_literal(tmp_xlsx: str) -> None:
    from sqlalchemy import text

    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    stmt = select(func.count(text("1"))).select_from(users)
    with pytest.raises(exc.CompileError, match="expression arguments"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_count_string_literal(tmp_xlsx: str) -> None:
    from sqlalchemy import literal

    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    stmt = select(func.count(literal("x"))).select_from(users)
    with pytest.raises(exc.CompileError, match="expression arguments"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_visit_label_and_group_by_paths(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)
    label = users.c.id.label("identifier")

    compiled_stmt = select(label).compile(dialect=engine.dialect)
    assert " AS " not in str(compiled_stmt)

    rendered = compiled_stmt.visit_label(label, render_label_as_label=label)
    assert rendered == "identifier"
    assert compiled_stmt.visit_label(label) == "id"

    truncated = users.c.id.label("very_long_identifier_name")
    truncated.name = elements._truncated_label("very_long_identifier_name")
    assert compiled_stmt.visit_label(truncated, render_label_as_label=truncated)

    plain_select = select(users.c.id)
    plain_compiled = plain_select.compile(dialect=engine.dialect)
    assert plain_compiled.group_by_clause(plain_select) == ""

    gb_select = select(users.c.id).group_by(users.c.id)
    gb_compiled = gb_select.compile(dialect=engine.dialect)
    assert "GROUP BY" in gb_compiled.group_by_clause(gb_select)

    preparer = engine.dialect.identifier_preparer
    assert preparer.quote_identifier("users") == "users"
    assert preparer.quote("users") == "users"

    engine.dispose()


def test_ddl_unsupported_construct_guards(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    with pytest.raises(exc.CompileError, match="CREATE INDEX"):
        CreateIndex(Index("ix_users_name", users.c.name)).compile(
            dialect=engine.dialect
        )
    with pytest.raises(exc.CompileError, match="DROP INDEX"):
        DropIndex(Index("ix_users_name_2", users.c.name)).compile(
            dialect=engine.dialect
        )

    unique = UniqueConstraint(users.c.name, name="uq_users_name")
    with pytest.raises(exc.CompileError, match="constraints"):
        AddConstraint(unique).compile(dialect=engine.dialect)
    with pytest.raises(exc.CompileError, match="constraints"):
        DropConstraint(unique).compile(dialect=engine.dialect)

    with pytest.raises(exc.CompileError, match="sequences"):
        CreateSequence(Sequence("seq_users")).compile(dialect=engine.dialect)
    with pytest.raises(exc.CompileError, match="sequences"):
        DropSequence(Sequence("seq_users_drop")).compile(dialect=engine.dialect)

    engine.dispose()


def test_dialect_autocommit_query_parsing_and_no_params_execution() -> None:
    engine = create_engine("excel:///test.xlsx")
    dialect = cast("Any", engine.dialect)

    _, kwargs_true = dialect.create_connect_args(
        make_url("excel:///test.xlsx?engine=openpyxl&autocommit=true")
    )
    assert kwargs_true["engine"] == "openpyxl"
    assert kwargs_true["autocommit"] is True

    _, kwargs_tuple = dialect.create_connect_args(
        make_url("excel:///test.xlsx?autocommit=true&autocommit=false")
    )
    assert kwargs_tuple["autocommit"] is True

    class CursorStub:
        def __init__(self) -> None:
            self.calls: list[tuple[str, object]] = []

        def execute(self, statement: str, parameters: object) -> None:
            self.calls.append((statement, parameters))

    cursor = CursorStub()
    dialect.do_execute_no_params(cursor, "SELECT   1\nFROM   users")
    assert cursor.calls == [("SELECT 1 FROM users", None)]

    assert dialect._check_unicode_returns(connection=None) is True
    assert dialect._check_unicode_description(connection=None) is True

    engine.dispose()
