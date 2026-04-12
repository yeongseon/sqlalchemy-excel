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
    union_all,
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


def test_compiler_join_compiles(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = select(users.c.id, orders.c.user_id).join(
        orders, users.c.id == orders.c.user_id
    )
    compiled = stmt.compile(dialect=engine.dialect)
    sql = str(compiled)
    assert "JOIN" in sql
    assert "ON" in sql
    # Table-qualified column names in JOIN context
    assert "users.id" in sql
    assert "orders.user_id" in sql
    engine.dispose()


def test_compiler_cte_returning_for_update_guards(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    with pytest.raises(exc.CompileError, match="CTE"):
        cte = select(users.c.id).cte("user_ids")
        select(cte.c.id).compile(dialect=engine.dialect)
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


def test_compiler_rejects_union(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    stmt = union_all(select(users.c.id), select(users.c.id))
    with pytest.raises(exc.CompileError, match=r"UNION|INTERSECT|EXCEPT"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_window_over(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    stmt = select(func.count().over()).select_from(users)
    with pytest.raises(exc.CompileError, match="window functions"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_aggregate_filter(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    stmt = select(func.count(users.c.id).filter(users.c.id > 0)).select_from(users)
    with pytest.raises(exc.CompileError, match="FILTER"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_scalar_subquery(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    scalar_sub = select(func.max(orders.c.user_id)).scalar_subquery()
    stmt = select(users).where(users.c.id == scalar_sub)
    with pytest.raises(exc.CompileError, match="only supports subqueries in WHERE"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()


def test_compiler_rejects_from_subquery(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    sub = select(users.c.id).subquery()
    stmt = select(sub.c.id)
    with pytest.raises(exc.CompileError, match="only supports subqueries in WHERE"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()


def test_compiler_rejects_correlated_subquery(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    correlated = select(orders.c.user_id).where(orders.c.user_id == users.c.id)
    stmt = select(users).where(users.c.id.in_(correlated))
    with pytest.raises(exc.CompileError, match="correlated"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()


def test_compiler_rejects_not_in(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, _ = _build_tables(metadata)

    stmt = select(users).where(users.c.id.not_in([1, 2, 3]))
    with pytest.raises(exc.CompileError, match="NOT IN"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()


def test_compiler_rejects_not_in_subquery(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    sub = select(orders.c.user_id)
    stmt = select(users).where(users.c.id.not_in(sub))
    with pytest.raises(exc.CompileError, match="NOT IN"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()

def test_compiler_rejects_update_with_subquery(tmp_xlsx: str) -> None:
    """UPDATE with IN subquery should be rejected."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    sub = select(orders.c.user_id)
    stmt = users.update().where(users.c.id.in_(sub)).values(name="x")
    with pytest.raises(exc.CompileError, match="does not support subqueries in UPDATE/DELETE"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()


def test_compiler_rejects_delete_with_subquery(tmp_xlsx: str) -> None:
    """DELETE with IN subquery should be rejected."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    sub = select(orders.c.user_id)
    stmt = users.delete().where(users.c.id.in_(sub))
    with pytest.raises(exc.CompileError, match="does not support subqueries in UPDATE/DELETE"):
        stmt.compile(dialect=engine.dialect)
    engine.dispose()


def test_compiler_rejects_nested_subquery(tmp_xlsx: str) -> None:
    """Nested subquery inside IN should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)
    items = Table(
        "items",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("order_id", Integer),
    )

    inner = select(items.c.order_id).where(items.c.id > 0)
    outer = select(orders.c.user_id).where(orders.c.id.in_(inner))
    stmt = select(users).where(users.c.id.in_(outer))
    with pytest.raises(exc.CompileError, match="nested subqueries"):
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


def test_compiler_visit_subquery_direct_guards(tmp_xlsx: str) -> None:
    """Exercise visit_subquery directly to cover depth/update checks."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    sub = select(orders.c.user_id).subquery()
    compiler_inst = select(users).compile(dialect=engine.dialect)

    # Without _in_in_clause, subquery is rejected
    with pytest.raises(exc.CompileError, match="only supports subqueries in WHERE"):
        compiler_inst.visit_subquery(sub)

    # With _in_in_clause but depth > 0, nested subquery rejected
    compiler_inst._in_in_clause = True
    compiler_inst._subquery_depth = 1
    with pytest.raises(exc.CompileError, match="nested subqueries"):
        compiler_inst.visit_subquery(sub)

    # Reset depth, it should work
    compiler_inst._subquery_depth = 0
    result = compiler_inst.visit_subquery(sub)
    assert "user_id" in result
    assert "orders" in result

    engine.dispose()


def test_compiler_visit_subquery_rejects_in_update_context(tmp_xlsx: str) -> None:
    """visit_subquery rejects subqueries when statement is UPDATE/DELETE."""
    from sqlalchemy import update as sa_update

    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    sub = select(orders.c.user_id).subquery()
    # Compile an UPDATE statement
    update_stmt = sa_update(users).values(name="x")
    compiler_inst = update_stmt.compile(dialect=engine.dialect)
    compiler_inst._in_in_clause = True

    with pytest.raises(exc.CompileError, match="does not support subqueries in UPDATE/DELETE"):
        compiler_inst.visit_subquery(sub)

    engine.dispose()


def test_compiler_rejects_join_with_distinct(tmp_xlsx: str) -> None:
    """JOIN + DISTINCT should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = (
        select(users.c.id, orders.c.user_id)
        .join(orders, users.c.id == orders.c.user_id)
        .distinct()
    )
    with pytest.raises(exc.CompileError, match="DISTINCT with JOIN"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_join_with_aggregate(tmp_xlsx: str) -> None:
    """JOIN + aggregate functions should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = select(func.count(users.c.id)).join(
        orders, users.c.id == orders.c.user_id
    )
    with pytest.raises(exc.CompileError, match="aggregate functions with JOIN"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_join_with_group_by(tmp_xlsx: str) -> None:
    """JOIN + GROUP BY should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = (
        select(users.c.name)
        .join(orders, users.c.id == orders.c.user_id)
        .group_by(users.c.name)
    )
    with pytest.raises(exc.CompileError, match="GROUP BY with JOIN"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_join_with_having(tmp_xlsx: str) -> None:
    """JOIN + HAVING should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = (
        select(users.c.name)
        .join(orders, users.c.id == orders.c.user_id)
        .group_by(users.c.name)
        .having(func.count(users.c.id) > 1)
    )
    with pytest.raises(exc.CompileError, match="(GROUP BY|HAVING) with JOIN"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_join_with_subquery(tmp_xlsx: str) -> None:
    """JOIN + WHERE IN (subquery) should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)
    admins = Table(
        "admins",
        metadata,
        Column("id", Integer, primary_key=True),
    )

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, users.c.id == orders.c.user_id)
        .where(users.c.id.in_(select(admins.c.id)))
    )
    with pytest.raises(exc.CompileError, match="subqueries with JOIN"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_subquery_containing_join(tmp_xlsx: str) -> None:
    """Subquery that itself contains a JOIN should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)
    admins = Table(
        "admins",
        metadata,
        Column("id", Integer, primary_key=True),
    )

    # Subquery: SELECT users.id FROM users JOIN orders ON users.id = orders.user_id
    inner_join_subquery = (
        select(users.c.id)
        .join(orders, users.c.id == orders.c.user_id)
    )
    stmt = select(admins.c.id).where(admins.c.id.in_(inner_join_subquery))
    with pytest.raises(exc.CompileError, match="JOIN inside subqueries"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_chained_joins(tmp_xlsx: str) -> None:
    """Chained JOIN (more than one JOIN) should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)
    admins = Table(
        "admins",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
    )

    stmt = (
        select(users.c.name)
        .join(orders, users.c.id == orders.c.user_id)
        .join(admins, users.c.id == admins.c.user_id)
    )
    with pytest.raises(exc.CompileError, match="only one JOIN per query"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_full_outer_join(tmp_xlsx: str) -> None:
    """FULL OUTER JOIN should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, users.c.id == orders.c.user_id, full=True)
    )
    with pytest.raises(exc.CompileError, match="FULL OUTER JOIN"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_non_equality_on_clause(tmp_xlsx: str) -> None:
    """Non-equality ON clause (e.g. >) should be rejected at compile time."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, users.c.id > orders.c.user_id)
    )
    with pytest.raises(exc.CompileError, match="'=' comparisons"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_rejects_true_on_clause(tmp_xlsx: str) -> None:
    """true() ON clause (cross-join-like) should be rejected at compile time."""
    from sqlalchemy import true

    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users, orders = _build_tables(metadata)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, true())
    )
    with pytest.raises(exc.CompileError, match="equality comparisons"):
        stmt.compile(dialect=engine.dialect)

    engine.dispose()


def test_compiler_accepts_and_combined_equality_on(tmp_xlsx: str) -> None:
    """AND-combined equality ON clause should be accepted."""
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    users = Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("dept_id", Integer),
    )
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("dept_id", Integer),
    )

    stmt = (
        select(users.c.id, orders.c.id)
        .join(
            orders,
            (users.c.id == orders.c.user_id) & (users.c.dept_id == orders.c.dept_id),
        )
    )
    # Should not raise
    compiled = stmt.compile(dialect=engine.dialect)
    sql_text = str(compiled)
    assert "JOIN" in sql_text
    assert "AND" in sql_text

    engine.dispose()
