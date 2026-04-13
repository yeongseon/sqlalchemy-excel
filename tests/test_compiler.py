"""Tests for ExcelCompiler — SQL compilation and rejection."""

from __future__ import annotations

import pytest
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
    create_engine,
    exc,
    select,
)


@pytest.fixture
def metadata():
    return MetaData()


@pytest.fixture
def users_table(metadata):
    return Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("age", Integer),
    )


@pytest.fixture
def orders_table(metadata):
    return Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
    )


@pytest.fixture
def excel_engine(tmp_xlsx):
    eng = create_engine(f"excel:///{tmp_xlsx}")
    yield eng
    eng.dispose()


class TestSelectCompilation:
    """Test SELECT statement compilation."""

    def test_select_all(self, excel_engine, users_table):
        stmt = select(users_table)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "SELECT" in sql
        assert "FROM" in sql
        assert "users" in sql

    def test_select_columns(self, excel_engine, users_table):
        stmt = select(users_table.c.id, users_table.c.name)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "id" in sql
        assert "name" in sql

    def test_select_where(self, excel_engine, users_table):
        stmt = select(users_table).where(users_table.c.id == 1)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "WHERE" in sql

    def test_select_order_by(self, excel_engine, users_table):
        stmt = select(users_table).order_by(users_table.c.name)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "ORDER BY" in sql

    def test_select_limit(self, excel_engine, users_table):
        stmt = select(users_table).limit(10)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "LIMIT" in sql

    def test_select_offset(self, excel_engine, users_table):
        stmt = select(users_table).limit(10).offset(5)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "LIMIT" in sql
        assert "OFFSET" in sql

    def test_select_offset_only(self, excel_engine, users_table):
        stmt = select(users_table).offset(5)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "OFFSET" in sql

    def test_select_distinct(self, excel_engine, users_table):
        stmt = select(users_table.c.name).distinct()
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "DISTINCT" in sql

    def test_select_distinct_with_limit_offset(self, excel_engine, users_table):
        stmt = select(users_table.c.name).distinct().limit(10).offset(5)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "DISTINCT" in sql
        assert "LIMIT" in sql
        assert "OFFSET" in sql

    def test_select_where_in_subquery(self, excel_engine, users_table, orders_table):
        stmt = select(users_table).where(
            users_table.c.id.in_(select(orders_table.c.user_id))
        )
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "SELECT" in sql
        assert "IN" in sql


class TestColumnAliasCompilation:
    """Test column alias (AS) SQL compilation."""

    def test_label_emits_as(self, excel_engine, users_table):
        stmt = select(users_table.c.name.label("n"))
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert " AS " in sql
        assert "n" in sql

    def test_aggregate_label(self, excel_engine, users_table):
        from sqlalchemy import func

        stmt = select(func.count(users_table.c.id).label("total")).select_from(users_table)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert " AS " in sql
        assert "total" in sql

    def test_join_label(self, excel_engine, users_table, orders_table):
        stmt = select(
            users_table.c.name.label("user_name"),
            orders_table.c.amount.label("order_amount"),
        ).join(orders_table, users_table.c.id == orders_table.c.user_id)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "AS user_name" in sql
        assert "AS order_amount" in sql

    def test_mixed_label_and_bare(self, excel_engine, users_table):
        stmt = select(users_table.c.name.label("n"), users_table.c.age)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "AS n" in sql
        assert "AS age" not in sql


class TestCompilationRejection:
    """Test that unsupported features raise CompileError."""

    def test_join_compiles(self, excel_engine, users_table, orders_table):
        stmt = select(users_table.c.id, orders_table.c.user_id).join(
            orders_table, users_table.c.id == orders_table.c.user_id
        )
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "JOIN" in sql
        assert "ON" in sql
        assert "users.id" in sql
        assert "orders.user_id" in sql
    def test_group_by_compiles(self, excel_engine, users_table):
        stmt = select(users_table.c.name).group_by(users_table.c.name)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "GROUP BY" in sql
        assert "name" in sql

    def test_having_compiles(self, excel_engine, users_table):
        from sqlalchemy import func

        stmt = (
            select(users_table.c.name, func.count(users_table.c.id))
            .group_by(users_table.c.name)
            .having(func.count(users_table.c.id) > 1)
        )
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "GROUP BY" in sql
        assert "HAVING" in sql

    def test_aggregate_count_compiles(self, excel_engine, users_table):
        from sqlalchemy import func

        stmt = select(func.count(users_table.c.id)).select_from(users_table)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "count" in sql.lower()

    def test_aggregate_sum_compiles(self, excel_engine, users_table):
        from sqlalchemy import func

        stmt = select(func.sum(users_table.c.age)).select_from(users_table)
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "sum" in sql.lower()

    def test_unsupported_aggregate_function_rejected(self, excel_engine, users_table):
        from sqlalchemy import func

        stmt = select(func.median(users_table.c.age)).select_from(users_table)
        with pytest.raises(exc.CompileError, match="function"):
            stmt.compile(dialect=excel_engine.dialect)


class TestMultiColumnOrderBy:
    """Test multi-column ORDER BY SQL compilation."""

    def test_multi_column_order_by(self, excel_engine, users_table):
        stmt = (
            select(users_table)
            .order_by(users_table.c.name.asc(), users_table.c.age.desc())
        )
        compiled = stmt.compile(dialect=excel_engine.dialect)
        sql = str(compiled)
        assert "ORDER BY" in sql
        assert "name ASC" in sql
        assert "age DESC" in sql
        # Verify column order: name comes before age
        name_pos = sql.index("name ASC")
        age_pos = sql.index("age DESC")
        assert name_pos < age_pos
