"""Tests for ExcelDDLCompiler — CREATE TABLE and DROP TABLE."""

from __future__ import annotations

import pytest
from sqlalchemy import (
    CheckConstraint,
    Column,
    ForeignKey,
    Integer,
    MetaData,
    String,
    Table,
    UniqueConstraint,
    exc,
    inspect,
)
from sqlalchemy.schema import CreateTable, DropTable


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


class TestDDLCompilation:
    """Test DDL statement compilation."""

    def test_create_table_sql(self, engine, users_table):
        create = CreateTable(users_table)
        compiled = create.compile(dialect=engine.dialect)
        sql = str(compiled).strip()
        # Should produce: CREATE TABLE users (id INTEGER, name TEXT, age INTEGER)
        assert sql.startswith("CREATE TABLE")
        assert "users" in sql
        assert "id" in sql
        assert "INTEGER" in sql
        assert "TEXT" in sql
        assert "PRIMARY KEY" in sql

    def test_create_table_sql_includes_not_null(self, engine, metadata):
        accounts = Table(
            "accounts",
            metadata,
            Column("id", Integer, primary_key=True),
            Column("email", String, nullable=False),
        )
        sql = str(CreateTable(accounts).compile(dialect=engine.dialect)).strip()
        assert "email TEXT NOT NULL" in sql

    def test_create_table_sql_emits_table_level_primary_key_for_composite(
        self, engine, metadata
    ):
        memberships = Table(
            "memberships",
            metadata,
            Column("user_id", Integer, primary_key=True),
            Column("group_id", Integer, primary_key=True),
        )
        sql = str(CreateTable(memberships).compile(dialect=engine.dialect)).strip()
        assert "PRIMARY KEY (user_id, group_id)" in sql

    def test_create_table_warns_for_unsupported_unique_and_check(
        self, engine, metadata
    ):
        users = Table(
            "users",
            metadata,
            Column("id", Integer, primary_key=True),
            Column("name", String, unique=True),
            Column("age", Integer),
            UniqueConstraint("age", name="uq_users_age"),
            CheckConstraint("age > 0", name="ck_users_age_positive"),
        )

        with pytest.warns(UserWarning) as captured:
            sql = str(CreateTable(users).compile(dialect=engine.dialect)).strip()

        messages = [str(item.message) for item in captured]
        assert any("UNIQUE" in message for message in messages)
        assert any("CHECK" in message for message in messages)
        assert "UNIQUE" not in sql
        assert "CHECK" not in sql

    def test_create_table_warns_for_unsupported_foreign_key(self, engine, metadata):
        Table(
            "parents",
            metadata,
            Column("id", Integer, primary_key=True),
        )
        children = Table(
            "children",
            metadata,
            Column("id", Integer, primary_key=True),
            Column("parent_id", Integer, ForeignKey("parents.id")),
        )

        with pytest.warns(UserWarning, match="FOREIGN KEY"):
            str(CreateTable(children).compile(dialect=engine.dialect)).strip()

    def test_schema_qualified_table_compile_raises(self, engine, metadata):
        users = Table(
            "users",
            metadata,
            Column("id", Integer, primary_key=True),
            schema="myschema",
        )

        with pytest.raises(exc.CompileError, match="does not support schemas"):
            str(CreateTable(users).compile(dialect=engine.dialect))
        with pytest.raises(exc.CompileError, match="does not support schemas"):
            str(DropTable(users).compile(dialect=engine.dialect))

    def test_create_table_deduplicates_unique_warning_for_same_column(
        self, engine, metadata
    ):
        users = Table(
            "users",
            metadata,
            Column("id", Integer, primary_key=True),
            Column("email", String, unique=True),
            UniqueConstraint("email", name="uq_users_email"),
        )

        with pytest.warns(UserWarning) as captured:
            str(CreateTable(users).compile(dialect=engine.dialect)).strip()

        unique_messages = [
            str(item.message) for item in captured if "UNIQUE" in str(item.message)
        ]
        assert len(unique_messages) == 1

    def test_drop_table_sql(self, engine, users_table):
        drop = DropTable(users_table)
        compiled = drop.compile(dialect=engine.dialect)
        sql = str(compiled).strip()
        assert sql == "DROP TABLE users"


def test_split_sql_list_quote_aware_handles_single_quoted_comma() -> None:
    """Verify _split_sql_list_quote_aware keeps commas inside single quotes."""
    from sqlalchemy_excel.dialect import _split_sql_list_quote_aware

    sql = "id INTEGER PRIMARY KEY, note TEXT DEFAULT 'a,b' NOT NULL"
    parts = _split_sql_list_quote_aware(sql)
    assert len(parts) == 2
    assert parts[0].strip() == "id INTEGER PRIMARY KEY"
    assert "DEFAULT 'a,b'" in parts[1]


class TestDDLExecution:
    """Test DDL statement execution (integration)."""

    def test_create_table_creates_worksheet(self, engine, metadata, users_table):
        metadata.create_all(engine)
        insp = inspect(engine)
        tables = insp.get_table_names()
        assert "users" in tables

    def test_create_table_writes_metadata(self, engine, metadata, users_table):
        metadata.create_all(engine)
        with engine.connect() as conn:
            raw = conn.connection.dbapi_connection
            import excel_dbapi

            meta = excel_dbapi.read_table_metadata(raw, "users")
            assert meta is not None
            names = [c["name"] for c in meta]
            assert "id" in names
            assert "name" in names
            assert "age" in names

    def test_drop_table_removes_worksheet(self, engine, metadata, users_table):
        metadata.create_all(engine)
        metadata.drop_all(engine)
        insp = inspect(engine)
        tables = insp.get_table_names()
        assert "users" not in tables
