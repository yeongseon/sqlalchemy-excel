"""Tests for ExcelDDLCompiler — CREATE TABLE and DROP TABLE."""

from __future__ import annotations

import pytest
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
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

    def test_drop_table_sql(self, engine, users_table):
        drop = DropTable(users_table)
        compiled = drop.compile(dialect=engine.dialect)
        sql = str(compiled).strip()
        assert sql == "DROP TABLE users"


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
