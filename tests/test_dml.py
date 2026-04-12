"""Tests for DML operations — INSERT, SELECT, UPDATE, DELETE round-trip."""

from __future__ import annotations

import pytest
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
    delete,
    insert,
    select,
    update,
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
def populated_engine(engine, metadata, users_table):
    """Engine with 'users' table created."""
    metadata.create_all(engine)
    return engine


class TestInsert:
    """Test INSERT operations."""

    def test_insert_single_row(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            conn.execute(insert(users_table).values(id=1, name="Alice", age=30))
            conn.commit()

            result = conn.execute(select(users_table))
            rows = result.fetchall()
            assert len(rows) == 1
            assert rows[0][0] == 1
            assert rows[0][1] == "Alice"
            assert rows[0][2] == 30

    def test_insert_multiple_rows(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            conn.execute(insert(users_table).values(id=1, name="Alice", age=30))
            conn.execute(insert(users_table).values(id=2, name="Bob", age=25))
            conn.commit()

            result = conn.execute(select(users_table))
            rows = result.fetchall()
            assert len(rows) == 2


class TestSelect:
    """Test SELECT operations."""

    def _seed(self, conn, users_table):
        conn.execute(insert(users_table).values(id=1, name="Alice", age=30))
        conn.execute(insert(users_table).values(id=2, name="Bob", age=25))
        conn.execute(insert(users_table).values(id=3, name="Charlie", age=35))
        conn.commit()

    def test_select_all(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            self._seed(conn, users_table)
            result = conn.execute(select(users_table))
            rows = result.fetchall()
            assert len(rows) == 3

    def test_select_where_eq(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            self._seed(conn, users_table)
            stmt = select(users_table).where(users_table.c.name == "Alice")
            result = conn.execute(stmt)
            rows = result.fetchall()
            assert len(rows) == 1
            assert rows[0][1] == "Alice"

    def test_select_where_gt(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            self._seed(conn, users_table)
            stmt = select(users_table).where(users_table.c.age > 28)
            result = conn.execute(stmt)
            rows = result.fetchall()
            assert len(rows) == 2

    def test_select_order_by_asc(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            self._seed(conn, users_table)
            stmt = select(users_table).order_by(users_table.c.age.asc())
            result = conn.execute(stmt)
            rows = result.fetchall()
            ages = [row[2] for row in rows]
            assert ages == [25, 30, 35]

    def test_select_order_by_desc(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            self._seed(conn, users_table)
            stmt = select(users_table).order_by(users_table.c.age.desc())
            result = conn.execute(stmt)
            rows = result.fetchall()
            ages = [row[2] for row in rows]
            assert ages == [35, 30, 25]

    def test_select_limit(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            self._seed(conn, users_table)
            stmt = select(users_table).limit(2)
            result = conn.execute(stmt)
            rows = result.fetchall()
            assert len(rows) == 2

    def test_select_specific_columns(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            self._seed(conn, users_table)
            stmt = select(users_table.c.name, users_table.c.age)
            result = conn.execute(stmt)
            rows = result.fetchall()
            assert len(rows) == 3
            # Each row should have 2 columns
            assert len(rows[0]) == 2


class TestUpdate:
    """Test UPDATE operations."""

    def test_update_with_where(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            conn.execute(insert(users_table).values(id=1, name="Alice", age=30))
            conn.commit()

            conn.execute(
                update(users_table).where(users_table.c.id == 1).values(age=31)
            )
            conn.commit()

            result = conn.execute(select(users_table).where(users_table.c.id == 1))
            rows = result.fetchall()
            assert rows[0][2] == 31


class TestDelete:
    """Test DELETE operations."""

    def test_delete_with_where(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            conn.execute(insert(users_table).values(id=1, name="Alice", age=30))
            conn.execute(insert(users_table).values(id=2, name="Bob", age=25))
            conn.commit()

            conn.execute(delete(users_table).where(users_table.c.id == 1))
            conn.commit()

            result = conn.execute(select(users_table))
            rows = result.fetchall()
            assert len(rows) == 1
            assert rows[0][1] == "Bob"

    def test_delete_all(self, populated_engine, users_table):
        with populated_engine.connect() as conn:
            conn.execute(insert(users_table).values(id=1, name="Alice", age=30))
            conn.execute(insert(users_table).values(id=2, name="Bob", age=25))
            conn.commit()

            conn.execute(delete(users_table))
            conn.commit()

            result = conn.execute(select(users_table))
            rows = result.fetchall()
            assert len(rows) == 0
