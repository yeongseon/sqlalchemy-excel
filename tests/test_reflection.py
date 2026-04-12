"""Tests for reflection — Inspector integration."""

from __future__ import annotations

import pytest
from sqlalchemy import (
    Column,
    Float,
    Integer,
    MetaData,
    String,
    Table,
    inspect,
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
        Column("score", Float),
    )


@pytest.fixture
def populated_engine(engine, metadata, users_table):
    """Engine with a 'users' table already created."""
    metadata.create_all(engine)
    return engine


class TestGetTableNames:
    """Test get_table_names."""

    def test_no_tables_initially(self, engine):
        insp = inspect(engine)
        # A fresh xlsx might have a default sheet
        tables = insp.get_table_names()
        assert isinstance(tables, list)

    def test_created_table_appears(self, populated_engine):
        insp = inspect(populated_engine)
        tables = insp.get_table_names()
        assert "users" in tables

    def test_metadata_sheet_excluded(self, populated_engine):
        insp = inspect(populated_engine)
        tables = insp.get_table_names()
        assert "__excel_meta__" not in tables


class TestHasTable:
    """Test has_table."""

    def test_existing_table(self, populated_engine):
        insp = inspect(populated_engine)
        assert insp.has_table("users")

    def test_nonexistent_table(self, populated_engine):
        insp = inspect(populated_engine)
        assert not insp.has_table("nonexistent")


class TestGetColumns:
    """Test get_columns — from metadata sheet."""

    def test_column_names(self, populated_engine):
        insp = inspect(populated_engine)
        columns = insp.get_columns("users")
        names = [c["name"] for c in columns]
        assert "id" in names
        assert "name" in names
        assert "score" in names

    def test_column_count(self, populated_engine):
        insp = inspect(populated_engine)
        columns = insp.get_columns("users")
        assert len(columns) == 3

    def test_column_has_type(self, populated_engine):
        insp = inspect(populated_engine)
        columns = insp.get_columns("users")
        for col in columns:
            assert "type" in col
            assert col["type"] is not None


class TestGetPKConstraint:
    """Test get_pk_constraint — from metadata sheet."""

    def test_pk_columns(self, populated_engine):
        insp = inspect(populated_engine)
        pk = insp.get_pk_constraint("users")
        assert "id" in pk["constrained_columns"]

    def test_pk_excludes_non_pk(self, populated_engine):
        insp = inspect(populated_engine)
        pk = insp.get_pk_constraint("users")
        assert "name" not in pk["constrained_columns"]
        assert "score" not in pk["constrained_columns"]


class TestEmptyResults:
    """Test reflection methods that always return empty."""

    def test_no_foreign_keys(self, populated_engine):
        insp = inspect(populated_engine)
        fks = insp.get_foreign_keys("users")
        assert fks == []

    def test_no_indexes(self, populated_engine):
        insp = inspect(populated_engine)
        indexes = insp.get_indexes("users")
        assert indexes == []

    def test_no_views(self, populated_engine):
        insp = inspect(populated_engine)
        views = insp.get_view_names()
        assert views == []
