from __future__ import annotations

from typing import Any

import pytest
from sqlalchemy import Column, Integer, MetaData, String, Table, inspect
from sqlalchemy.sql.ddl import ExecutableDDLElement


class AddColumn(ExecutableDDLElement):
    __visit_name__ = "add_column"

    def __init__(self, table_name: str, column: Column[Any]) -> None:
        self.table_name = table_name
        self.column = column
        self.schema = None


class DropColumn(ExecutableDDLElement):
    __visit_name__ = "drop_column"

    def __init__(self, table_name: str, column_name: str) -> None:
        self.table_name = table_name
        self.column_name = column_name
        self.schema = None


class RenameColumn(ExecutableDDLElement):
    __visit_name__ = "rename_column"

    def __init__(
        self,
        table_name: str,
        old_column_name: str,
        new_column_name: str,
    ) -> None:
        self.table_name = table_name
        self.column_name = old_column_name
        self.new_column_name = new_column_name
        self.schema = None


def _create_users_table(metadata: MetaData) -> Table:
    return Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("age", Integer),
    )


def test_alter_table_add_column(engine) -> None:
    metadata = MetaData()
    _create_users_table(metadata)
    metadata.create_all(engine)

    ddl = AddColumn("users", Column("email", String))
    compiled = str(ddl.compile(dialect=engine.dialect)).strip()
    assert compiled == "ALTER TABLE users ADD COLUMN email TEXT"

    with engine.begin() as conn:
        conn.execute(ddl)

    columns = [col["name"] for col in inspect(engine).get_columns("users")]
    assert columns == ["id", "name", "age", "email"]


def test_alter_table_add_column_with_constraints_reflection_round_trip(engine) -> None:
    metadata = MetaData()
    _create_users_table(metadata)
    metadata.create_all(engine)

    ddl = AddColumn(
        "users",
        Column("external_id", Integer, nullable=False, primary_key=True),
    )
    compiled = str(ddl.compile(dialect=engine.dialect)).strip()
    assert compiled == "ALTER TABLE users ADD COLUMN external_id INTEGER PRIMARY KEY"

    with engine.begin() as conn:
        conn.execute(ddl)

    inspector = inspect(engine)
    columns = inspector.get_columns("users")
    external_id = next(col for col in columns if col["name"] == "external_id")
    assert external_id["nullable"] is False
    assert "external_id" in inspector.get_pk_constraint("users")["constrained_columns"]


def test_alter_table_add_column_warns_for_unsupported_unique(engine) -> None:
    ddl = AddColumn("users", Column("email", String, unique=True))

    with pytest.warns(UserWarning, match="UNIQUE"):
        compiled = str(ddl.compile(dialect=engine.dialect)).strip()

    assert compiled == "ALTER TABLE users ADD COLUMN email TEXT"


def test_alter_table_drop_column(engine) -> None:
    metadata = MetaData()
    _create_users_table(metadata)
    metadata.create_all(engine)

    ddl = DropColumn("users", "age")
    compiled = str(ddl.compile(dialect=engine.dialect)).strip()
    assert compiled == "ALTER TABLE users DROP COLUMN age"

    with engine.begin() as conn:
        conn.execute(ddl)

    columns = [col["name"] for col in inspect(engine).get_columns("users")]
    assert columns == ["id", "name"]


def test_alter_table_rename_column(engine) -> None:
    metadata = MetaData()
    _create_users_table(metadata)
    metadata.create_all(engine)

    ddl = RenameColumn("users", "name", "full_name")
    compiled = str(ddl.compile(dialect=engine.dialect)).strip()
    assert compiled == "ALTER TABLE users RENAME COLUMN name TO full_name"

    with engine.begin() as conn:
        conn.execute(ddl)

    columns = [col["name"] for col in inspect(engine).get_columns("users")]
    assert columns == ["id", "full_name", "age"]


def test_alter_table_preserves_existing_pk_and_nullability(engine) -> None:
    metadata = MetaData()
    Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("age", Integer, nullable=False),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(AddColumn("users", Column("email", String)))

    inspector = inspect(engine)
    columns_after_add = {col["name"]: col for col in inspector.get_columns("users")}
    assert columns_after_add["age"]["nullable"] is False
    assert inspector.get_pk_constraint("users")["constrained_columns"] == ["id"]

    with engine.begin() as conn:
        conn.execute(RenameColumn("users", "age", "years"))

    inspector = inspect(engine)
    columns_after_rename = {col["name"]: col for col in inspector.get_columns("users")}
    assert columns_after_rename["years"]["nullable"] is False
    assert inspector.get_pk_constraint("users")["constrained_columns"] == ["id"]

    with engine.begin() as conn:
        conn.execute(DropColumn("users", "name"))

    inspector = inspect(engine)
    final_columns = {col["name"]: col for col in inspector.get_columns("users")}
    assert "name" not in final_columns
    assert final_columns["years"]["nullable"] is False
    assert inspector.get_pk_constraint("users")["constrained_columns"] == ["id"]
