from __future__ import annotations

from typing import Any

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
    assert (
        compiled
        == "ALTER TABLE users ADD COLUMN external_id INTEGER NOT NULL PRIMARY KEY"
    )

    with engine.begin() as conn:
        conn.execute(ddl)

    inspector = inspect(engine)
    columns = inspector.get_columns("users")
    external_id = next(col for col in columns if col["name"] == "external_id")
    assert external_id["nullable"] is False
    assert "external_id" in inspector.get_pk_constraint("users")["constrained_columns"]


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
