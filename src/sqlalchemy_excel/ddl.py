"""DDL compiler for the Excel dialect.

CREATE TABLE → creates a worksheet and writes column metadata.
DROP TABLE   → deletes the worksheet and removes metadata.
ALTER TABLE  → supports ADD/DROP/RENAME COLUMN.
"""

from __future__ import annotations

from typing import Any

from sqlalchemy import exc
from sqlalchemy.sql import compiler


class ExcelDDLCompiler(compiler.DDLCompiler):
    """Compiles DDL statements for excel-dbapi."""

    def _format_alter_table_name(self, operation: Any) -> str:
        schema = getattr(operation, "schema", None)
        if schema is not None:
            raise exc.CompileError("Excel dialect does not support schemas")

        table_name = getattr(operation, "table_name", None)
        if table_name is None:
            table = getattr(operation, "table", None)
            table_name = getattr(table, "name", None)

        if not isinstance(table_name, str) or not table_name:
            raise exc.CompileError("ALTER TABLE operation requires a table name")

        return self.preparer.quote(table_name)

    def _format_alter_column_name(self, column_name: Any) -> str:
        if not isinstance(column_name, str) or not column_name:
            raise exc.CompileError("ALTER TABLE operation requires a column name")
        return self.preparer.quote(column_name)

    def visit_create_table(self, create: Any, **kw: Any) -> str:
        """Compile CREATE TABLE into SQL that excel-dbapi's parser accepts.

        Format: CREATE TABLE name (col1 TYPE, col2 TYPE, ...)
        """
        table = create.element
        table_name = self.preparer.format_table(table)

        columns = []
        for col in table.columns:
            col_name = self.preparer.format_column(col)
            col_type = self.dialect.type_compiler.process(col.type)
            columns.append(f"{col_name} {col_type}")

        return f"CREATE TABLE {table_name} ({', '.join(columns)})"

    def visit_drop_table(self, drop: Any, **kw: Any) -> str:
        """Compile DROP TABLE."""
        table = drop.element
        table_name = self.preparer.format_table(table)
        return f"DROP TABLE {table_name}"

    def visit_add_column(self, create: Any, **kw: Any) -> str:
        table_name = self._format_alter_table_name(create)
        column = getattr(create, "column", None)
        if column is None:
            raise exc.CompileError("ALTER TABLE ADD COLUMN requires a column")

        col_name = self.preparer.format_column(column)
        col_type = self.dialect.type_compiler.process(column.type)
        constraints: list[str] = []
        if column.nullable is False:
            constraints.append("NOT NULL")
        if column.primary_key is True:
            constraints.append("PRIMARY KEY")
        suffix = f" {' '.join(constraints)}" if constraints else ""
        return f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type}{suffix}"

    def visit_drop_column(self, drop: Any, **kw: Any) -> str:
        table_name = self._format_alter_table_name(drop)

        column_name = getattr(drop, "column_name", None)
        if column_name is None:
            column = getattr(drop, "column", None)
            column_name = getattr(column, "name", column)

        column_name = self._format_alter_column_name(column_name)
        return f"ALTER TABLE {table_name} DROP COLUMN {column_name}"

    def visit_rename_column(self, rename: Any, **kw: Any) -> str:
        table_name = self._format_alter_table_name(rename)

        old_column_name = getattr(rename, "column_name", None)
        if old_column_name is None:
            old_column_name = getattr(rename, "old_column_name", None)
        if old_column_name is None:
            old_column_name = getattr(rename, "old_name", None)

        new_column_name = getattr(rename, "new_column_name", None)
        if new_column_name is None:
            new_column_name = getattr(rename, "name", None)
        if new_column_name is None:
            new_column_name = getattr(rename, "new_name", None)

        old_name = self._format_alter_column_name(old_column_name)
        new_name = self._format_alter_column_name(new_column_name)
        return f"ALTER TABLE {table_name} RENAME COLUMN {old_name} TO {new_name}"

    def visit_create_index(
        self,
        create: Any,
        include_schema: Any = False,
        include_table_schema: Any = True,
        **kw: Any,
    ) -> str:
        raise exc.CompileError("Excel dialect does not support CREATE INDEX")

    def visit_drop_index(self, drop: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support DROP INDEX")

    def visit_add_constraint(self, create: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support constraints")

    def visit_drop_constraint(self, drop: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support constraints")

    def visit_create_sequence(
        self,
        create: Any,
        prefix: Any = None,
        **kw: Any,
    ) -> str:
        raise exc.CompileError("Excel dialect does not support sequences")

    def visit_drop_sequence(self, drop: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support sequences")
