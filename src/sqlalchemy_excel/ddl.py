"""DDL compiler for the Excel dialect.

CREATE TABLE → creates a worksheet and writes column metadata.
DROP TABLE   → deletes the worksheet and removes metadata.
ALTER TABLE  → supports ADD/DROP/RENAME COLUMN.
"""

from __future__ import annotations

import warnings
from typing import Any

from sqlalchemy import exc
from sqlalchemy.schema import CheckConstraint, UniqueConstraint
from sqlalchemy.sql import compiler


class ExcelDDLCompiler(compiler.DDLCompiler):
    """Compiles DDL statements for excel-dbapi."""

    @staticmethod
    def _warn_unsupported_constraints(
        column: Any,
        context_msg: str,
        seen_unique: set[tuple[str, ...]],
        seen_check: set[str],
    ) -> None:
        if column.unique is True:
            key = (column.name,)
            if key not in seen_unique:
                warnings.warn(
                    f"Excel dialect does not enforce UNIQUE constraints ({context_msg})",
                    stacklevel=2,
                )
                seen_unique.add(key)

        for constraint in column.constraints:
            if isinstance(constraint, UniqueConstraint):
                unique_key = tuple(col.name for col in constraint.columns) or (
                    column.name,
                )
                if unique_key not in seen_unique:
                    warnings.warn(
                        f"Excel dialect does not enforce UNIQUE constraints ({context_msg})",
                        stacklevel=2,
                    )
                    seen_unique.add(unique_key)
            if isinstance(constraint, CheckConstraint):
                check_key = str(constraint.sqltext)
                if check_key not in seen_check:
                    warnings.warn(
                        f"Excel dialect does not enforce CHECK constraints ({context_msg})",
                        stacklevel=2,
                    )
                    seen_check.add(check_key)

    @staticmethod
    def _warn_unsupported_generated_columns(column: Any) -> None:
        if getattr(column, "computed", None) is not None:
            warnings.warn(
                f"Column '{column.name}': Computed columns are not supported by excel dialect; the expression will be ignored",
                stacklevel=2,
            )
        if getattr(column, "identity", None) is not None:
            warnings.warn(
                f"Column '{column.name}': Identity columns are not supported by excel dialect; auto-increment will not be applied",
                stacklevel=2,
            )

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
        pk_columns = {col.name for col in table.primary_key.columns}
        pk_columns_ordered = [
            self.preparer.format_column(col) for col in table.primary_key.columns
        ]
        inline_pk = len(pk_columns) == 1
        seen_unique: set[tuple[str, ...]] = set()
        seen_check: set[str] = set()
        for col in table.columns:
            col_name = self.preparer.format_column(col)
            col_type = self.dialect.type_compiler.process(col.type)
            constraints: list[str] = []
            if col.nullable is False and col.name not in pk_columns:
                constraints.append("NOT NULL")
            if inline_pk and col.name in pk_columns:
                constraints.append("PRIMARY KEY")
            self._warn_unsupported_constraints(
                col,
                context_msg="CREATE TABLE",
                seen_unique=seen_unique,
                seen_check=seen_check,
            )
            self._warn_unsupported_generated_columns(col)
            suffix = f" {' '.join(constraints)}" if constraints else ""
            columns.append(f"{col_name} {col_type}{suffix}")

        if len(pk_columns_ordered) > 1:
            columns.append(f"PRIMARY KEY ({', '.join(pk_columns_ordered)})")

        for constraint in table.constraints:
            if isinstance(constraint, UniqueConstraint):
                unique_key = tuple(col.name for col in constraint.columns)
                if unique_key not in seen_unique:
                    warnings.warn(
                        "Excel dialect does not enforce UNIQUE constraints (CREATE TABLE)",
                        stacklevel=2,
                    )
                    seen_unique.add(unique_key)
            if isinstance(constraint, CheckConstraint):
                check_key = str(constraint.sqltext)
                if check_key not in seen_check:
                    warnings.warn(
                        "Excel dialect does not enforce CHECK constraints (CREATE TABLE)",
                        stacklevel=2,
                    )
                    seen_check.add(check_key)

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
        if column.nullable is False and column.primary_key is not True:
            constraints.append("NOT NULL")
        if column.primary_key is True:
            constraints.append("PRIMARY KEY")
        self._warn_unsupported_constraints(
            column,
            context_msg="ALTER TABLE ADD COLUMN",
            seen_unique=set(),
            seen_check=set(),
        )
        self._warn_unsupported_generated_columns(column)
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
