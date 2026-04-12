"""DDL compiler for the Excel dialect.

CREATE TABLE → creates a worksheet and writes column metadata.
DROP TABLE   → deletes the worksheet and removes metadata.
ALTER TABLE  → rejected (not supported).
"""

from __future__ import annotations

from typing import Any

from sqlalchemy import exc
from sqlalchemy.sql import compiler


class ExcelDDLCompiler(compiler.DDLCompiler):
    """Compiles DDL statements for excel-dbapi."""

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
