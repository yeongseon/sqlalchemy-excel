"""Type compiler for the Excel dialect.

Maps SQLAlchemy types to the type names that excel-dbapi understands:
    TEXT, INTEGER, FLOAT, BOOLEAN, DATE, DATETIME
"""

from __future__ import annotations

from typing import Any

from sqlalchemy import exc
from sqlalchemy.sql import compiler


class ExcelTypeCompiler(compiler.GenericTypeCompiler):
    """Maps SQLAlchemy column types to Excel-compatible type strings."""

    def visit_STRING(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_TEXT(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_NVARCHAR(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_VARCHAR(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_CHAR(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_NCHAR(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_CLOB(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_INTEGER(self, type_: Any, **kw: Any) -> str:
        return "INTEGER"

    def visit_SMALLINT(self, type_: Any, **kw: Any) -> str:
        return "INTEGER"

    def visit_BIGINT(self, type_: Any, **kw: Any) -> str:
        return "INTEGER"

    def visit_FLOAT(self, type_: Any, **kw: Any) -> str:
        return "FLOAT"

    def visit_REAL(self, type_: Any, **kw: Any) -> str:
        return "FLOAT"

    def visit_DOUBLE(self, type_: Any, **kw: Any) -> str:
        return "FLOAT"

    def visit_DOUBLE_PRECISION(self, type_: Any, **kw: Any) -> str:
        return "FLOAT"

    def visit_NUMERIC(self, type_: Any, **kw: Any) -> str:
        return "FLOAT"

    def visit_DECIMAL(self, type_: Any, **kw: Any) -> str:
        return "FLOAT"

    def visit_BOOLEAN(self, type_: Any, **kw: Any) -> str:
        return "BOOLEAN"

    def visit_DATE(self, type_: Any, **kw: Any) -> str:
        return "DATE"

    def visit_DATETIME(self, type_: Any, **kw: Any) -> str:
        return "DATETIME"

    def visit_TIMESTAMP(self, type_: Any, **kw: Any) -> str:
        return "DATETIME"

    def visit_TIME(self, type_: Any, **kw: Any) -> str:
        return "TEXT"

    def visit_BLOB(self, type_: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support BLOB type")

    def visit_BINARY(self, type_: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support BINARY type")

    def visit_VARBINARY(self, type_: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support VARBINARY type")

    def visit_JSON(self, type_: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support JSON type")

    def visit_ARRAY(self, type_: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support ARRAY type")

    def visit_large_binary(self, type_: Any, **kw: Any) -> str:
        raise exc.CompileError("Excel dialect does not support LargeBinary type")

    def visit_uuid(self, type_: Any, **kw: Any) -> str:
        return "TEXT"
