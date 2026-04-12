from __future__ import annotations

import pytest
from sqlalchemy import exc
from sqlalchemy import types as sa_types

from sqlalchemy_excel.dialect import ExcelDialect
from sqlalchemy_excel.types import ExcelTypeCompiler


@pytest.fixture
def type_compiler() -> ExcelTypeCompiler:
    return ExcelTypeCompiler(ExcelDialect())


def test_type_compiler_core_mappings_via_process(
    type_compiler: ExcelTypeCompiler,
) -> None:
    assert type_compiler.process(sa_types.Text()) == "TEXT"
    assert type_compiler.process(sa_types.Integer()) == "INTEGER"
    assert type_compiler.process(sa_types.Float()) == "FLOAT"
    assert type_compiler.process(sa_types.Boolean()) == "BOOLEAN"
    assert type_compiler.process(sa_types.Date()) == "DATE"
    assert type_compiler.process(sa_types.DateTime()) == "DATETIME"


def test_type_compiler_all_text_visitors(type_compiler: ExcelTypeCompiler) -> None:
    assert type_compiler.visit_STRING(sa_types.String()) == "TEXT"
    assert type_compiler.visit_TEXT(sa_types.Text()) == "TEXT"
    assert type_compiler.visit_NVARCHAR(sa_types.NVARCHAR()) == "TEXT"
    assert type_compiler.visit_VARCHAR(sa_types.VARCHAR()) == "TEXT"
    assert type_compiler.visit_CHAR(sa_types.CHAR()) == "TEXT"
    assert type_compiler.visit_NCHAR(sa_types.NCHAR()) == "TEXT"
    assert type_compiler.visit_CLOB(sa_types.CLOB()) == "TEXT"


def test_type_compiler_all_numeric_visitors(type_compiler: ExcelTypeCompiler) -> None:
    assert type_compiler.visit_INTEGER(sa_types.Integer()) == "INTEGER"
    assert type_compiler.visit_SMALLINT(sa_types.SmallInteger()) == "INTEGER"
    assert type_compiler.visit_BIGINT(sa_types.BigInteger()) == "INTEGER"
    assert type_compiler.visit_FLOAT(sa_types.Float()) == "FLOAT"
    assert type_compiler.visit_REAL(sa_types.REAL()) == "FLOAT"
    assert type_compiler.visit_DOUBLE(sa_types.DOUBLE()) == "FLOAT"
    assert type_compiler.visit_DOUBLE_PRECISION(sa_types.DOUBLE_PRECISION()) == "FLOAT"
    assert type_compiler.visit_NUMERIC(sa_types.Numeric()) == "FLOAT"
    assert type_compiler.visit_DECIMAL(sa_types.DECIMAL()) == "FLOAT"


def test_type_compiler_all_temporal_and_misc_visitors(
    type_compiler: ExcelTypeCompiler,
) -> None:
    assert type_compiler.visit_BOOLEAN(sa_types.Boolean()) == "BOOLEAN"
    assert type_compiler.visit_DATE(sa_types.Date()) == "DATE"
    assert type_compiler.visit_DATETIME(sa_types.DateTime()) == "DATETIME"
    assert type_compiler.visit_TIMESTAMP(sa_types.TIMESTAMP()) == "DATETIME"
    assert type_compiler.visit_TIME(sa_types.TIME()) == "TEXT"
    assert type_compiler.visit_uuid(sa_types.Uuid()) == "TEXT"


@pytest.mark.parametrize(
    ("call", "message"),
    [
        (lambda c: c.visit_BLOB(sa_types.BLOB()), "BLOB"),
        (lambda c: c.visit_BINARY(sa_types.BINARY()), "BINARY"),
        (lambda c: c.visit_VARBINARY(sa_types.VARBINARY()), "VARBINARY"),
        (lambda c: c.visit_JSON(sa_types.JSON()), "JSON"),
        (lambda c: c.visit_ARRAY(sa_types.ARRAY(sa_types.Integer())), "ARRAY"),
        (lambda c: c.visit_large_binary(sa_types.LargeBinary()), "LargeBinary"),
    ],
)
def test_type_compiler_unsupported_type_visitors_raise_compile_error(
    type_compiler: ExcelTypeCompiler,
    call,
    message: str,
) -> None:
    with pytest.raises(exc.CompileError, match=message):
        call(type_compiler)
