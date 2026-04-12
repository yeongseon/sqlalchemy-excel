"""Tests for ExcelTypeCompiler — type mapping."""

from __future__ import annotations

import pytest
from sqlalchemy import create_engine, exc
from sqlalchemy import types as sa_types


@pytest.fixture
def compiler(tmp_xlsx):
    engine = create_engine(f"excel:///{tmp_xlsx}")
    comp = engine.dialect.type_compiler
    yield comp
    engine.dispose()


class TestTypeMapping:
    """Test type mapping to Excel type strings."""

    def test_string_to_text(self, compiler):
        assert compiler.process(sa_types.String()) == "TEXT"

    def test_text_to_text(self, compiler):
        assert compiler.process(sa_types.Text()) == "TEXT"

    def test_varchar_to_text(self, compiler):
        assert compiler.process(sa_types.VARCHAR()) == "TEXT"

    def test_integer_to_integer(self, compiler):
        assert compiler.process(sa_types.Integer()) == "INTEGER"

    def test_smallint_to_integer(self, compiler):
        assert compiler.process(sa_types.SmallInteger()) == "INTEGER"

    def test_bigint_to_integer(self, compiler):
        assert compiler.process(sa_types.BigInteger()) == "INTEGER"

    def test_float_to_float(self, compiler):
        assert compiler.process(sa_types.Float()) == "FLOAT"

    def test_numeric_to_float(self, compiler):
        assert compiler.process(sa_types.Numeric()) == "FLOAT"

    def test_boolean_to_boolean(self, compiler):
        assert compiler.process(sa_types.Boolean()) == "BOOLEAN"

    def test_date_to_date(self, compiler):
        assert compiler.process(sa_types.Date()) == "DATE"

    def test_datetime_to_datetime(self, compiler):
        assert compiler.process(sa_types.DateTime()) == "DATETIME"


class TestUnsupportedTypes:
    """Test that unsupported types raise CompileError."""

    def test_blob_rejected(self, compiler):
        with pytest.raises(exc.CompileError, match="BLOB"):
            compiler.process(sa_types.BLOB())

    def test_json_rejected(self, compiler):
        with pytest.raises(exc.CompileError, match="JSON"):
            compiler.process(sa_types.JSON())

    def test_large_binary_rejected(self, compiler):
        with pytest.raises(exc.CompileError, match="LargeBinary"):
            compiler.process(sa_types.LargeBinary())
