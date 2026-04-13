"""Tests for ExcelDialect — URL parsing, feature flags, connect args."""

from __future__ import annotations

import pytest
from sqlalchemy import create_engine
from sqlalchemy.engine import make_url


class TestURLParsing:
    """Test URL parsing and create_connect_args."""

    def test_relative_path(self):
        url = make_url("excel:///data.xlsx")
        assert url.database == "data.xlsx"

    def test_absolute_path(self):
        url = make_url("excel:////home/user/data.xlsx")
        assert url.database == "/home/user/data.xlsx"

    def test_nested_relative_path(self):
        url = make_url("excel:///path/to/data.xlsx")
        assert url.database == "path/to/data.xlsx"

    def test_empty_path_raises(self, tmp_path):
        with pytest.raises((ValueError, Exception)):
            create_engine("excel://").connect()


class TestDialectFlags:
    """Test dialect feature flags."""

    def test_name(self, engine):
        assert engine.dialect.name == "excel"

    def test_driver(self, engine):
        assert engine.dialect.driver == "dbapi"

    def test_paramstyle(self, engine):
        assert engine.dialect.default_paramstyle == "qmark"

    def test_supports_alter(self, engine):
        assert engine.dialect.supports_alter is True

    def test_no_sequences(self, engine):
        assert engine.dialect.supports_sequences is False

    def test_no_schemas(self, engine):
        assert engine.dialect.supports_schemas is False

    def test_no_views(self, engine):
        assert engine.dialect.supports_views is False

    def test_no_statement_cache(self, engine):
        assert engine.dialect.supports_statement_cache is False


class TestImportDbapi:
    """Test import_dbapi."""

    def test_import_dbapi_returns_module(self):
        from sqlalchemy_excel.dialect import ExcelDialect

        dbapi = ExcelDialect.import_dbapi()
        assert hasattr(dbapi, "connect")
        assert hasattr(dbapi, "apilevel")
        assert dbapi.apilevel == "2.0"
        assert dbapi.paramstyle == "qmark"


class TestConnectArgs:
    """Test create_connect_args."""

    def test_connect_args_file_path(self, tmp_xlsx):
        engine = create_engine(f"excel:///{tmp_xlsx}")
        dialect = engine.dialect
        url = make_url(f"excel:///{tmp_xlsx}")
        args, kwargs = dialect.create_connect_args(url)
        assert args == []
        assert kwargs["file_path"] == tmp_xlsx
        assert kwargs["engine"] == "openpyxl"
        assert kwargs["create"] is True
        assert kwargs["autocommit"] is False
        engine.dispose()


class TestConnection:
    """Test basic connection operations."""

    def test_connect_creates_file(self, tmp_xlsx):
        import os

        engine = create_engine(f"excel:///{tmp_xlsx}")
        with engine.connect() as _conn:
            pass
        assert os.path.exists(tmp_xlsx)
        engine.dispose()

    def test_do_ping(self, engine):
        with engine.connect() as conn:
            raw = conn.connection.dbapi_connection
            assert engine.dialect.do_ping(raw) is True

    def test_is_disconnect_returns_false(self, engine):
        assert engine.dialect.is_disconnect(Exception(), None, None) is False
