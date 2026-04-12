"""Tests for ExcelGraphDialect — URL parsing and Graph API integration."""

from __future__ import annotations

import httpx
import pytest
from sqlalchemy import create_engine, text
from sqlalchemy.dialects import registry
from sqlalchemy.engine import make_url


# ---------------------------------------------------------------------------
# Mock transport (minimal Graph API stub)
# ---------------------------------------------------------------------------

def _graph_handler(request: httpx.Request) -> httpx.Response:
    """Stateless mock handler for read-only Graph API tests."""
    path = request.url.path
    method = request.method

    if path.endswith("/createSession"):
        return httpx.Response(201, json={"id": "sess-graph-test"})
    if path.endswith("/closeSession"):
        return httpx.Response(204)

    if (
        path.endswith("/worksheets") or "/worksheets?" in str(request.url)
    ) and method == "GET":
        return httpx.Response(
            200,
            json={"value": [{"id": "ws-sheet1", "name": "Sheet1"}]},
        )

    if "usedRange" in path and method == "GET":
        return httpx.Response(
            200,
            json={
                "values": [
                    ["id", "name", "value"],
                    [1, "Alice", 100],
                    [2, "Bob", 200],
                ]
            },
        )

    return httpx.Response(404)


# ---------------------------------------------------------------------------
# URL Parsing Tests
# ---------------------------------------------------------------------------

class TestGraphURLParsing:
    def test_url_components(self):
        url = make_url("excel+graph:///drv-abc/itm-xyz")
        assert url.get_backend_name() == "excel"
        assert url.get_driver_name() == "graph"
        assert url.database == "drv-abc/itm-xyz"

    def test_url_with_host_ignored(self):
        """Host part (tenant_id) is allowed but unused."""
        url = make_url("excel+graph://my-tenant/drv-abc/itm-xyz")
        assert url.host == "my-tenant"
        assert url.database == "drv-abc/itm-xyz"

    def test_create_connect_args_basic(self):
        dialect = registry.load("excel.graph")()
        url = make_url("excel+graph:///drv-abc/itm-xyz")
        args, kwargs = dialect.create_connect_args(url)
        assert args == []
        assert kwargs["file_path"] == "msgraph://drives/drv-abc/items/itm-xyz"
        assert kwargs["engine"] == "graph"
        assert kwargs["autocommit"] is True
        assert kwargs["create"] is False

    def test_create_connect_args_url_decoding(self):
        """Drive/item IDs with percent-encoded chars should be decoded."""
        dialect = registry.load("excel.graph")()
        url = make_url("excel+graph:///b%21abc/itm%2D123")
        _, kwargs = dialect.create_connect_args(url)
        assert kwargs["file_path"] == "msgraph://drives/b!abc/items/itm-123"

    def test_create_connect_args_readonly_false(self):
        dialect = registry.load("excel.graph")()
        url = make_url("excel+graph:///drv/itm?readonly=false")
        _, kwargs = dialect.create_connect_args(url)
        assert kwargs.get("readonly") is False

    def test_create_connect_args_readonly_true(self):
        dialect = registry.load("excel.graph")()
        url = make_url("excel+graph:///drv/itm?readonly=true")
        _, kwargs = dialect.create_connect_args(url)
        assert kwargs.get("readonly") is True

    def test_empty_path_raises(self):
        dialect = registry.load("excel.graph")()
        url = make_url("excel+graph://")
        with pytest.raises(ValueError, match="No drive/item path"):
            _ = dialect.create_connect_args(url)

    def test_single_segment_raises(self):
        dialect = registry.load("excel.graph")()
        url = make_url("excel+graph:///only-one")
        with pytest.raises(ValueError, match="path segments"):
            _ = dialect.create_connect_args(url)

    def test_three_segments_raises(self):
        dialect = registry.load("excel.graph")()
        url = make_url("excel+graph:///a/b/c")
        with pytest.raises(ValueError, match="path segments"):
            _ = dialect.create_connect_args(url)


# ---------------------------------------------------------------------------
# Dialect Feature Flags
# ---------------------------------------------------------------------------

class TestGraphDialectFlags:
    def test_driver(self):
        d = registry.load("excel.graph")()
        assert d.driver == "graph"

    def test_name(self):
        d = registry.load("excel.graph")()
        assert d.name == "excel"


# ---------------------------------------------------------------------------
# Integration: SELECT via mock transport
# ---------------------------------------------------------------------------

class TestGraphDialectIntegration:
    def test_select_via_engine(self):
        """Full round-trip: create_engine → connect → SELECT."""
        transport = httpx.MockTransport(_graph_handler)
        engine = create_engine(
            "excel+graph:///drv-test/itm-test",
            connect_args={
                "credential": "test-token",
                "transport": transport,
            },
        )
        with engine.connect() as conn:
            result = conn.execute(text("SELECT * FROM Sheet1"))
            rows = result.fetchall()
            assert len(rows) == 2
            assert rows[0] == (1, "Alice", 100)
        engine.dispose()

    def test_select_with_where(self):
        transport = httpx.MockTransport(_graph_handler)
        engine = create_engine(
            "excel+graph:///drv-test/itm-test",
            connect_args={
                "credential": "test-token",
                "transport": transport,
            },
        )
        with engine.connect() as conn:
            result = conn.execute(
                text("SELECT name FROM Sheet1 WHERE id = :id"),
                {"id": 1},
            )
            rows = result.fetchall()
            assert len(rows) == 1
            assert rows[0] == ("Alice",)
        engine.dispose()
