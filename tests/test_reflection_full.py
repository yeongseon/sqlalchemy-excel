from __future__ import annotations

import pytest
from openpyxl import load_workbook
from sqlalchemy import create_engine, exc, inspect, text


def _engine_for(tmp_path):
    return create_engine(f"excel:///{tmp_path / 'reflection.xlsx'}")


def test_reflection_fallback_columns_and_empty_pk(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.execute(text("CREATE TABLE users (id INTEGER, name TEXT)"))

    insp = inspect(engine)
    columns = insp.get_columns("users")
    assert [column["name"] for column in columns] == ["id", "name"]

    pk = insp.get_pk_constraint("users")
    assert pk == {"constrained_columns": [], "name": None}

    engine.dispose()


def test_reflection_empty_return_methods(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.execute(text("CREATE TABLE users (id INTEGER, name TEXT)"))

    insp = inspect(engine)
    assert insp.get_view_names() == []
    assert insp.get_foreign_keys("users") == []
    assert insp.get_indexes("users") == []
    assert insp.get_unique_constraints("users") == []
    assert insp.get_check_constraints("users") == []
    assert insp.get_table_comment("users") == {"text": None}
    assert insp.get_schema_names() == []

    engine.dispose()


def test_reflection_cleans_stale_metadata_when_sheet_is_missing(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.connect() as conn:
        import excel_dbapi

        raw_conn = conn.connection.dbapi_connection
        excel_dbapi.write_table_metadata(
            raw_conn,
            "ghost",
            [
                {
                    "name": "id",
                    "type_name": "INTEGER",
                    "nullable": False,
                    "primary_key": True,
                }
            ],
        )
        assert excel_dbapi.read_table_metadata(raw_conn, "ghost") is not None

    inspector = inspect(engine)
    with pytest.raises(exc.NoSuchTableError):
        inspector.get_columns("ghost")
    with pytest.raises(exc.NoSuchTableError):
        inspector.get_pk_constraint("ghost")

    with engine.connect() as conn:
        import excel_dbapi

        raw_conn = conn.connection.dbapi_connection
        assert excel_dbapi.read_table_metadata(raw_conn, "ghost") is None

    engine.dispose()


def test_reflection_rebuilds_columns_when_metadata_headers_are_stale(tmp_path) -> None:
    workbook_path = tmp_path / "reflection.xlsx"
    engine = create_engine(f"excel:///{workbook_path}")

    with engine.begin() as conn:
        conn.execute(text("CREATE TABLE users (id INTEGER, name TEXT)"))

    workbook = load_workbook(workbook_path)
    worksheet = workbook["users"]
    worksheet["A1"] = "user_id"
    worksheet["B1"] = "full_name"
    worksheet["C1"] = "nickname"
    workbook.save(workbook_path)

    engine.dispose()
    engine = create_engine(f"excel:///{workbook_path}")

    inspector = inspect(engine)
    columns = inspector.get_columns("users")
    assert [column["name"] for column in columns] == [
        "user_id",
        "full_name",
        "nickname",
    ]

    engine.dispose()


def test_table_scoped_reflection_methods_raise_for_missing_table(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    inspector = inspect(engine)

    with pytest.raises(exc.NoSuchTableError):
        inspector.get_foreign_keys("missing")
    with pytest.raises(exc.NoSuchTableError):
        inspector.get_indexes("missing")
    with pytest.raises(exc.NoSuchTableError):
        inspector.get_unique_constraints("missing")
    with pytest.raises(exc.NoSuchTableError):
        inspector.get_check_constraints("missing")
    with pytest.raises(exc.NoSuchTableError):
        inspector.get_table_comment("missing")

    engine.dispose()
