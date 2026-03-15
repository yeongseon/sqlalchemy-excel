from __future__ import annotations

import importlib
from typing import TYPE_CHECKING

import pytest
from openpyxl import Workbook
from sqlalchemy import select

_conftest = importlib.import_module("tests.conftest")
Base = _conftest.Base
SimpleUser = _conftest.SimpleUser

_sqlalchemy_excel = importlib.import_module("sqlalchemy_excel")
ExcelImporter = _sqlalchemy_excel.ExcelImporter
ExcelMapping = _sqlalchemy_excel.ExcelMapping
ImportResult = _sqlalchemy_excel.ImportResult

_sqlalchemy_excel_exceptions = importlib.import_module("sqlalchemy_excel.exceptions")
ImportError_ = _sqlalchemy_excel_exceptions.ImportError_

_sqlalchemy_excel_strategies = importlib.import_module(
    "sqlalchemy_excel.load.strategies"
)
DryRunStrategy = _sqlalchemy_excel_strategies.DryRunStrategy
InsertStrategy = _sqlalchemy_excel_strategies.InsertStrategy
UpsertStrategy = _sqlalchemy_excel_strategies.UpsertStrategy

if TYPE_CHECKING:
    from pathlib import Path


def _create_test_xlsx(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = SimpleUser.__tablename__
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


def test_import_result_total() -> None:
    result = ImportResult(inserted=5, updated=2, skipped=1, failed=0)

    assert result.total == 8


def test_import_result_summary() -> None:
    result = ImportResult(inserted=3, updated=1, skipped=2, failed=4)

    summary = result.summary()

    assert "inserted=3" in summary
    assert "updated=1" in summary
    assert "skipped=2" in summary
    assert "failed=4" in summary
    assert "total=10" in summary


def test_import_result_defaults() -> None:
    result = ImportResult()

    assert result.inserted == 0
    assert result.updated == 0
    assert result.skipped == 0
    assert result.failed == 0
    assert result.errors == []
    assert result.duration_ms == 0.0


def test_insert_basic(session) -> None:
    strategy = InsertStrategy()
    rows = [
        {"id": 1, "name": "Alice", "email": "alice@example.com", "age": 30},
        {"id": 2, "name": "Bob", "email": "bob@example.com", "age": 28},
        {"id": 3, "name": "Carol", "email": "carol@example.com", "age": None},
    ]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=100,
    )

    users = session.execute(select(SimpleUser).order_by(SimpleUser.id)).scalars().all()
    assert result.inserted == 3
    assert result.failed == 0
    assert [user.name for user in users] == ["Alice", "Bob", "Carol"]


def test_insert_duplicate_key(session) -> None:
    session.add(SimpleUser(id=1, name="Existing", email="existing@example.com", age=20))
    session.commit()

    strategy = InsertStrategy()
    rows = [{"id": 1, "name": "Duplicate", "email": "dup@example.com", "age": 44}]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=100,
    )

    assert result.inserted == 0
    assert result.failed == 1

    # After a failed insert the session may need rollback before querying
    if session.is_active:
        try:
            count = session.query(SimpleUser).count()
        except Exception:
            session.rollback()
            count = session.query(SimpleUser).count()
    else:
        session.rollback()
        count = session.query(SimpleUser).count()
    assert count == 1


def test_insert_batch_size(session) -> None:
    strategy = InsertStrategy()
    rows = [
        {"id": 1, "name": "U1", "email": "u1@example.com", "age": 21},
        {"id": 2, "name": "U2", "email": "u2@example.com", "age": 22},
        {"id": 3, "name": "U3", "email": "u3@example.com", "age": 23},
        {"id": 4, "name": "U4", "email": "u4@example.com", "age": 24},
        {"id": 5, "name": "U5", "email": "u5@example.com", "age": 25},
    ]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=2,
    )

    assert result.inserted == 5
    assert result.failed == 0
    assert session.query(SimpleUser).count() == 5


def test_upsert_insert_new(session) -> None:
    strategy = UpsertStrategy()
    rows = [{"id": 1, "name": "Alice", "email": "alice@example.com", "age": 30}]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=100,
    )

    assert result.inserted == 1
    assert result.updated == 0
    assert session.query(SimpleUser).count() == 1


def test_upsert_update_existing(session) -> None:
    session.add(SimpleUser(id=1, name="Before", email="before@example.com", age=22))
    session.commit()

    strategy = UpsertStrategy()
    rows = [{"id": 1, "name": "After", "email": "after@example.com", "age": 29}]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=100,
    )

    updated = session.execute(select(SimpleUser).where(SimpleUser.id == 1)).scalar_one()
    assert result.inserted == 0
    assert result.updated == 1
    assert updated.name == "After"
    assert updated.email == "after@example.com"
    assert updated.age == 29


def test_upsert_mixed(session) -> None:
    session.add(SimpleUser(id=1, name="Existing", email="existing@example.com", age=20))
    session.commit()

    strategy = UpsertStrategy()
    rows = [
        {"id": 1, "name": "Updated", "email": "updated@example.com", "age": 31},
        {"id": 2, "name": "New", "email": "new@example.com", "age": 26},
    ]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=100,
    )

    by_id = {
        user.id: user
        for user in session.execute(select(SimpleUser).order_by(SimpleUser.id))
        .scalars()
        .all()
    }
    assert result.inserted == 1
    assert result.updated == 1
    assert by_id[1].name == "Updated"
    assert by_id[2].name == "New"


def test_upsert_requires_key_columns(session) -> None:
    strategy = UpsertStrategy()

    with pytest.raises(ValueError, match="requires at least one key column"):
        _ = strategy.execute(
            session=session,
            model_class=SimpleUser,
            rows=[],
            key_columns=[],
            batch_size=100,
        )


def test_dry_run_counts_without_persisting(session) -> None:
    strategy = DryRunStrategy()
    rows = [
        {"id": 1, "name": "A", "email": "a@example.com", "age": 10},
        {"id": 2, "name": "B", "email": "b@example.com", "age": 11},
        {"id": 3, "name": "C", "email": "c@example.com", "age": 12},
    ]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=100,
    )

    assert result.inserted == 3
    assert result.failed == 0
    assert session.query(SimpleUser).count() == 0


def test_importer_insert(tmp_path: Path, session) -> None:
    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)
    source = tmp_path / "users_insert.xlsx"
    _create_test_xlsx(
        source,
        headers=["id", "name", "email", "age"],
        rows=[
            [1, "Alice", "alice@example.com", 30],
            [2, "Bob", "bob@example.com", 28],
        ],
    )

    result = importer.insert(source)

    users = session.execute(select(SimpleUser).order_by(SimpleUser.id)).scalars().all()
    assert result.inserted == 2
    assert result.failed == 0
    assert [user.email for user in users] == ["alice@example.com", "bob@example.com"]


def test_importer_dry_run(tmp_path: Path, session) -> None:
    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)
    source = tmp_path / "users_dry_run.xlsx"
    _create_test_xlsx(
        source,
        headers=["id", "name", "email", "age"],
        rows=[
            [1, "Alice", "alice@example.com", 30],
            [2, "Bob", "bob@example.com", 28],
        ],
    )

    result = importer.dry_run(source)

    assert result.inserted == 2
    assert result.failed == 0
    assert session.query(SimpleUser).count() == 0


def test_importer_empty_mappings_raises(session) -> None:
    with pytest.raises(ImportError_, match="At least one ExcelMapping is required"):
        _ = ExcelImporter([], session)


def test_importer_insert_skip_validation(tmp_path: Path, session, engine) -> None:
    assert issubclass(SimpleUser, Base)
    assert engine is not None

    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)
    source = tmp_path / "users_skip_validation.xlsx"
    _create_test_xlsx(
        source,
        headers=["id", "name", "email", "age"],
        rows=[[1, "Alice", "alice@example.com", 30]],
    )

    result = importer.insert(source, validate=False)

    assert result.inserted == 1
    assert result.failed == 0
    assert session.query(SimpleUser).count() == 1


def test_upsert_missing_key_column_records_failure(session) -> None:
    strategy = UpsertStrategy()
    rows = [
        {"name": "Missing Id", "email": "missing@example.com", "age": 31},
        {"id": 2, "name": "Valid", "email": "valid@example.com", "age": 30},
    ]

    result = strategy.execute(
        session=session,
        model_class=SimpleUser,
        rows=rows,
        key_columns=["id"],
        batch_size=100,
    )

    assert result.failed == 1
    assert result.inserted == 1
    assert any("Missing upsert key column(s): id" in error for error in result.errors)


def test_importer_aligns_special_headers_by_normalization(
    tmp_path: Path, session
) -> None:
    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)
    source = tmp_path / "users_special_headers.xlsx"
    _create_test_xlsx(
        source,
        headers=["ID", "Name!!!", "Email@", "Age($)"],
        rows=[[1, "Alice", "alice@example.com", 30]],
    )

    result = importer.insert(source, validate=False)

    assert result.failed == 0
    assert result.inserted == 1
    saved = session.execute(select(SimpleUser).where(SimpleUser.id == 1)).scalar_one()
    assert saved.name == "Alice"


def test_importer_create_reader_uses_read_only(session) -> None:
    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)

    reader = importer._create_reader()

    assert reader.read_only is True


def test_importer_insert_validation_failure_blocks_write(tmp_path: Path, session) -> None:
    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)
    source = tmp_path / "users_invalid_insert.xlsx"
    _create_test_xlsx(
        source,
        headers=["id", "name", "email", "age"],
        rows=[
            [1, "Alice", "alice@example.com", "not-an-int"],
        ],
    )

    result = importer.insert(source, validate=True)

    assert result.failed == 1
    assert result.inserted == 0
    assert session.query(SimpleUser).count() == 0


def test_importer_upsert_validation_failure_blocks_write(tmp_path: Path, session) -> None:
    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)
    source = tmp_path / "users_invalid_upsert.xlsx"
    _create_test_xlsx(
        source,
        headers=["id", "name", "email", "age"],
        rows=[
            [1, None, "alice@example.com", 10],
        ],
    )

    result = importer.upsert(source, validate=True)

    assert result.failed == 1
    assert result.inserted == 0
    assert result.updated == 0


def test_importer_dry_run_does_not_persist_even_with_valid_data(
    tmp_path: Path, session
) -> None:
    mapping = ExcelMapping.from_model(SimpleUser)
    importer = ExcelImporter([mapping], session)
    source = tmp_path / "users_valid_dry_run.xlsx"
    _create_test_xlsx(
        source,
        headers=["id", "name", "email", "age"],
        rows=[[99, "Dry", "dry@example.com", 21]],
    )

    result = importer.dry_run(source, validate=True)

    assert result.inserted == 1
    assert result.failed == 0
    assert session.query(SimpleUser).count() == 0
