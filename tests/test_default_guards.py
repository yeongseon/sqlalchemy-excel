from __future__ import annotations

import warnings

from sqlalchemy import (
    Column,
    Computed,
    Identity,
    Integer,
    MetaData,
    String,
    Table,
    create_engine,
    text,
)


def test_server_default_emits_warning(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    _ = Table(
        "with_server_default",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String, server_default=text("'anon'")),
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        metadata.create_all(engine)

    messages = [str(item.message) for item in caught]
    assert any(
        "Excel dialect does not support server_default; value will be ignored"
        in message
        for message in messages
    )
    engine.dispose()


def test_explicit_autoincrement_emits_warning(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    _ = Table(
        "with_autoincrement",
        metadata,
        Column("id", Integer, primary_key=True, autoincrement=True),
        Column("name", String),
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        metadata.create_all(engine)

    messages = [str(item.message) for item in caught]
    assert any(
        "Excel dialect does not support autoincrement=True; value must be set explicitly"
        in message
        for message in messages
    )
    engine.dispose()


def test_normal_columns_do_not_emit_default_warnings(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    _ = Table(
        "no_guard_warning",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        metadata.create_all(engine)

    messages = [str(item.message) for item in caught]
    assert not any("Excel dialect does not support" in message for message in messages)
    engine.dispose()


def test_computed_column_emits_warning(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    _ = Table(
        "with_computed",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("name_upper", String, Computed("name")),
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        metadata.create_all(engine)

    messages = [str(item.message) for item in caught]
    assert any(
        "Computed columns are not supported by excel dialect; the expression will be ignored"
        in message
        for message in messages
    )
    engine.dispose()


def test_identity_column_emits_warning(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    metadata = MetaData()
    _ = Table(
        "with_identity",
        metadata,
        Column("id", Integer, Identity(), primary_key=True),
        Column("name", String),
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        metadata.create_all(engine)

    messages = [str(item.message) for item in caught]
    assert any(
        "Identity columns are not supported by excel dialect; auto-increment will not be applied"
        in message
        for message in messages
    )
    engine.dispose()
