"""Bulk and merge write path tests.

These operations work with the Excel dialect, but SQLAlchemy bulk APIs still
use their normal caveats (for example, bypassing unit-of-work bookkeeping).
"""

from __future__ import annotations

from typing import Any, cast

from sqlalchemy import Float, Integer, String, create_engine, insert, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column


class Base(DeclarativeBase):
    pass


class BulkRow(Base):
    __tablename__ = "bulk_rows"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String)
    score: Mapped[float | None] = mapped_column(Float, nullable=True)
    note: Mapped[str | None] = mapped_column(String, nullable=True)


def test_session_merge_updates_detached_object(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    Base.metadata.create_all(engine)

    with Session(engine) as session:
        session.add(BulkRow(id=1, name="before", score=1.0, note="old"))
        session.commit()

    detached = BulkRow(id=1, name="after", score=2.5, note=None)
    with Session(engine) as session:
        merged = session.merge(detached)
        session.commit()
        assert merged.name == "after"

    with Session(engine) as session:
        row = session.query(BulkRow).filter(BulkRow.id == 1).one()
        assert row.name == "after"
        assert row.score == 2.5
        assert row.note is None

    engine.dispose()


def test_bulk_save_objects_writes_rows(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    Base.metadata.create_all(engine)

    with Session(engine) as session:
        session.bulk_save_objects(
            [
                BulkRow(id=1, name="alpha", score=1.5, note="a"),
                BulkRow(id=2, name="beta", score=None, note=None),
            ]
        )
        session.commit()

    with Session(engine) as session:
        rows = session.query(BulkRow).order_by(BulkRow.id).all()
        assert [(row.id, row.name, row.score, row.note) for row in rows] == [
            (1, "alpha", 1.5, "a"),
            (2, "beta", None, None),
        ]

    engine.dispose()


def test_bulk_insert_mappings_writes_rows(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    Base.metadata.create_all(engine)

    with Session(engine) as session:
        session.bulk_insert_mappings(
            cast("Any", BulkRow),
            [
                {"id": 1, "name": "one", "score": 3.0, "note": "n1"},
                {"id": 2, "name": "two", "score": None, "note": None},
            ],
        )
        session.commit()

    with Session(engine) as session:
        rows = session.query(BulkRow).order_by(BulkRow.id).all()
        assert [(row.id, row.name, row.score, row.note) for row in rows] == [
            (1, "one", 3.0, "n1"),
            (2, "two", None, None),
        ]

    engine.dispose()


def test_core_executemany_insert_varied_types(tmp_xlsx: str) -> None:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    Base.metadata.create_all(engine)

    with engine.begin() as conn:
        _ = conn.execute(
            insert(BulkRow),
            [
                {"id": 10, "name": "int", "score": 10.0, "note": "ok"},
                {"id": 11, "name": "none", "score": None, "note": None},
                {"id": 12, "name": "float", "score": 3.14159, "note": "pi"},
            ],
        )

    with engine.connect() as conn:
        rows = conn.execute(select(BulkRow.__table__).order_by(BulkRow.id)).all()
        assert rows == [
            (10, "int", 10.0, "ok"),
            (11, "none", None, None),
            (12, "float", 3.14159, "pi"),
        ]

    engine.dispose()
