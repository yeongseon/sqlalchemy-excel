import pytest
from sqlalchemy import Column, Integer, MetaData, String, Table, select
from sqlalchemy.exc import InvalidRequestError

from sqlalchemy_excel import insert


@pytest.fixture
def metadata():
    return MetaData()


@pytest.fixture
def items_table(metadata):
    return Table(
        "items",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("age", Integer),
        Column("code", String),
    )


@pytest.fixture
def composite_items_table(metadata):
    return Table(
        "composite_items",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("code", String, primary_key=True),
        Column("name", String),
        Column("age", Integer),
    )


@pytest.fixture
def populated_engine(engine, metadata, items_table, composite_items_table):
    metadata.create_all(engine)
    return engine


class TestUpsertDoNothing:
    def test_conflict_skipped(self, populated_engine, items_table):
        with populated_engine.begin() as conn:
            conn.execute(insert(items_table).values(id=1, name="Alice", age=30, code="A"))

        with populated_engine.begin() as conn:
            stmt = insert(items_table).values(id=1, name="New", age=40, code="B")
            conn.execute(stmt.on_conflict_do_nothing(index_elements=["id"]))

        with populated_engine.connect() as conn:
            rows = conn.execute(select(items_table).order_by(items_table.c.id)).all()
            assert rows == [(1, "Alice", 30, "A")]

    def test_no_conflict_inserts(self, populated_engine, items_table):
        with populated_engine.begin() as conn:
            conn.execute(insert(items_table).values(id=1, name="Alice", age=30, code="A"))

        with populated_engine.begin() as conn:
            stmt = insert(items_table).values(id=2, name="Bob", age=25, code="B")
            conn.execute(stmt.on_conflict_do_nothing(index_elements=["id"]))

        with populated_engine.connect() as conn:
            rows = conn.execute(select(items_table).order_by(items_table.c.id)).all()
            assert rows == [(1, "Alice", 30, "A"), (2, "Bob", 25, "B")]

    def test_multi_row_mixed(self, populated_engine, items_table):
        with populated_engine.begin() as conn:
            conn.execute(insert(items_table), [{"id": 1, "name": "Alice", "age": 30, "code": "A"}])

        with populated_engine.begin() as conn:
            stmt = insert(items_table).on_conflict_do_nothing(index_elements=["id"])
            conn.execute(
                stmt,
                [
                    {"id": 1, "name": "Conflict", "age": 99, "code": "X"},
                    {"id": 2, "name": "Bob", "age": 25, "code": "B"},
                    {"id": 3, "name": "Cara", "age": 27, "code": "C"},
                ],
            )

        with populated_engine.connect() as conn:
            rows = conn.execute(select(items_table).order_by(items_table.c.id)).all()
            assert rows == [
                (1, "Alice", 30, "A"),
                (2, "Bob", 25, "B"),
                (3, "Cara", 27, "C"),
            ]


class TestUpsertDoUpdate:
    def test_literal_set(self, populated_engine, items_table):
        with populated_engine.begin() as conn:
            conn.execute(insert(items_table).values(id=1, name="Alice", age=30, code="A"))

        with populated_engine.begin() as conn:
            stmt = insert(items_table).values(id=1, name="New", age=99, code="B")
            conn.execute(
                stmt.on_conflict_do_update(
                    index_elements=["id"],
                    set_={"age": 42},
                )
            )

        with populated_engine.connect() as conn:
            rows = conn.execute(select(items_table).order_by(items_table.c.id)).all()
            assert rows == [(1, "Alice", 42, "A")]

    def test_excluded_reference(self, populated_engine, items_table):
        with populated_engine.begin() as conn:
            conn.execute(insert(items_table).values(id=1, name="Alice", age=30, code="A"))

        with populated_engine.begin() as conn:
            stmt = insert(items_table).values(id=1, name="Renamed", age=31, code="A")
            conn.execute(
                stmt.on_conflict_do_update(
                    index_elements=["id"],
                    set_={"name": stmt.excluded.name},
                )
            )

        with populated_engine.connect() as conn:
            rows = conn.execute(select(items_table).order_by(items_table.c.id)).all()
            assert rows == [(1, "Renamed", 30, "A")]

    def test_expression_set_value(self, populated_engine, items_table):
        with populated_engine.begin() as conn:
            conn.execute(insert(items_table).values(id=1, name="Alice", age=30, code="A"))

        stmt = insert(items_table).values(id=1, name="New", age=99, code="B")
        upsert_stmt = stmt.on_conflict_do_update(
            index_elements=["id"],
            set_={"age": items_table.c.age + 1},
        )
        compiled = " ".join(str(upsert_stmt.compile(dialect=populated_engine.dialect)).split())
        assert "items.age" not in compiled
        assert " SET age = (age + " in compiled

        with populated_engine.begin() as conn:
            conn.execute(upsert_stmt)

        with populated_engine.connect() as conn:
            rows = conn.execute(select(items_table).order_by(items_table.c.id)).all()
            assert rows == [(1, "Alice", 31, "A")]

    def test_set_with_table_columns(self, populated_engine, items_table):
        upsert_stmt = insert(items_table).on_conflict_do_update(
            index_elements=["id"],
            set_=items_table.c,
        )
        compiled = " ".join(str(upsert_stmt.compile(dialect=populated_engine.dialect)).split())
        assert "items." not in compiled
        assert " SET id = id, name = name, age = age, code = code" in compiled

    def test_multi_column_conflict_target(self, populated_engine, composite_items_table):
        with populated_engine.begin() as conn:
            conn.execute(
                insert(composite_items_table).values(id=1, code="A", name="Alice", age=30)
            )

        with populated_engine.begin() as conn:
            stmt = insert(composite_items_table).values(
                id=1,
                code="A",
                name="Alice2",
                age=33,
            )
            conn.execute(
                stmt.on_conflict_do_update(
                    index_elements=["id", "code"],
                    set_={"age": stmt.excluded.age},
                )
            )

        with populated_engine.connect() as conn:
            rows = conn.execute(
                select(composite_items_table).order_by(
                    composite_items_table.c.id,
                    composite_items_table.c.code,
                )
            ).all()
            assert rows == [(1, "A", "Alice", 33)]

    def test_multi_row(self, populated_engine, items_table):
        with populated_engine.begin() as conn:
            conn.execute(
                insert(items_table),
                [
                    {"id": 1, "name": "Alice", "age": 30, "code": "A"},
                    {"id": 2, "name": "Bob", "age": 25, "code": "B"},
                ],
            )

        with populated_engine.begin() as conn:
            stmt = insert(items_table).on_conflict_do_update(
                index_elements=["id"],
                set_={"age": 77},
            )
            conn.execute(
                stmt,
                [
                    {"id": 1, "name": "Ignore1", "age": 999, "code": "X"},
                    {"id": 2, "name": "Ignore2", "age": 999, "code": "Y"},
                    {"id": 3, "name": "Cara", "age": 26, "code": "C"},
                ],
            )

        with populated_engine.connect() as conn:
            rows = conn.execute(select(items_table).order_by(items_table.c.id)).all()
            assert rows == [
                (1, "Alice", 77, "A"),
                (2, "Bob", 77, "B"),
                (3, "Cara", 26, "C"),
            ]


def test_double_on_conflict_raises(items_table):
    stmt = insert(items_table).values(id=1, name="Alice", age=30, code="A")
    stmt = stmt.on_conflict_do_nothing(index_elements=["id"])

    with pytest.raises(InvalidRequestError, match="already has an ON CONFLICT clause"):
        stmt.on_conflict_do_update(index_elements=["id"], set_={"age": 40})


def test_where_rejected(items_table):
    stmt = insert(items_table).values(id=1, name="Alice", age=30, code="A")

    with pytest.raises(ValueError, match="does not support WHERE clause"):
        stmt.on_conflict_do_update(
            index_elements=["id"],
            set_={"age": 1},
            where=items_table.c.age > 10,
        )


def test_index_where_rejected(items_table):
    stmt = insert(items_table).values(id=1, name="Alice", age=30, code="A")

    with pytest.raises(ValueError, match="does not support index_where"):
        stmt.on_conflict_do_nothing(
            index_elements=["id"],
            index_where=items_table.c.id > 0,
        )


def test_missing_index_elements_do_nothing(items_table):
    stmt = insert(items_table).values(id=1, name="Alice", age=30, code="A")

    with pytest.raises(ValueError, match="requires index_elements"):
        stmt.on_conflict_do_nothing()


def test_missing_index_elements_do_update(items_table):
    stmt = insert(items_table).values(id=1, name="Alice", age=30, code="A")

    with pytest.raises(ValueError, match="requires index_elements"):
        stmt.on_conflict_do_update(set_={"age": 1})


def test_index_where_rejected_do_update(items_table):
    """index_where must be rejected on do_update too, not just do_nothing."""
    stmt = insert(items_table).values(id=1, name="Alice", age=30, code="A")

    with pytest.raises(ValueError, match="does not support index_where"):
        stmt.on_conflict_do_update(
            index_elements=["id"],
            set_={"age": 1},
            index_where=items_table.c.id > 0,
        )


def test_foreign_table_column_key_rejected(items_table, composite_items_table):
    """set_ keys from a different table must be rejected at compile time."""
    stmt = insert(items_table).values(id=1, name="Alice", age=30, code="A")
    upsert = stmt.on_conflict_do_update(
        index_elements=["id"],
        set_={composite_items_table.c.age: 99},
    )
    from sqlalchemy.exc import CompileError
    from sqlalchemy_excel.dialect import ExcelDialect

    with pytest.raises(CompileError, match="cannot be used as an ON CONFLICT"):
        upsert.compile(dialect=ExcelDialect())
