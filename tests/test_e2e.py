from __future__ import annotations

import pytest
import sqlalchemy as sa
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
    create_engine,
    delete,
    exc,
    func,
    insert,
    inspect,
    select,
    true,
    update,
)
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column


class Base(DeclarativeBase):
    pass


class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String)
    age: Mapped[int] = mapped_column(Integer)


def _engine_for(tmp_path):
    return create_engine(f"excel:///{tmp_path / 'test.xlsx'}")


def _users_table(metadata: MetaData) -> Table:
    return Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("age", Integer),
    )


def test_e2e_core_crud_round_trip(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))

    with engine.connect() as conn:
        rows = conn.execute(select(users).order_by(users.c.id)).all()
        assert rows == [(1, "Alice", 30)]

    with engine.begin() as conn:
        conn.execute(update(users).where(users.c.id == 1).values(age=31))

    with engine.connect() as conn:
        updated = conn.execute(select(users.c.age).where(users.c.id == 1)).scalar_one()
        assert updated == 31

    with engine.begin() as conn:
        conn.execute(delete(users).where(users.c.id == 1))

    with engine.connect() as conn:
        remaining = conn.execute(select(users)).all()
        assert remaining == []

    engine.dispose()


def test_multi_row_insert_e2e(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Charlie", "age": 35},
            ],
        )

    with engine.connect() as conn:
        rows = conn.execute(select(users).order_by(users.c.id)).all()
        assert rows == [(1, "Alice", 30), (2, "Bob", 25), (3, "Charlie", 35)]

    engine.dispose()


def test_insert_from_select_e2e(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    source = Table(
        "source",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    target = Table(
        "target",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(source), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
        conn.execute(
            target.insert().from_select(["id", "name"], select(source.c.id, source.c.name))
        )

    with engine.connect() as conn:
        rows = conn.execute(select(target).order_by(target.c.id)).all()
        assert rows == [(1, "Alice"), (2, "Bob")]

    engine.dispose()


def test_insert_from_select_with_where_e2e(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    source = Table(
        "source",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    target = Table(
        "target",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(source),
            [
                {"id": 1, "name": "Alice"},
                {"id": 2, "name": "Bob"},
                {"id": 3, "name": "Charlie"},
            ],
        )
        conn.execute(
            target.insert().from_select(
                ["id", "name"],
                select(source.c.id, source.c.name).where(source.c.id >= 2),
            )
        )

    with engine.connect() as conn:
        rows = conn.execute(select(target).order_by(target.c.id)).all()
        assert rows == [(2, "Bob"), (3, "Charlie")]

    engine.dispose()


def test_e2e_session_add_commit_and_readback(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    Base.metadata.create_all(engine)

    with Session(engine) as session:
        session.add(User(id=1, name="Alice", age=30))
        session.commit()

    with Session(engine) as session:
        user = session.execute(select(User).where(User.id == 1)).scalar_one()
        assert user.name == "Alice"
        assert user.age == 30

    engine.dispose()


def test_e2e_session_filtered_ordered_limited_select(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    Base.metadata.create_all(engine)

    with Session(engine) as session:
        session.add_all(
            [
                User(id=1, name="Alice", age=30),
                User(id=2, name="Bob", age=22),
                User(id=3, name="Charlie", age=35),
            ]
        )
        session.commit()

    with Session(engine) as session:
        stmt = select(User).where(User.age >= 25).order_by(User.age.desc()).limit(2)
        names = [user.name for user in session.execute(stmt).scalars().all()]
        assert names == ["Charlie", "Alice"]

    engine.dispose()


def test_e2e_inspector_reflection(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    _users_table(metadata)
    metadata.create_all(engine)

    insp = inspect(engine)
    assert "users" in insp.get_table_names()
    assert insp.has_table("users")
    columns = insp.get_columns("users")
    names = [column["name"] for column in columns]
    assert names == ["id", "name", "age"]

    engine.dispose()


def test_e2e_create_and_drop_all_lifecycle(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    _users_table(metadata)

    metadata.create_all(engine)
    assert inspect(engine).has_table("users")

    metadata.drop_all(engine)
    assert not inspect(engine).has_table("users")

    engine.dispose()


def test_e2e_rollback_is_silent_noop_and_data_persists(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.rollback()

    with engine.connect() as conn:
        rows = conn.execute(select(users)).all()
        assert rows == [(1, "Alice", 30)]

    engine.dispose()


def test_e2e_join_inner(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Charlie", age=35))
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))
        conn.execute(insert(orders).values(id=2, user_id=1, amount=200))
        conn.execute(insert(orders).values(id=3, user_id=3, amount=300))

    with engine.connect() as conn:
        stmt = (
            select(users.c.name, orders.c.amount)
            .join(orders, users.c.id == orders.c.user_id)
            .order_by(orders.c.amount)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 100), ("Alice", 200), ("Charlie", 300)]

    engine.dispose()


def test_e2e_join_left(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))

    with engine.connect() as conn:
        stmt = (
            select(users.c.name, orders.c.amount)
            .join(orders, users.c.id == orders.c.user_id, isouter=True)
            .order_by(users.c.name)
        )
        rows = conn.execute(stmt).all()
        # Bob has no orders, so LEFT JOIN returns (Bob, None)
        assert rows == [("Alice", 100), ("Bob", None)]

    engine.dispose()


def test_e2e_subquery_in_where(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    admins = Table(
        "admins",
        metadata,
        Column("id", Integer, primary_key=True),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Charlie", age=35))
        conn.execute(insert(admins).values(id=1))
        conn.execute(insert(admins).values(id=3))

    with engine.connect() as conn:
        stmt = select(users.c.id, users.c.name).where(
            users.c.id.in_(select(admins.c.id))
        )
        rows = conn.execute(stmt).all()
        assert rows == [(1, "Alice"), (3, "Charlie")]

    engine.dispose()


def test_e2e_subquery_with_inner_where(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    admins = Table(
        "admins",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("role", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Charlie", age=35))
        conn.execute(insert(admins).values(id=1, role="admin"))
        conn.execute(insert(admins).values(id=3, role="editor"))

    with engine.connect() as conn:
        stmt = select(users.c.id, users.c.name).where(
            users.c.id.in_(select(admins.c.id).where(admins.c.role == "admin"))
        )
        rows = conn.execute(stmt).all()
        assert rows == [(1, "Alice")]

    engine.dispose()


def test_e2e_aggregate_count(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Alice", age=35))

    with engine.connect() as conn:
        stmt = select(func.count(users.c.id)).select_from(users)
        result = conn.execute(stmt).scalar_one()
        assert result == 3

    engine.dispose()


def test_e2e_aggregate_sum_avg_min_max(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=20))
        conn.execute(insert(users).values(id=3, name="Charlie", age=40))

    with engine.connect() as conn:
        stmt = select(func.sum(users.c.age)).select_from(users)
        assert conn.execute(stmt).scalar_one() == 90

        stmt = select(func.avg(users.c.age)).select_from(users)
        assert conn.execute(stmt).scalar_one() == 30.0

        stmt = select(func.min(users.c.age)).select_from(users)
        assert conn.execute(stmt).scalar_one() == 20

        stmt = select(func.max(users.c.age)).select_from(users)
        assert conn.execute(stmt).scalar_one() == 40

    engine.dispose()


def test_e2e_where_aggregate_rejected(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))

    with engine.connect() as conn:
        stmt = select(users.c.name).where(func.count(users.c.id) > 1)
        with pytest.raises(
            (exc.DBAPIError, ValueError),
            match="Aggregate functions are not allowed in WHERE clause; use HAVING instead",
        ):
            conn.execute(stmt).all()

    engine.dispose()


def test_e2e_group_by(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Alice", age=35))

    with engine.connect() as conn:
        stmt = (
            select(users.c.name, func.count(users.c.id))
            .group_by(users.c.name)
            .order_by(users.c.name)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 2), ("Bob", 1)]

    engine.dispose()


def test_e2e_group_by_having(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Alice", age=35))

    with engine.connect() as conn:
        stmt = (
            select(users.c.name, func.count(users.c.id))
            .group_by(users.c.name)
            .having(func.count(users.c.id) > 1)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 2)]

    engine.dispose()


def test_e2e_group_by_having_aggregate_not_in_select(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Alice", age=35))

    with engine.connect() as conn:
        stmt = (
            select(users.c.name)
            .group_by(users.c.name)
            .having(func.count(users.c.id) > 1)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice",)]

    engine.dispose()


def test_e2e_group_by_order_by_group_key_not_in_select(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Alice", age=35))

    with engine.connect() as conn:
        stmt = (
            select(func.count(users.c.id))
            .group_by(users.c.name)
            .order_by(users.c.name)
        )
        rows = conn.execute(stmt).all()
        assert rows == [(2,), (1,)]

    engine.dispose()


def test_e2e_rejects_aggregate_arithmetic_projection(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))

    with engine.connect() as conn:
        stmt = select(func.sum(users.c.age) + 1).select_from(users)
        with pytest.raises(
            (exc.DBAPIError, ValueError),
            match="Unsupported column expression",
        ):
            conn.execute(stmt).all()

    engine.dispose()

def test_e2e_offset_compiles_and_executes(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(users).values(id=3, name="Charlie", age=35))

    with engine.connect() as conn:
        rows = conn.execute(
            select(users).order_by(users.c.id).limit(2).offset(1)
        ).all()
        assert rows == [(2, "Bob", 25), (3, "Charlie", 35)]

    engine.dispose()


def test_e2e_distinct_compiles_and_executes(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Alice", age=25))
        conn.execute(insert(users).values(id=3, name="Bob", age=30))

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name).distinct()
        ).all()
        names = [r[0] for r in rows]
        assert sorted(names) == ["Alice", "Bob"]

    engine.dispose()


def test_union_e2e(tmp_path) -> None:
    db = tmp_path / "test.xlsx"
    engine = create_engine(f"excel:///{db}")
    metadata = MetaData()
    sheet1 = Table(
        "Sheet1",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    sheet2 = Table(
        "Sheet2",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
        conn.execute(insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}])

    t1 = Table("Sheet1", MetaData(), autoload_with=engine)
    t2 = Table("Sheet2", MetaData(), autoload_with=engine)
    stmt = sa.union(select(t1.c.id, t1.c.name), select(t2.c.id, t2.c.name))
    with engine.connect() as conn:
        rows = conn.execute(stmt).all()

    assert sorted(rows) == [(1, "Alice"), (2, "Bob"), (3, "Cara")]
    engine.dispose()


def test_union_all_e2e(tmp_path) -> None:
    db = tmp_path / "test.xlsx"
    engine = create_engine(f"excel:///{db}")
    metadata = MetaData()
    sheet1 = Table(
        "Sheet1",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    sheet2 = Table(
        "Sheet2",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
        conn.execute(insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}])

    t1 = Table("Sheet1", MetaData(), autoload_with=engine)
    t2 = Table("Sheet2", MetaData(), autoload_with=engine)
    stmt = sa.union_all(select(t1.c.id, t1.c.name), select(t2.c.id, t2.c.name))
    with engine.connect() as conn:
        rows = conn.execute(stmt).all()

    assert sorted(rows) == [(1, "Alice"), (2, "Bob"), (2, "Bob"), (3, "Cara")]
    engine.dispose()


def test_intersect_e2e(tmp_path) -> None:
    db = tmp_path / "test.xlsx"
    engine = create_engine(f"excel:///{db}")
    metadata = MetaData()
    sheet1 = Table(
        "Sheet1",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    sheet2 = Table(
        "Sheet2",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
        conn.execute(insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}])

    t1 = Table("Sheet1", MetaData(), autoload_with=engine)
    t2 = Table("Sheet2", MetaData(), autoload_with=engine)
    stmt = sa.intersect(select(t1.c.id, t1.c.name), select(t2.c.id, t2.c.name))
    with engine.connect() as conn:
        rows = conn.execute(stmt).all()

    assert rows == [(2, "Bob")]
    engine.dispose()


def test_except_e2e(tmp_path) -> None:
    db = tmp_path / "test.xlsx"
    engine = create_engine(f"excel:///{db}")
    metadata = MetaData()
    sheet1 = Table(
        "Sheet1",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    sheet2 = Table(
        "Sheet2",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}])
        conn.execute(insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}])

    t1 = Table("Sheet1", MetaData(), autoload_with=engine)
    t2 = Table("Sheet2", MetaData(), autoload_with=engine)
    stmt = sa.except_(select(t1.c.id, t1.c.name), select(t2.c.id, t2.c.name))
    with engine.connect() as conn:
        rows = conn.execute(stmt).all()

    assert rows == [(1, "Alice")]
    engine.dispose()


def test_e2e_nested_subquery_rejected(tmp_path) -> None:
    """Nested subquery should fail at compile time, not just at DB-API level."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    admins = Table(
        "admins",
        metadata,
        Column("id", Integer, primary_key=True),
    )
    items = Table(
        "items",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("admin_id", Integer),
    )
    metadata.create_all(engine)

    inner = select(items.c.admin_id).where(items.c.id > 0)
    outer = select(admins.c.id).where(admins.c.id.in_(inner))
    stmt = select(users).where(users.c.id.in_(outer))
    with pytest.raises(exc.CompileError, match="nested subqueries"):
        with engine.connect() as conn:
            conn.execute(stmt).all()

    engine.dispose()


def test_e2e_join_with_as_alias(tmp_path) -> None:
    """JOIN with AS alias syntax (SQLAlchemy's default) works end-to-end."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))
        conn.execute(insert(orders).values(id=2, user_id=2, amount=200))

    with engine.connect() as conn:
        # SQLAlchemy emits 'FROM users AS users_1 JOIN orders AS orders_1 ON ...'
        # The parser must accept AS aliases
        stmt = (
            select(users.c.name, orders.c.amount)
            .join(orders, users.c.id == orders.c.user_id)
            .order_by(orders.c.amount)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 100), ("Bob", 200)]

    engine.dispose()


def test_e2e_chained_join_three_tables(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
    )
    items = Table(
        "items",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("order_id", Integer),
        Column("sku", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(orders).values(id=10, user_id=1, amount=100))
        conn.execute(insert(orders).values(id=11, user_id=2, amount=200))
        conn.execute(insert(items).values(id=100, order_id=10, sku="A-1"))
        conn.execute(insert(items).values(id=101, order_id=11, sku="B-1"))

    with engine.connect() as conn:
        stmt = (
            select(users.c.name, orders.c.amount, items.c.sku)
            .join(orders, users.c.id == orders.c.user_id)
            .join(items, orders.c.id == items.c.order_id)
            .where(orders.c.amount >= 100)
            .order_by(orders.c.amount)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 100, "A-1"), ("Bob", 200, "B-1")]

    engine.dispose()


def test_e2e_right_join_via_swapped_left_join(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))
        conn.execute(insert(orders).values(id=2, user_id=999, amount=999))

    with engine.connect() as conn:
        stmt = (
            select(orders.c.id, users.c.name)
            .select_from(orders.join(users, users.c.id == orders.c.user_id, isouter=True))
            .order_by(orders.c.id)
        )
        rows = conn.execute(stmt).all()
        assert rows == [(1, "Alice"), (2, None)]

    engine.dispose()


def test_e2e_full_outer_join_rejected_at_compile_time(tmp_path) -> None:
    """FULL OUTER JOIN raises CompileError, not a later DBAPI error."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
    )
    metadata.create_all(engine)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, users.c.id == orders.c.user_id, full=True)
    )
    with pytest.raises(exc.CompileError, match="FULL OUTER JOIN"):
        with engine.connect() as conn:
            conn.execute(stmt).all()

    engine.dispose()


def test_e2e_non_equality_on_clause_rejected_at_compile_time(tmp_path) -> None:
    """Non-equality ON clause (e.g. true()) raises CompileError."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
    )
    metadata.create_all(engine)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, true())
    )
    with pytest.raises(exc.CompileError, match="equality comparisons"):
        with engine.connect() as conn:
            conn.execute(stmt).all()

    engine.dispose()


def test_e2e_or_on_clause_rejected_at_compile_time(tmp_path) -> None:
    """OR-combined ON clause raises CompileError."""
    from sqlalchemy import or_

    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("dept_id", Integer),
    )
    metadata.create_all(engine)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(
            orders,
            or_(users.c.id == orders.c.user_id, users.c.id == orders.c.dept_id),
        )
    )
    with pytest.raises(exc.CompileError, match="OR"):
        with engine.connect() as conn:
            conn.execute(stmt).all()

    engine.dispose()


def test_e2e_literal_on_clause_rejected_at_compile_time(tmp_path) -> None:
    """ON clause with literal operand (col == 1) raises CompileError."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
    )
    metadata.create_all(engine)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, users.c.id == 1)
    )
    with pytest.raises(exc.CompileError, match="column references"):
        with engine.connect() as conn:
            conn.execute(stmt).all()

    engine.dispose()


def test_e2e_same_side_on_clause_rejected_at_compile_time(tmp_path) -> None:
    """ON clause with same-side columns (users.id == users.age) raises CompileError."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = Table(
        "users",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("age", Integer),
    )
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
    )
    metadata.create_all(engine)

    stmt = (
        select(users.c.name, orders.c.user_id)
        .join(orders, users.c.id == users.c.age)
    )
    with pytest.raises(exc.CompileError, match="different join sources"):
        with engine.connect() as conn:
            conn.execute(stmt).all()

    engine.dispose()


def test_e2e_compound_with_mixed_join_branches(tmp_path) -> None:
    """UNION where branch 1 has a JOIN and branch 2 does not.

    Regression test: _has_join must reset between compound branches.
    """
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.execute(insert(users).values(id=2, name="Bob", age=25))
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))

    with engine.connect() as conn:
        # Branch 1: JOIN (users + orders), branch 2: single-table (users)
        stmt = sa.union(
            select(users.c.id)
            .join(orders, users.c.id == orders.c.user_id),
            select(users.c.id),
        )
        rows = conn.execute(stmt).all()
        # UNION deduplication: ids 1 (from join) and 1,2 (from users) → {1, 2}
        assert sorted(rows) == [(1,), (2,)]

    engine.dispose()


def test_e2e_compound_with_outer_order_by(tmp_path) -> None:
    """Compound-level ORDER BY sorts the entire result set."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Charlie", "age": 35},
            ],
        )

    t = Table("users", MetaData(), autoload_with=engine)
    with engine.connect() as conn:
        # UNION ALL with compound-level ORDER BY DESC
        stmt = (
            sa.union_all(
                select(t.c.id).where(t.c.id <= 2),
                select(t.c.id).where(t.c.id >= 2),
            )
            .order_by(t.c.id.desc())
        )
        rows = conn.execute(stmt).all()
        # UNION ALL keeps dupes: [1, 2] + [2, 3] = [1, 2, 2, 3], ORDER BY DESC → [3, 2, 2, 1]
        assert rows == [(3,), (2,), (2,), (1,)]

    engine.dispose()


def test_e2e_compound_with_outer_order_by_and_limit(tmp_path) -> None:
    """Compound-level ORDER BY + LIMIT."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Charlie", "age": 35},
            ],
        )

    t = Table("users", MetaData(), autoload_with=engine)
    with engine.connect() as conn:
        stmt = (
            sa.union_all(
                select(t.c.id).where(t.c.id <= 2),
                select(t.c.id).where(t.c.id >= 2),
            )
            .order_by(t.c.id.desc())
            .limit(2)
        )
        rows = conn.execute(stmt).all()
        # [3, 2, 2, 1] limited to 2 → [3, 2]
        assert rows == [(3,), (2,)]

    engine.dispose()
