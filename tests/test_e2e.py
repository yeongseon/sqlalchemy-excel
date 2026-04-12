from __future__ import annotations

import pytest
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


def test_e2e_join_is_rejected(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
    )

    stmt = select(users).join(orders, users.c.id == orders.c.user_id)
    with pytest.raises(exc.CompileError, match="JOIN"):
        stmt.compile(dialect=engine.dialect)

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
