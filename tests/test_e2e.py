from __future__ import annotations

import pytest
import sqlalchemy as sa
from sqlalchemy import (
    Column,
    Integer,
    MetaData,
    String,
    Table,
    case,
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


def _employees_table(metadata: MetaData) -> Table:
    return Table(
        "employees",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("team", String),
        Column("dept", String),
    )


def _arithmetic_table(metadata: MetaData) -> Table:
    return Table(
        "arith",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("price", Integer),
        Column("qty", Integer),
        Column("tax", Integer),
    )


def _seed_arithmetic_data(engine: sa.engine.Engine, table: Table) -> None:
    with engine.begin() as conn:
        conn.execute(
            insert(table),
            [
                {"id": 1, "price": 100, "qty": 5, "tax": 10},
                {"id": 2, "price": 200, "qty": 3, "tax": 20},
                {"id": 3, "price": 300, "qty": 0, "tax": 30},
            ],
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
        conn.execute(
            insert(source), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]
        )
        conn.execute(
            target.insert().from_select(
                ["id", "name"], select(source.c.id, source.c.name)
            )
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


def test_e2e_rollback_reverts_uncommitted_insert(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(insert(users).values(id=1, name="Alice", age=30))
        conn.rollback()

    with engine.connect() as conn:
        rows = conn.execute(select(users)).all()
        assert rows == []

    engine.dispose()


def test_e2e_rollback_reverts_insert_update_delete(tmp_path) -> None:
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
            ],
        )

    with engine.connect() as conn:
        conn.execute(insert(users).values(id=3, name="Charlie", age=40))
        conn.execute(update(users).where(users.c.id == 1).values(age=99))
        conn.execute(delete(users).where(users.c.id == 2))
        conn.rollback()

    with engine.connect() as conn:
        rows = conn.execute(select(users).order_by(users.c.id)).all()
        assert rows == [(1, "Alice", 30), (2, "Bob", 25)]

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


def test_e2e_right_join_raw_sql(tmp_path) -> None:
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
        rows = conn.execute(
            sa.text(
                "SELECT users.name, orders.amount "
                "FROM users RIGHT JOIN orders ON users.id = orders.user_id"
            )
        ).all()
        assert rows == [("Alice", 100)]

    engine.dispose()


def test_e2e_full_outer_join_raw_sql(tmp_path) -> None:
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
        conn.execute(insert(orders).values(id=2, user_id=999, amount=200))

    with engine.connect() as conn:
        rows = conn.execute(
            sa.text(
                "SELECT users.name, orders.amount "
                "FROM users FULL OUTER JOIN orders ON users.id = orders.user_id"
            )
        ).all()
        assert set(rows) == {("Alice", 100), ("Bob", None), (None, 200)}

    engine.dispose()


def test_e2e_cross_join_raw_sql(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    teams = Table(
        "teams",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.execute(insert(teams), [{"id": 1, "name": "A"}, {"id": 2, "name": "B"}])

    with engine.connect() as conn:
        rows = conn.execute(
            sa.text(
                "SELECT users.id, teams.id FROM users CROSS JOIN teams ORDER BY users.id, teams.id"
            )
        ).all()
        assert rows == [(1, 1), (1, 2), (2, 1), (2, 2)]

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


def test_e2e_update_with_subquery_in_where(tmp_path) -> None:
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
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Charlie", "age": 35},
            ],
        )
        conn.execute(insert(admins), [{"id": 1}, {"id": 3}])

    with engine.begin() as conn:
        stmt = users.update().where(users.c.id.in_(select(admins.c.id))).values(age=99)
        result = conn.execute(stmt)
        assert result.rowcount == 2

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name, users.c.age).order_by(users.c.id)
        ).all()
        assert rows == [("Alice", 99), ("Bob", 25), ("Charlie", 99)]

    engine.dispose()


def test_e2e_delete_with_subquery_in_where(tmp_path) -> None:
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
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Charlie", "age": 35},
            ],
        )
        conn.execute(insert(admins), [{"id": 1}, {"id": 3}])

    with engine.begin() as conn:
        stmt = users.delete().where(users.c.id.notin_(select(admins.c.id)))
        result = conn.execute(stmt)
        assert result.rowcount == 1

    with engine.connect() as conn:
        rows = conn.execute(select(users.c.name).order_by(users.c.id)).all()
        assert rows == [("Alice",), ("Charlie",)]

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


def test_count_distinct_basic(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    employees = _employees_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(employees),
            [
                {"id": 1, "team": "A", "dept": "Sales"},
                {"id": 2, "team": "A", "dept": "Sales"},
                {"id": 3, "team": "A", "dept": "Support"},
                {"id": 4, "team": "B", "dept": None},
            ],
        )

    with engine.connect() as conn:
        stmt = select(sa.func.count(sa.distinct(employees.c.dept))).select_from(
            employees
        )
        unique_depts = conn.execute(stmt).scalar_one()
        assert unique_depts == 2

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


def test_e2e_column_alias(tmp_path) -> None:
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
            ],
        )

    with engine.connect() as conn:
        stmt = select(users.c.name.label("n"), users.c.age.label("a"))
        result = conn.execute(stmt)
        keys = list(result.keys())
        assert "n" in keys
        assert "a" in keys
        rows = result.all()
        assert len(rows) == 2

    engine.dispose()


def test_e2e_arithmetic_multiplication(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select((arith.c.price * arith.c.qty).label("total")).order_by(
            arith.c.id
        )
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(500,), (600,), (0,)]

    engine.dispose()


def test_e2e_arithmetic_addition(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select((arith.c.price + arith.c.tax).label("sum")).order_by(
            arith.c.id
        )
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(110,), (220,), (330,)]

    engine.dispose()


def test_e2e_arithmetic_subtraction(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select((arith.c.price - arith.c.tax).label("diff")).order_by(
            arith.c.id
        )
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(90,), (180,), (270,)]

    engine.dispose()


def test_e2e_arithmetic_complex_expression(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select(
            ((arith.c.price + arith.c.tax) * arith.c.qty).label("total")
        ).order_by(arith.c.id)
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(550,), (660,), (0,)]

    engine.dispose()


def test_e2e_arithmetic_unary_negation(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select((-arith.c.price).label("neg")).order_by(arith.c.id)
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(-100,), (-200,), (-300,)]

    engine.dispose()


def test_e2e_arithmetic_with_alias(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select((arith.c.price * arith.c.qty).label("custom_name")).order_by(
            arith.c.id
        )
        result = conn.execute(stmt)
        cursor = result.cursor
        assert cursor is not None
        assert cursor.description is not None
        assert [desc[0] for desc in cursor.description] == ["custom_name"]
        rows = result.fetchall()
        assert rows == [(500,), (600,), (0,)]

    engine.dispose()


def test_e2e_arithmetic_with_where(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = (
            sa.select((arith.c.price * arith.c.qty).label("total"))
            .where(arith.c.qty > 0)
            .order_by(arith.c.id)
        )
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(500,), (600,)]

    engine.dispose()


def test_e2e_arithmetic_null_propagation(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.begin() as conn:
        conn.execute(insert(arith).values(id=4, price=None, qty=2, tax=10))

    with engine.connect() as conn:
        stmt = sa.select((arith.c.price * arith.c.qty).label("total")).where(
            arith.c.id == 4
        )
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(None,)]

    engine.dispose()


def test_e2e_arithmetic_with_literal(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select(
            (arith.c.price * sa.literal_column("2")).label("doubled")
        ).order_by(arith.c.id)
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(200,), (400,), (600,)]

    engine.dispose()


def test_e2e_arithmetic_order_by_alias(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    arith = _arithmetic_table(metadata)
    metadata.create_all(engine)
    _seed_arithmetic_data(engine, arith)

    with engine.connect() as conn:
        stmt = sa.select((arith.c.price * arith.c.qty).label("total")).order_by(
            sa.text("total")
        )
        result = conn.execute(stmt)
        rows = result.fetchall()
        assert rows == [(0,), (500,), (600,)]

    engine.dispose()


def test_e2e_aggregate_alias(tmp_path) -> None:
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
            ],
        )

    with engine.connect() as conn:
        stmt = select(func.count(users.c.id).label("total")).select_from(users)
        result = conn.execute(stmt)
        row = result.one()
        assert row[0] == 2

    engine.dispose()


def test_count_distinct_with_alias(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    employees = _employees_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(employees),
            [
                {"id": 1, "team": "A", "dept": "Sales"},
                {"id": 2, "team": "A", "dept": "Sales"},
                {"id": 3, "team": "A", "dept": "Support"},
            ],
        )

    with engine.connect() as conn:
        stmt = select(
            sa.func.count(sa.distinct(employees.c.dept)).label("unique_depts")
        ).select_from(employees)
        result = conn.execute(stmt)
        assert list(result.keys()) == ["unique_depts"]
        row = result.one()
        assert row[0] == 2

    engine.dispose()


def test_e2e_order_by_alias(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Charlie", "age": 30},
                {"id": 2, "name": "Alice", "age": 25},
                {"id": 3, "name": "Bob", "age": 35},
            ],
        )

    with engine.connect() as conn:
        label_col = users.c.name.label("n")
        stmt = select(label_col).order_by(label_col)
        result = conn.execute(stmt)
        names = [row[0] for row in result]
        assert names == ["Alice", "Bob", "Charlie"]

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


def test_count_distinct_with_group_by(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    employees = _employees_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(employees),
            [
                {"id": 1, "team": "A", "dept": "Sales"},
                {"id": 2, "team": "A", "dept": "Sales"},
                {"id": 3, "team": "A", "dept": "Support"},
                {"id": 4, "team": "B", "dept": "Finance"},
                {"id": 5, "team": "B", "dept": "Finance"},
                {"id": 6, "team": "B", "dept": None},
            ],
        )

    with engine.connect() as conn:
        stmt = (
            select(employees.c.team, sa.func.count(sa.distinct(employees.c.dept)))
            .group_by(employees.c.team)
            .order_by(employees.c.team)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("A", 2), ("B", 1)]

    engine.dispose()


def test_count_distinct_having(tmp_path) -> None:
    """HAVING with COUNT(DISTINCT col) filter."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    employees = _employees_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(employees),
            [
                {"id": 1, "team": "A", "dept": "Sales"},
                {"id": 2, "team": "A", "dept": "Support"},
                {"id": 3, "team": "A", "dept": "Finance"},
                {"id": 4, "team": "B", "dept": "Sales"},
                {"id": 5, "team": "B", "dept": "Sales"},
                {"id": 6, "team": "C", "dept": "Sales"},
                {"id": 7, "team": "C", "dept": "Support"},
            ],
        )

    with engine.connect() as conn:
        stmt = (
            select(employees.c.team, sa.func.count(sa.distinct(employees.c.dept)))
            .group_by(employees.c.team)
            .having(sa.func.count(sa.distinct(employees.c.dept)) > 1)
            .order_by(employees.c.team)
        )
        rows = conn.execute(stmt).all()
        # A: 3 distinct depts (>1), B: 1 distinct dept (not >1), C: 2 distinct depts (>1)
        assert rows == [("A", 3), ("C", 2)]

    engine.dispose()


def test_count_distinct_order_by(tmp_path) -> None:
    """ORDER BY COUNT(DISTINCT col)."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    employees = _employees_table(metadata)
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(employees),
            [
                {"id": 1, "team": "A", "dept": "Sales"},
                {"id": 2, "team": "A", "dept": "Support"},
                {"id": 3, "team": "A", "dept": "Finance"},
                {"id": 4, "team": "B", "dept": "Sales"},
                {"id": 5, "team": "B", "dept": "Sales"},
                {"id": 6, "team": "C", "dept": "Sales"},
                {"id": 7, "team": "C", "dept": "Support"},
            ],
        )

    with engine.connect() as conn:
        cnt_distinct = sa.func.count(sa.distinct(employees.c.dept))
        stmt = (
            select(employees.c.team, cnt_distinct)
            .group_by(employees.c.team)
            .order_by(cnt_distinct.desc())
        )
        rows = conn.execute(stmt).all()
        # A: 3 distinct depts, C: 2 distinct depts, B: 1 distinct dept
        assert rows == [("A", 3), ("C", 2), ("B", 1)]

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
            select(func.count(users.c.id)).group_by(users.c.name).order_by(users.c.name)
        )
        rows = conn.execute(stmt).all()
        assert rows == [(2,), (1,)]

    engine.dispose()


def test_e2e_group_by_with_join(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
        Column("status", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Cara", "age": 40},
            ],
        )
        conn.execute(
            insert(orders),
            [
                {"id": 10, "user_id": 1, "amount": 100, "status": "paid"},
                {"id": 11, "user_id": 1, "amount": 150, "status": "paid"},
                {"id": 12, "user_id": 2, "amount": 50, "status": "pending"},
            ],
        )

    with engine.connect() as conn:
        stmt = (
            sa.select(users.c.name, sa.func.count())
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .group_by(users.c.name)
            .order_by(users.c.name)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 2), ("Bob", 1)]

    engine.dispose()


def test_e2e_group_by_join_with_having(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
        Column("status", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.execute(
            insert(orders),
            [
                {"id": 10, "user_id": 1, "amount": 100, "status": "paid"},
                {"id": 11, "user_id": 1, "amount": 150, "status": "paid"},
                {"id": 12, "user_id": 2, "amount": 50, "status": "pending"},
            ],
        )

    with engine.connect() as conn:
        order_count = sa.func.count()
        stmt = (
            sa.select(users.c.name, order_count)
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .group_by(users.c.name)
            .having(order_count > 1)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 2)]

    engine.dispose()


def test_e2e_group_by_join_with_sum(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
        Column("status", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.execute(
            insert(orders),
            [
                {"id": 10, "user_id": 1, "amount": 100, "status": "paid"},
                {"id": 11, "user_id": 1, "amount": 150, "status": "paid"},
                {"id": 12, "user_id": 2, "amount": 50, "status": "pending"},
            ],
        )

    with engine.connect() as conn:
        total_amount = sa.func.sum(orders.c.amount)
        stmt = (
            sa.select(users.c.name, total_amount)
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .group_by(users.c.name)
            .order_by(users.c.name)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 250), ("Bob", 50)]

    engine.dispose()


def test_e2e_group_by_join_order_by_aggregate(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
        Column("status", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Cara", "age": 40},
            ],
        )
        conn.execute(
            insert(orders),
            [
                {"id": 10, "user_id": 1, "amount": 100, "status": "paid"},
                {"id": 11, "user_id": 1, "amount": 150, "status": "paid"},
                {"id": 12, "user_id": 2, "amount": 50, "status": "pending"},
                {"id": 13, "user_id": 3, "amount": 10, "status": "paid"},
                {"id": 14, "user_id": 3, "amount": 20, "status": "pending"},
                {"id": 15, "user_id": 3, "amount": 30, "status": "failed"},
            ],
        )

    with engine.connect() as conn:
        order_count = sa.func.count()
        stmt = (
            sa.select(users.c.name, order_count)
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .group_by(users.c.name)
            .order_by(order_count.desc(), users.c.name.asc())
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Cara", 3), ("Alice", 2), ("Bob", 1)]

    engine.dispose()


def test_e2e_group_by_join_multiple_columns(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    orders = Table(
        "orders",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("user_id", Integer),
        Column("amount", Integer),
        Column("status", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.execute(
            insert(orders),
            [
                {"id": 10, "user_id": 1, "amount": 100, "status": "paid"},
                {"id": 11, "user_id": 1, "amount": 150, "status": "pending"},
                {"id": 12, "user_id": 2, "amount": 50, "status": "paid"},
                {"id": 13, "user_id": 2, "amount": 75, "status": "pending"},
            ],
        )

    with engine.connect() as conn:
        stmt = (
            sa.select(users.c.name, orders.c.status, sa.func.count())
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .group_by(users.c.name, orders.c.status)
            .order_by(users.c.name, orders.c.status)
        )
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "paid", 1),
            ("Alice", "pending", 1),
            ("Bob", "paid", 1),
            ("Bob", "pending", 1),
        ]

    engine.dispose()


def test_e2e_group_by_join_explicit_alias(tmp_path) -> None:
    """Regression: GROUP BY + JOIN with explicit SQLAlchemy aliases."""
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
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.execute(
            insert(orders),
            [
                {"id": 10, "user_id": 1, "amount": 100},
                {"id": 11, "user_id": 1, "amount": 150},
                {"id": 12, "user_id": 2, "amount": 50},
            ],
        )

    u = users.alias("u")
    o = orders.alias("o")

    with engine.connect() as conn:
        stmt = (
            sa.select(u.c.name, sa.func.sum(o.c.amount).label("total"))
            .select_from(u.join(o, u.c.id == o.c.user_id))
            .group_by(u.c.name)
            .order_by(u.c.name)
        )
        rows = conn.execute(stmt).all()
        assert rows == [("Alice", 250), ("Bob", 50)]

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
        rows = conn.execute(select(users).order_by(users.c.id).limit(2).offset(1)).all()
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
        rows = conn.execute(select(users.c.name).distinct()).all()
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
        conn.execute(
            insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]
        )
        conn.execute(
            insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}]
        )

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
        conn.execute(
            insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]
        )
        conn.execute(
            insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}]
        )

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
        conn.execute(
            insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]
        )
        conn.execute(
            insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}]
        )

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
        conn.execute(
            insert(sheet1), [{"id": 1, "name": "Alice"}, {"id": 2, "name": "Bob"}]
        )
        conn.execute(
            insert(sheet2), [{"id": 2, "name": "Bob"}, {"id": 3, "name": "Cara"}]
        )

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
    with (
        pytest.raises(exc.CompileError, match="nested subqueries"),
        engine.connect() as conn,
    ):
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
            .select_from(
                orders.join(users, users.c.id == orders.c.user_id, isouter=True)
            )
            .order_by(orders.c.id)
        )
        rows = conn.execute(stmt).all()
        assert rows == [(1, "Alice"), (2, None)]

    engine.dispose()


def test_e2e_full_outer_join_basic(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    t1 = Table(
        "t1",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("val1", String),
    )
    t2 = Table(
        "t2",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("val2", String),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(t1),
            [{"id": 1, "val1": "a1"}, {"id": 2, "val1": "a2"}, {"id": 4, "val1": "a4"}],
        )
        conn.execute(
            insert(t2),
            [{"id": 1, "val2": "b1"}, {"id": 2, "val2": "b2"}, {"id": 3, "val2": "b3"}],
        )

    with engine.connect() as conn:
        stmt = select(t1.c.id, t1.c.val1, t2.c.id, t2.c.val2).select_from(
            t1.outerjoin(t2, t1.c.id == t2.c.id, full=True)
        )
        rows = conn.execute(stmt).all()

    rows_sorted = sorted(
        rows, key=lambda r: (r[0] is None, r[0] or 0, r[2] is None, r[2] or 0)
    )
    assert rows_sorted == [
        (1, "a1", 1, "b1"),
        (2, "a2", 2, "b2"),
        (4, "a4", None, None),
        (None, None, 3, "b3"),
    ]
    engine.dispose()


def test_e2e_full_outer_join_all_match(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    t1 = Table(
        "t1", metadata, Column("id", Integer, primary_key=True), Column("val1", String)
    )
    t2 = Table(
        "t2", metadata, Column("id", Integer, primary_key=True), Column("val2", String)
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(t1), [{"id": 1, "val1": "a1"}, {"id": 2, "val1": "a2"}])
        conn.execute(insert(t2), [{"id": 1, "val2": "b1"}, {"id": 2, "val2": "b2"}])

    with engine.connect() as conn:
        full_stmt = select(t1.c.id, t1.c.val1, t2.c.val2).select_from(
            t1.outerjoin(t2, t1.c.id == t2.c.id, full=True)
        )
        inner_stmt = select(t1.c.id, t1.c.val1, t2.c.val2).select_from(
            t1.join(t2, t1.c.id == t2.c.id)
        )
        full_rows = sorted(conn.execute(full_stmt).all())
        inner_rows = sorted(conn.execute(inner_stmt).all())

    assert full_rows == inner_rows == [(1, "a1", "b1"), (2, "a2", "b2")]
    engine.dispose()


def test_e2e_full_outer_join_with_where(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    t1 = Table(
        "t1", metadata, Column("id", Integer, primary_key=True), Column("val1", String)
    )
    t2 = Table(
        "t2", metadata, Column("id", Integer, primary_key=True), Column("val2", String)
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(t1),
            [{"id": 1, "val1": "a1"}, {"id": 2, "val1": "a2"}, {"id": 4, "val1": "a4"}],
        )
        conn.execute(
            insert(t2),
            [{"id": 1, "val2": "b1"}, {"id": 2, "val2": "b2"}, {"id": 3, "val2": "b3"}],
        )

    with engine.connect() as conn:
        stmt = (
            select(t1.c.id, t2.c.id)
            .select_from(t1.outerjoin(t2, t1.c.id == t2.c.id, full=True))
            .where(sa.or_(t1.c.id == 4, t2.c.id == 3))
        )
        rows = sorted(
            conn.execute(stmt).all(),
            key=lambda r: (r[0] is None, r[0] or 0, r[1] is None, r[1] or 0),
        )

    assert rows == [(4, None), (None, 3)]
    engine.dispose()


def test_e2e_full_outer_join_select_star(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    t1 = Table(
        "t1", metadata, Column("id", Integer, primary_key=True), Column("val1", String)
    )
    t2 = Table(
        "t2", metadata, Column("id", Integer, primary_key=True), Column("val2", String)
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(t1), [{"id": 1, "val1": "a1"}, {"id": 4, "val1": "a4"}])
        conn.execute(insert(t2), [{"id": 1, "val2": "b1"}, {"id": 3, "val2": "b3"}])

    with engine.connect() as conn:
        stmt = select(t1, t2).select_from(
            t1.outerjoin(t2, t1.c.id == t2.c.id, full=True)
        )
        rows = sorted(
            conn.execute(stmt).all(),
            key=lambda r: (r[0] is None, r[0] or 0, r[2] is None, r[2] or 0),
        )

    assert rows == [
        (1, "a1", 1, "b1"),
        (4, "a4", None, None),
        (None, None, 3, "b3"),
    ]
    engine.dispose()


def test_e2e_cross_join_basic(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    t1 = Table(
        "t1", metadata, Column("id", Integer, primary_key=True), Column("val1", String)
    )
    t2 = Table(
        "t2", metadata, Column("id", Integer, primary_key=True), Column("val2", String)
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(t1), [{"id": 1, "val1": "a1"}, {"id": 2, "val1": "a2"}])
        conn.execute(insert(t2), [{"id": 10, "val2": "b10"}, {"id": 20, "val2": "b20"}])

    with engine.connect() as conn:
        stmt = select(t1.c.id, t2.c.id).select_from(t1.join(t2, sa.true()))
        rows = conn.execute(stmt).all()

    assert len(rows) == 4
    assert set(rows) == {(1, 10), (1, 20), (2, 10), (2, 20)}
    engine.dispose()


def test_e2e_cross_join_with_where(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    t1 = Table(
        "t1", metadata, Column("id", Integer, primary_key=True), Column("val1", String)
    )
    t2 = Table(
        "t2", metadata, Column("id", Integer, primary_key=True), Column("val2", String)
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(
            insert(t1),
            [{"id": 1, "val1": "a1"}, {"id": 2, "val1": "a2"}, {"id": 3, "val1": "a3"}],
        )
        conn.execute(
            insert(t2),
            [{"id": 2, "val2": "b2"}, {"id": 3, "val2": "b3"}, {"id": 9, "val2": "b9"}],
        )

    with engine.connect() as conn:
        stmt = (
            select(t1.c.id, t2.c.id)
            .select_from(t1.join(t2, sa.true()))
            .where(sa.and_(t1.c.id >= 2, t2.c.id <= 3))
            .order_by(t1.c.id)
        )
        rows = conn.execute(stmt).all()

    assert rows == [(2, 2), (2, 3), (3, 2), (3, 3)]
    engine.dispose()


def test_e2e_cross_join_select_star(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    t1 = Table(
        "t1", metadata, Column("id", Integer, primary_key=True), Column("val1", String)
    )
    t2 = Table(
        "t2", metadata, Column("id", Integer, primary_key=True), Column("val2", String)
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(t1), [{"id": 1, "val1": "a1"}, {"id": 2, "val1": "a2"}])
        conn.execute(insert(t2), [{"id": 10, "val2": "b10"}, {"id": 20, "val2": "b20"}])

    with engine.connect() as conn:
        stmt = select(t1, t2).select_from(t1.join(t2, sa.true()))
        rows = conn.execute(stmt).all()

    assert set(rows) == {
        (1, "a1", 10, "b10"),
        (1, "a1", 20, "b20"),
        (2, "a2", 10, "b10"),
        (2, "a2", 20, "b20"),
    }
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

    stmt = select(users.c.name, orders.c.user_id).join(
        orders,
        or_(users.c.id == orders.c.user_id, users.c.id == orders.c.dept_id),
    )
    with pytest.raises(exc.CompileError, match="OR"), engine.connect() as conn:
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

    stmt = select(users.c.name, orders.c.user_id).join(orders, users.c.id == 1)
    with (
        pytest.raises(exc.CompileError, match="column references"),
        engine.connect() as conn,
    ):
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

    stmt = select(users.c.name, orders.c.user_id).join(
        orders, users.c.id == users.c.age
    )
    with (
        pytest.raises(exc.CompileError, match="different join sources"),
        engine.connect() as conn,
    ):
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
            select(users.c.id).join(orders, users.c.id == orders.c.user_id),
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
        stmt = sa.union_all(
            select(t.c.id).where(t.c.id <= 2),
            select(t.c.id).where(t.c.id >= 2),
        ).order_by(t.c.id.desc())
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


def test_e2e_compound_with_branch_local_limit(tmp_path) -> None:
    """Branch-local LIMIT restricts rows from that branch only."""
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
        # Branch 1: all users with id <= 3 (all 3 rows)
        # Branch 2: users ordered by id DESC, LIMIT 1 → only id=3
        # UNION ALL → [1, 2, 3] + [3] = [1, 2, 3, 3]
        stmt = sa.union_all(
            select(t.c.id).where(t.c.id <= 3),
            select(t.c.id).order_by(t.c.id.desc()).limit(1),
        )
        rows = conn.execute(stmt).all()
        ids = sorted(r[0] for r in rows)
        # Without branch-local LIMIT, branch 2 would contribute all 3 rows.
        # With branch-local LIMIT 1, branch 2 contributes only 1 row.
        assert len(rows) == 4  # 3 from branch 1 + 1 from branch 2
        assert ids == [1, 2, 3, 3]

    engine.dispose()


def test_e2e_compound_mixed_operators(tmp_path) -> None:
    """Mixed compound operators with grouping raise CompileError.

    sa.union(a, sa.intersect(b, c)) produces grouped SQL that the
    Excel dialect cannot handle. CompileError is raised at compile time.
    """
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
                {"id": 4, "name": "Diana", "age": 28},
            ],
        )

    t = Table("users", MetaData(), autoload_with=engine)
    with engine.connect() as conn:
        stmt = sa.union(
            select(t.c.id).where(t.c.id.in_([1, 2, 3])),
            sa.intersect(
                select(t.c.id).where(t.c.id.in_([2, 3, 4])),
                select(t.c.id).where(t.c.id.in_([3, 4])),
            ),
        )
        with pytest.raises(exc.CompileError, match="grouped/nested compound"):
            conn.execute(stmt)

    engine.dispose()


# ===================================================================
# Phase 10: NOT / Parenthesized WHERE / NOT IN / NOT LIKE / NOT BETWEEN
# ===================================================================


def _seed_phase10(engine: sa.engine.Engine, users: Table) -> None:
    """Seed data for Phase 10 tests."""
    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Charlie", "age": 35},
                {"id": 4, "name": "Diana", "age": 28},
                {"id": 5, "name": "Eve", "age": 45},
            ],
        )


def test_e2e_not_in_literals(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name).where(users.c.name.not_in(["Alice", "Bob"]))
        ).all()
        names = sorted(r[0] for r in rows)
        assert names == ["Charlie", "Diana", "Eve"]

    engine.dispose()


def test_e2e_not_in_subquery(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    scores = Table(
        "scores",
        metadata,
        Column("user_id", Integer),
        Column("score", Integer),
    )
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
        conn.execute(
            insert(scores),
            [{"user_id": 1, "score": 90}, {"user_id": 2, "score": 80}],
        )

    with engine.connect() as conn:
        sub = select(scores.c.user_id)
        rows = conn.execute(select(users.c.name).where(users.c.id.not_in(sub))).all()
        names = [r[0] for r in rows]
        assert names == ["Charlie"]

    engine.dispose()


def test_e2e_not_like(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name).where(users.c.name.not_like("A%"))
        ).all()
        names = sorted(r[0] for r in rows)
        assert names == ["Bob", "Charlie", "Diana", "Eve"]

    engine.dispose()


def test_e2e_not_between(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name).where(~users.c.age.between(26, 34))
        ).all()
        names = sorted(r[0] for r in rows)
        # age NOT BETWEEN 26 AND 34: Bob(25), Charlie(35), Eve(45)
        assert names == ["Bob", "Charlie", "Eve"]

    engine.dispose()


def test_e2e_parenthesized_or_and(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.connect() as conn:
        # (age < 30 OR age > 40) AND name != 'Eve'
        rows = conn.execute(
            select(users.c.name).where(
                sa.and_(
                    sa.or_(users.c.age < 30, users.c.age > 40),
                    users.c.name != "Eve",
                )
            )
        ).all()
        names = sorted(r[0] for r in rows)
        # age<30: Bob(25), Diana(28); age>40: Eve(45)
        # AND name != 'Eve': Bob(25), Diana(28)
        assert names == ["Bob", "Diana"]

    engine.dispose()


def test_e2e_not_operator(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.connect() as conn:
        # NOT (age >= 30)
        rows = conn.execute(select(users.c.name).where(~(users.c.age >= 30))).all()
        names = sorted(r[0] for r in rows)
        # NOT age>=30: Bob(25), Diana(28)
        assert names == ["Bob", "Diana"]

    engine.dispose()


def test_e2e_parenthesized_complex(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.connect() as conn:
        # (name NOT LIKE 'A%' AND age >= 30) OR name = 'Eve'
        rows = conn.execute(
            select(users.c.name).where(
                sa.or_(
                    sa.and_(
                        users.c.name.not_like("A%"),
                        users.c.age >= 30,
                    ),
                    users.c.name == "Eve",
                )
            )
        ).all()
        names = sorted(r[0] for r in rows)
        # NOT LIKE 'A%' AND age>=30: Charlie(35), Eve(45)
        # OR name='Eve': already included
        assert names == ["Charlie", "Eve"]

    engine.dispose()


def test_e2e_update_with_not_in(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.begin() as conn:
        conn.execute(
            update(users).where(users.c.name.not_in(["Alice", "Bob"])).values(age=99)
        )

    with engine.connect() as conn:
        rows = conn.execute(select(users.c.name).where(users.c.age == 99)).all()
        names = sorted(r[0] for r in rows)
        assert names == ["Charlie", "Diana", "Eve"]

    engine.dispose()


def test_e2e_delete_with_not_between(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_phase10(engine, users)

    with engine.begin() as conn:
        conn.execute(delete(users).where(~users.c.age.between(25, 35)))

    with engine.connect() as conn:
        rows = conn.execute(select(users.c.name).order_by(users.c.id)).all()
        names = [r[0] for r in rows]
        # Kept: Bob(25), Alice(30), Charlie(35), Diana(28)
        assert sorted(names) == ["Alice", "Bob", "Charlie", "Diana"]

    engine.dispose()


# ──────────────────────────────────────────────────────────────────────
# Phase 11: Multi-column ORDER BY
# ──────────────────────────────────────────────────────────────────────


def _seed_multi_order(engine, users):
    """Seed data with deliberate ties for multi-column ORDER BY tests."""
    with engine.begin() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
                {"id": 3, "name": "Charlie", "age": 30},
                {"id": 4, "name": "Diana", "age": 25},
                {"id": 5, "name": "Eve", "age": 30},
            ],
        )


def test_e2e_multi_order_by_age_asc_name_desc(tmp_path) -> None:
    """ORDER BY age ASC, name DESC — ties in age broken by name descending."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_multi_order(engine, users)

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name, users.c.age).order_by(
                users.c.age.asc(), users.c.name.desc()
            )
        ).all()
        # age 25: Diana, Bob (DESC) → Diana, Bob
        # age 30: Eve, Charlie, Alice (DESC) → Eve, Charlie, Alice
        assert rows == [
            ("Diana", 25),
            ("Bob", 25),
            ("Eve", 30),
            ("Charlie", 30),
            ("Alice", 30),
        ]

    engine.dispose()


def test_e2e_multi_order_by_name_asc_age_desc(tmp_path) -> None:
    """ORDER BY name ASC, age DESC — primary sort by name."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_multi_order(engine, users)

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name, users.c.age).order_by(
                users.c.name.asc(), users.c.age.desc()
            )
        ).all()
        assert rows == [
            ("Alice", 30),
            ("Bob", 25),
            ("Charlie", 30),
            ("Diana", 25),
            ("Eve", 30),
        ]

    engine.dispose()


def test_e2e_multi_order_by_with_limit(tmp_path) -> None:
    """Multi-column ORDER BY + LIMIT."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_multi_order(engine, users)

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name, users.c.age)
            .order_by(users.c.age.asc(), users.c.name.asc())
            .limit(3)
        ).all()
        # age 25: Bob, Diana (ASC) → Bob, Diana
        # age 30: Alice (first of 3 at 30)
        assert rows == [
            ("Bob", 25),
            ("Diana", 25),
            ("Alice", 30),
        ]

    engine.dispose()


def test_e2e_multi_order_by_with_where(tmp_path) -> None:
    """Multi-column ORDER BY + WHERE filter."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_multi_order(engine, users)

    with engine.connect() as conn:
        rows = conn.execute(
            select(users.c.name, users.c.age)
            .where(users.c.age == 30)
            .order_by(users.c.name.desc(), users.c.id.asc())
        ).all()
        # Only age=30: Eve, Charlie, Alice (DESC by name)
        assert rows == [
            ("Eve", 30),
            ("Charlie", 30),
            ("Alice", 30),
        ]

    engine.dispose()


def test_e2e_multi_order_by_compound_union(tmp_path) -> None:
    """Compound UNION ALL with multi-column ORDER BY."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)
    _seed_multi_order(engine, users)

    t = Table("users", MetaData(), autoload_with=engine)
    with engine.connect() as conn:
        # UNION ALL with compound-level multi-column ORDER BY
        q1 = select(t.c.name, t.c.age).where(t.c.age <= 25)
        q2 = select(t.c.name, t.c.age).where(t.c.age >= 30)
        stmt = sa.union_all(q1, q2).order_by(t.c.age.asc(), t.c.name.asc())
        rows = conn.execute(stmt).all()
        # age 25: Bob, Diana; age 30: Alice, Charlie, Eve
        assert rows == [
            ("Bob", 25),
            ("Diana", 25),
            ("Alice", 30),
            ("Charlie", 30),
            ("Eve", 30),
        ]

    engine.dispose()


# ---- Phase 14: SELECT * with JOIN (via literal_column('*')) ----


def test_e2e_select_star_inner_join(tmp_path) -> None:
    """SELECT * FROM users JOIN orders ON ... returns all columns from both tables."""
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
        stmt = (
            sa.select(sa.literal_column("*"))
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .order_by(users.c.id)
        )
        result = conn.execute(stmt)
        rows = result.all()
        # All 6 columns: users(id, name, age), orders(id, user_id, amount)
        assert len(rows) == 2
        assert rows[0] == (1, "Alice", 30, 1, 1, 100)
        assert rows[1] == (2, "Bob", 25, 2, 2, 200)

    engine.dispose()


def test_e2e_select_star_left_join(tmp_path) -> None:
    """SELECT * with LEFT JOIN includes unmatched left rows with NULLs."""
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
            sa.select(sa.literal_column("*"))
            .select_from(
                users.join(orders, users.c.id == orders.c.user_id, isouter=True)
            )
            .order_by(users.c.id)
        )
        rows = conn.execute(stmt).all()
        # Alice matches, Bob has no order (NULLs for orders columns)
        assert len(rows) == 2
        assert rows[0] == (1, "Alice", 30, 1, 1, 100)
        assert rows[1] == (2, "Bob", 25, None, None, None)

    engine.dispose()


def test_e2e_select_star_chained_join(tmp_path) -> None:
    """SELECT * with chained JOINs across three tables."""
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
        conn.execute(insert(orders).values(id=10, user_id=1, amount=100))
        conn.execute(insert(items).values(id=100, order_id=10, sku="A-1"))

    with engine.connect() as conn:
        stmt = sa.select(sa.literal_column("*")).select_from(
            users.join(orders, users.c.id == orders.c.user_id).join(
                items, orders.c.id == items.c.order_id
            )
        )
        rows = conn.execute(stmt).all()
        # 9 columns: users(3) + orders(3) + items(3)
        assert len(rows) == 1
        assert rows[0] == (1, "Alice", 30, 10, 1, 100, 100, 10, "A-1")

    engine.dispose()


def test_e2e_select_star_with_where(tmp_path) -> None:
    """SELECT * with JOIN and WHERE clause filters correctly."""
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
        stmt = (
            sa.select(sa.literal_column("*"))
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .where(users.c.name == "Alice")
        )
        rows = conn.execute(stmt).all()
        assert len(rows) == 1
        assert rows[0] == (1, "Alice", 30, 1, 1, 100)

    engine.dispose()


def test_e2e_select_star_description_columns(tmp_path) -> None:
    """Verify result proxy exposes qualified column names for SELECT * with JOIN."""
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
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))

    with engine.connect() as conn:
        stmt = sa.select(sa.literal_column("*")).select_from(
            users.join(orders, users.c.id == orders.c.user_id)
        )
        result = conn.execute(stmt)
        cursor = result.cursor
        assert cursor is not None
        assert cursor.description is not None
        col_names = [desc[0] for desc in cursor.description]
        # SA compiles to: SELECT * FROM users JOIN orders ON users.id = orders.user_id
        # excel-dbapi expands * using table names as source refs
        assert col_names == [
            "users.id",
            "users.name",
            "users.age",
            "orders.id",
            "orders.user_id",
            "orders.amount",
        ]
        # Fetch to ensure result is consumable
        rows = result.all()
        assert len(rows) == 1

    engine.dispose()


def test_e2e_select_star_empty_result_has_description(tmp_path) -> None:
    """SELECT * with JOIN + impossible WHERE still populates cursor.description."""
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
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))

    with engine.connect() as conn:
        stmt = (
            sa.select(sa.literal_column("*"))
            .select_from(users.join(orders, users.c.id == orders.c.user_id))
            .where(users.c.id == 999)
        )
        result = conn.execute(stmt)
        # Check description BEFORE consuming rows — SA closes cursor after .all()
        cursor = result.cursor
        assert cursor is not None
        assert cursor.description is not None
        desc = cursor.description
        assert [d[0] for d in desc] == [
            "users.id",
            "users.name",
            "users.age",
            "orders.id",
            "orders.user_id",
            "orders.amount",
        ]
        rows = result.all()
        assert len(rows) == 0

    engine.dispose()


def test_e2e_select_star_labeled_rejected(tmp_path) -> None:
    """SELECT * AS alias with JOIN is rejected by the dbapi layer."""
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
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))

    with engine.connect() as conn:
        stmt = sa.select(sa.literal_column("*").label("all_cols")).select_from(
            users.join(orders, users.c.id == orders.c.user_id)
        )
        with pytest.raises(exc.ProgrammingError):
            conn.execute(stmt)

    engine.dispose()


def test_e2e_select_star_mixed_columns_rejected(tmp_path) -> None:
    """SELECT *, users.id with JOIN is rejected by the dbapi layer."""
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
        conn.execute(insert(orders).values(id=1, user_id=1, amount=100))

    with engine.connect() as conn:
        stmt = sa.select(sa.literal_column("*"), users.c.id).select_from(
            users.join(orders, users.c.id == orders.c.user_id)
        )
        with pytest.raises(exc.ProgrammingError):
            conn.execute(stmt)

    engine.dispose()


def test_e2e_aggregate_with_arithmetic_arg_rejected(tmp_path) -> None:
    """SUM(price * qty) via SA is rejected by the compiler guard."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    products = Table(
        "products",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("price", Integer),
        Column("qty", Integer),
    )
    metadata.create_all(engine)

    with engine.begin() as conn:
        conn.execute(insert(products).values(id=1, price=10, qty=3))

    with engine.connect() as conn:
        stmt = sa.select(sa.func.sum(products.c.price * products.c.qty))
        with pytest.raises(exc.CompileError):
            conn.execute(stmt)

    engine.dispose()


# ──────────────────────────────────────────────────────────────────────
# Phase 19: CASE WHEN expressions
# ──────────────────────────────────────────────────────────────────────


def _status_table(metadata: MetaData) -> Table:
    return Table(
        "people",
        metadata,
        Column("id", Integer, primary_key=True),
        Column("name", String),
        Column("age", Integer),
        Column("status", String),
    )


def _seed_status_data(engine: sa.engine.Engine, table: Table) -> None:
    with engine.begin() as conn:
        conn.execute(
            insert(table),
            [
                {"id": 1, "name": "Alice", "age": 30, "status": "active"},
                {"id": 2, "name": "Bob", "age": 25, "status": "inactive"},
                {"id": 3, "name": "Charlie", "age": 35, "status": "active"},
                {"id": 4, "name": "Diana", "age": 22, "status": "pending"},
            ],
        )


def test_e2e_case_when_searched_basic(tmp_path) -> None:
    """Searched CASE WHEN with multiple conditions."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = select(
            people.c.name,
            case(
                (people.c.age >= 30, sa.literal("senior")),
                else_=sa.literal("junior"),
            ).label("category"),
        ).order_by(people.c.id)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "senior"),
            ("Bob", "junior"),
            ("Charlie", "senior"),
            ("Diana", "junior"),
        ]

    engine.dispose()


def test_e2e_case_when_no_else(tmp_path) -> None:
    """CASE WHEN with no ELSE clause returns None for unmatched rows."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = select(
            people.c.name,
            case(
                (people.c.age >= 30, sa.literal("old")),
            ).label("tag"),
        ).order_by(people.c.id)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "old"),
            ("Bob", None),
            ("Charlie", "old"),
            ("Diana", None),
        ]

    engine.dispose()


def test_e2e_case_when_multiple_conditions(tmp_path) -> None:
    """CASE WHEN with multiple WHEN branches."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = select(
            people.c.name,
            case(
                (people.c.age >= 35, sa.literal("senior")),
                (people.c.age >= 25, sa.literal("mid")),
                else_=sa.literal("junior"),
            ).label("tier"),
        ).order_by(people.c.id)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "mid"),
            ("Bob", "mid"),
            ("Charlie", "senior"),
            ("Diana", "junior"),
        ]

    engine.dispose()


def test_e2e_case_when_with_alias(tmp_path) -> None:
    """CASE WHEN result columns expose correct alias via cursor.description."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = select(
            people.c.name,
            case(
                (people.c.status == "active", sa.literal("A")),
                else_=sa.literal("I"),
            ).label("code"),
        )
        result = conn.execute(stmt)
        keys = list(result.keys())
        assert "code" in keys
        rows = result.all()
        assert len(rows) == 4

    engine.dispose()


def test_e2e_case_when_in_update(tmp_path) -> None:
    """CASE WHEN in UPDATE SET clause."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.begin() as conn:
        stmt = update(people).values(
            age=case(
                (people.c.status == "active", 99),
                else_=0,
            )
        )
        conn.execute(stmt)

    with engine.connect() as conn:
        rows = conn.execute(
            select(people.c.name, people.c.age).order_by(people.c.id)
        ).all()
        assert rows == [
            ("Alice", 99),
            ("Bob", 0),
            ("Charlie", 99),
            ("Diana", 0),
        ]

    engine.dispose()


def test_e2e_case_when_in_update_with_where(tmp_path) -> None:
    """CASE WHEN in UPDATE SET with WHERE filter."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.begin() as conn:
        stmt = (
            update(people)
            .where(people.c.age >= 25)
            .values(
                age=case(
                    (people.c.status == "active", 100),
                    else_=50,
                )
            )
        )
        result = conn.execute(stmt)
        assert result.rowcount == 3  # Alice(30), Bob(25), Charlie(35)

    with engine.connect() as conn:
        rows = conn.execute(
            select(people.c.name, people.c.age).order_by(people.c.id)
        ).all()
        # Alice: active → 100, Bob: inactive → 50, Charlie: active → 100
        # Diana: age 22 < 25, not matched by WHERE → stays 22
        assert rows == [
            ("Alice", 100),
            ("Bob", 50),
            ("Charlie", 100),
            ("Diana", 22),
        ]

    engine.dispose()


def test_e2e_case_when_multiple_columns(tmp_path) -> None:
    """Multiple CASE WHEN columns in same SELECT."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = select(
            people.c.name,
            case(
                (people.c.age >= 30, sa.literal("senior")),
                else_=sa.literal("junior"),
            ).label("tier"),
            case(
                (people.c.status == "active", sa.literal("A")),
                (people.c.status == "inactive", sa.literal("I")),
                else_=sa.literal("P"),
            ).label("code"),
        ).order_by(people.c.id)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "senior", "A"),
            ("Bob", "junior", "I"),
            ("Charlie", "senior", "A"),
            ("Diana", "junior", "P"),
        ]

    engine.dispose()


def test_e2e_case_when_numeric_result(tmp_path) -> None:
    """CASE WHEN with numeric results."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = select(
            people.c.name,
            case(
                (people.c.age >= 35, 100),
                (people.c.age >= 25, 50),
                else_=10,
            ).label("score"),
        ).order_by(people.c.id)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", 50),
            ("Bob", 50),
            ("Charlie", 100),
            ("Diana", 10),
        ]

    engine.dispose()


def test_e2e_case_when_with_order_by(tmp_path) -> None:
    """CASE WHEN with ORDER BY on another column."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = select(
            people.c.name,
            case(
                (people.c.status == "active", sa.literal("yes")),
                else_=sa.literal("no"),
            ).label("active"),
        ).order_by(people.c.name)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "yes"),
            ("Bob", "no"),
            ("Charlie", "yes"),
            ("Diana", "no"),
        ]

    engine.dispose()


def test_e2e_case_when_with_where(tmp_path) -> None:
    """CASE WHEN combined with WHERE clause."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        stmt = (
            select(
                people.c.name,
                case(
                    (people.c.age >= 30, sa.literal("old")),
                    else_=sa.literal("young"),
                ).label("age_group"),
            )
            .where(people.c.status == "active")
            .order_by(people.c.id)
        )
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "old"),
            ("Charlie", "old"),
        ]

    engine.dispose()


def test_e2e_case_when_order_by_case_expression(tmp_path) -> None:
    """ORDER BY a CASE expression directly (not an alias)."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        case_expr = case(
            (people.c.status == "active", sa.literal(0)),
            else_=sa.literal(1),
        )
        stmt = select(people.c.name, people.c.status).order_by(
            case_expr.asc(), people.c.name
        )
        rows = conn.execute(stmt).all()
        # active first (0), then others (1), each sub-sorted by name
        assert rows == [
            ("Alice", "active"),
            ("Charlie", "active"),
            ("Bob", "inactive"),
            ("Diana", "pending"),
        ]

    engine.dispose()


def test_e2e_case_when_order_by_case_desc(tmp_path) -> None:
    """ORDER BY CASE expression DESC."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        case_expr = case(
            (people.c.status == "active", sa.literal(0)),
            else_=sa.literal(1),
        )
        stmt = select(people.c.name, people.c.status).order_by(
            case_expr.desc(), people.c.name
        )
        rows = conn.execute(stmt).all()
        # non-active first (1 DESC), then active (0 DESC), sub-sorted by name ASC
        assert rows == [
            ("Bob", "inactive"),
            ("Diana", "pending"),
            ("Alice", "active"),
            ("Charlie", "active"),
        ]

    engine.dispose()


def test_e2e_case_when_arithmetic_addition(tmp_path) -> None:
    """CASE expression used as operand in arithmetic (CASE...END + N)."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        expr = case(
            (people.c.status == "active", people.c.age),
            else_=sa.literal(0),
        ) + sa.literal(100)
        stmt = select(people.c.name, expr.label("boosted")).order_by(people.c.id)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", 130.0),
            ("Bob", 100.0),
            ("Charlie", 135.0),
            ("Diana", 100.0),
        ]

    engine.dispose()


def test_e2e_case_when_simple_case(tmp_path) -> None:
    """Simple CASE (CASE value WHEN match THEN ...) via searched CASE in SA."""
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    people = _status_table(metadata)
    metadata.create_all(engine)
    _seed_status_data(engine, people)

    with engine.connect() as conn:
        # SA 2.0 uses searched CASE syntax; we test mapping each status
        stmt = select(
            people.c.name,
            case(
                (people.c.status == "active", sa.literal("A")),
                (people.c.status == "inactive", sa.literal("I")),
                (people.c.status == "pending", sa.literal("P")),
                else_=sa.literal("?"),
            ).label("code"),
        ).order_by(people.c.id)
        rows = conn.execute(stmt).all()
        assert rows == [
            ("Alice", "A"),
            ("Bob", "I"),
            ("Charlie", "A"),
            ("Diana", "P"),
        ]

    engine.dispose()


def test_e2e_alter_table_add_column(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.commit()

    with engine.connect() as conn:
        conn.exec_driver_sql("ALTER TABLE users ADD COLUMN email TEXT")
        conn.commit()

    with engine.connect() as conn:
        rows = conn.exec_driver_sql(
            "SELECT id, name, email FROM users ORDER BY id"
        ).all()
        assert rows == [(1, "Alice", None), (2, "Bob", None)]

    engine.dispose()


def test_e2e_raw_create_table_reflects_declared_schema(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.exec_driver_sql(
            "CREATE TABLE users (id INTEGER PRIMARY KEY, age INTEGER NOT NULL, name TEXT)"
        )

    inspector = inspect(engine)
    columns = inspector.get_columns("users")
    assert [column["name"] for column in columns] == ["id", "age", "name"]

    id_column = next(column for column in columns if column["name"] == "id")
    age_column = next(column for column in columns if column["name"] == "age")
    assert isinstance(id_column["type"], sa.Integer)
    assert isinstance(age_column["type"], sa.Integer)
    assert age_column["nullable"] is False
    assert inspector.get_pk_constraint("users")["constrained_columns"] == ["id"]

    engine.dispose()


def test_e2e_raw_create_table_reflects_table_level_composite_primary_key(
    tmp_path,
) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.exec_driver_sql(
            "CREATE TABLE memberships (user_id INTEGER, group_id INTEGER, PRIMARY KEY (user_id, group_id))"
        )

    inspector = inspect(engine)
    columns = inspector.get_columns("memberships")
    assert [column["name"] for column in columns] == ["user_id", "group_id"]
    assert inspector.get_pk_constraint("memberships")["constrained_columns"] == [
        "user_id",
        "group_id",
    ]

    engine.dispose()


def test_e2e_raw_create_table_reflects_named_composite_primary_key(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.exec_driver_sql(
            "CREATE TABLE memberships (a INTEGER, b INTEGER, CONSTRAINT pk_memberships PRIMARY KEY (a, b))"
        )

    inspector = inspect(engine)
    assert inspector.get_pk_constraint("memberships")["constrained_columns"] == [
        "a",
        "b",
    ]

    engine.dispose()


def test_e2e_raw_create_table_numeric_aliases_reflect_as_float(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.exec_driver_sql(
            "CREATE TABLE metrics (x DECIMAL, y NUMERIC, z DOUBLE, w DOUBLE PRECISION)"
        )

    inspector = inspect(engine)
    columns = {
        column["name"]: column["type"] for column in inspector.get_columns("metrics")
    }
    assert isinstance(columns["x"], sa.Float)
    assert isinstance(columns["y"], sa.Float)
    assert isinstance(columns["z"], sa.Float)
    assert isinstance(columns["w"], sa.Float)

    engine.dispose()


def test_e2e_raw_drop_table_removes_metadata(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.exec_driver_sql("CREATE TABLE users (id INTEGER, name TEXT)")

    with engine.connect() as conn:
        import excel_dbapi

        raw_conn = conn.connection.dbapi_connection
        assert excel_dbapi.read_table_metadata(raw_conn, "users") is not None

    with engine.begin() as conn:
        conn.exec_driver_sql("DROP TABLE users")

    inspector = inspect(engine)
    assert inspector.has_table("users") is False
    assert inspector.get_columns("users") == []

    with engine.connect() as conn:
        import excel_dbapi

        raw_conn = conn.connection.dbapi_connection
        assert excel_dbapi.read_table_metadata(raw_conn, "users") is None

    engine.dispose()


def test_e2e_alter_table_add_float_column_reflects_as_float(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.exec_driver_sql("ALTER TABLE users ADD COLUMN score FLOAT")
        conn.commit()

    columns = inspect(engine).get_columns("users")
    score = next(col for col in columns if col["name"] == "score")
    assert isinstance(score["type"], sa.Float)

    engine.dispose()


def test_e2e_raw_alter_add_numeric_aliases_reflect_as_float(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.exec_driver_sql("CREATE TABLE metrics (id INTEGER PRIMARY KEY)")
        conn.exec_driver_sql("ALTER TABLE metrics ADD COLUMN x DECIMAL")
        conn.exec_driver_sql("ALTER TABLE metrics ADD COLUMN y NUMERIC")
        conn.exec_driver_sql("ALTER TABLE metrics ADD COLUMN z DOUBLE")
        conn.exec_driver_sql("ALTER TABLE metrics ADD COLUMN w DOUBLE PRECISION")

    inspector = inspect(engine)
    columns = {
        column["name"]: column["type"] for column in inspector.get_columns("metrics")
    }
    assert isinstance(columns["x"], sa.Float)
    assert isinstance(columns["y"], sa.Float)
    assert isinstance(columns["z"], sa.Float)
    assert isinstance(columns["w"], sa.Float)

    engine.dispose()


def test_e2e_raw_alter_preserves_existing_pk_and_nullability(tmp_path) -> None:
    engine = _engine_for(tmp_path)

    with engine.begin() as conn:
        conn.exec_driver_sql(
            "CREATE TABLE users (id INTEGER PRIMARY KEY, age INTEGER NOT NULL, name TEXT)"
        )

    with engine.begin() as conn:
        conn.exec_driver_sql("ALTER TABLE users ADD COLUMN email TEXT")

    inspector = inspect(engine)
    columns_after_add = {col["name"]: col for col in inspector.get_columns("users")}
    assert columns_after_add["age"]["nullable"] is False
    assert inspector.get_pk_constraint("users")["constrained_columns"] == ["id"]

    with engine.begin() as conn:
        conn.exec_driver_sql("ALTER TABLE users RENAME COLUMN age TO years")

    inspector = inspect(engine)
    columns_after_rename = {col["name"]: col for col in inspector.get_columns("users")}
    assert columns_after_rename["years"]["nullable"] is False
    assert inspector.get_pk_constraint("users")["constrained_columns"] == ["id"]

    with engine.begin() as conn:
        conn.exec_driver_sql("ALTER TABLE users DROP COLUMN name")

    inspector = inspect(engine)
    final_columns = {col["name"]: col for col in inspector.get_columns("users")}
    assert "name" not in final_columns
    assert final_columns["years"]["nullable"] is False
    assert inspector.get_pk_constraint("users")["constrained_columns"] == ["id"]

    engine.dispose()


def test_e2e_alter_table_drop_column(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.commit()

    with engine.connect() as conn:
        conn.exec_driver_sql("ALTER TABLE users DROP COLUMN age")
        conn.commit()

    with engine.connect() as conn:
        rows = conn.exec_driver_sql("SELECT id, name FROM users ORDER BY id").all()
        assert rows == [(1, "Alice"), (2, "Bob")]

    engine.dispose()


def test_e2e_alter_table_rename_column(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.commit()

    with engine.connect() as conn:
        conn.exec_driver_sql("ALTER TABLE users RENAME COLUMN name TO full_name")
        conn.commit()

    with engine.connect() as conn:
        rows = conn.exec_driver_sql(
            "SELECT id, full_name, age FROM users ORDER BY id"
        ).all()
        assert rows == [(1, "Alice", 30), (2, "Bob", 25)]

    engine.dispose()


def test_e2e_alter_table_add_column_then_insert(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.commit()

    with engine.connect() as conn:
        conn.exec_driver_sql("ALTER TABLE users ADD COLUMN email TEXT")
        conn.exec_driver_sql(
            "INSERT INTO users (id, name, age, email) VALUES (3, 'Charlie', 35, 'charlie@example.com')"
        )
        conn.commit()

    with engine.connect() as conn:
        rows = conn.exec_driver_sql(
            "SELECT id, name, age, email FROM users ORDER BY id"
        ).all()
        assert rows == [
            (1, "Alice", 30, None),
            (2, "Bob", 25, None),
            (3, "Charlie", 35, "charlie@example.com"),
        ]

    engine.dispose()


def test_e2e_alter_table_multiple_operations_sequence(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.commit()

    with engine.connect() as conn:
        conn.exec_driver_sql("ALTER TABLE users ADD COLUMN email TEXT")
        conn.exec_driver_sql("ALTER TABLE users RENAME COLUMN name TO full_name")
        conn.exec_driver_sql("ALTER TABLE users DROP COLUMN age")
        conn.commit()

    with engine.connect() as conn:
        rows = conn.exec_driver_sql(
            "SELECT id, full_name, email FROM users ORDER BY id"
        ).all()
        assert rows == [(1, "Alice", None), (2, "Bob", None)]

    engine.dispose()


def test_e2e_alter_table_add_existing_column_error(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.commit()

    with engine.connect() as conn, pytest.raises(exc.ProgrammingError):
        conn.exec_driver_sql("ALTER TABLE users ADD COLUMN name TEXT")

    engine.dispose()


def test_e2e_alter_table_drop_nonexistent_column_error(tmp_path) -> None:
    engine = _engine_for(tmp_path)
    metadata = MetaData()
    users = _users_table(metadata)
    metadata.create_all(engine)

    with engine.connect() as conn:
        conn.execute(
            insert(users),
            [
                {"id": 1, "name": "Alice", "age": 30},
                {"id": 2, "name": "Bob", "age": 25},
            ],
        )
        conn.commit()

    with engine.connect() as conn, pytest.raises(exc.ProgrammingError):
        conn.exec_driver_sql("ALTER TABLE users DROP COLUMN missing_col")

    engine.dispose()
