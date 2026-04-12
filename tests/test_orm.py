"""ORM smoke tests — declarative model with Session."""

from __future__ import annotations

import pytest
from sqlalchemy import Integer, String, create_engine
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    Session,
    mapped_column,
)


class Base(DeclarativeBase):
    pass


class User(Base):
    __tablename__ = "users"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String)
    age: Mapped[int] = mapped_column(Integer)


class Product(Base):
    __tablename__ = "products"
    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    title: Mapped[str] = mapped_column(String)


@pytest.fixture
def orm_engine(tmp_xlsx):
    eng = create_engine(f"excel:///{tmp_xlsx}")
    Base.metadata.create_all(eng)
    yield eng
    eng.dispose()


class TestORMBasic:
    """Basic ORM operations."""

    def test_add_and_query(self, orm_engine):
        with Session(orm_engine) as session:
            session.add(User(id=1, name="Alice", age=30))
            session.commit()

        with Session(orm_engine) as session:
            users = session.query(User).all()
            assert len(users) == 1
            assert users[0].name == "Alice"
            assert users[0].age == 30

    def test_add_multiple(self, orm_engine):
        with Session(orm_engine) as session:
            session.add_all(
                [
                    User(id=1, name="Alice", age=30),
                    User(id=2, name="Bob", age=25),
                ]
            )
            session.commit()

        with Session(orm_engine) as session:
            users = session.query(User).all()
            assert len(users) == 2

    def test_filter(self, orm_engine):
        with Session(orm_engine) as session:
            session.add_all(
                [
                    User(id=1, name="Alice", age=30),
                    User(id=2, name="Bob", age=25),
                    User(id=3, name="Charlie", age=35),
                ]
            )
            session.commit()

        with Session(orm_engine) as session:
            young = session.query(User).filter(User.age < 30).all()
            assert len(young) == 1
            assert young[0].name == "Bob"

    def test_update_via_orm(self, orm_engine):
        with Session(orm_engine) as session:
            session.add(User(id=1, name="Alice", age=30))
            session.commit()

        with Session(orm_engine) as session:
            user = session.query(User).filter(User.id == 1).one()
            user.age = 31
            session.commit()

        with Session(orm_engine) as session:
            user = session.query(User).filter(User.id == 1).one()
            assert user.age == 31

    def test_delete_via_orm(self, orm_engine):
        with Session(orm_engine) as session:
            session.add(User(id=1, name="Alice", age=30))
            session.commit()

        with Session(orm_engine) as session:
            user = session.query(User).filter(User.id == 1).one()
            session.delete(user)
            session.commit()

        with Session(orm_engine) as session:
            users = session.query(User).all()
            assert len(users) == 0


class TestORMMultipleTables:
    """Test ORM with multiple tables."""

    def test_multiple_models(self, orm_engine):
        with Session(orm_engine) as session:
            session.add(User(id=1, name="Alice", age=30))
            session.add(Product(id=1, title="Widget"))
            session.commit()

        with Session(orm_engine) as session:
            users = session.query(User).all()
            products = session.query(Product).all()
            assert len(users) == 1
            assert len(products) == 1
            assert products[0].title == "Widget"
