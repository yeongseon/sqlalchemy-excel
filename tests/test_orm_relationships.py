"""Relationship boundary tests for ORM use with the Excel dialect.

Supported today:
- one-to-many inserts and joined eager loading

Known limitations:
- lazy relationship loading can return incorrect results
- many-to-many relationship loaders generate SQL patterns the backend can't parse
"""

from __future__ import annotations

from collections.abc import Iterator

import pytest
from sqlalchemy import ForeignKey, Integer, String, Table, Column, create_engine, select
from sqlalchemy import exc as sa_exc
from sqlalchemy.engine import Engine
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    Session,
    joinedload,
    mapped_column,
    relationship,
)


class Base(DeclarativeBase):
    pass


author_book = Table(
    "author_book",
    Base.metadata,
    Column("author_id", ForeignKey("authors.id"), primary_key=True),
    Column("book_id", ForeignKey("books.id"), primary_key=True),
)


class Parent(Base):
    __tablename__ = "parents"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String)
    children: Mapped[list["Child"]] = relationship(back_populates="parent")


class Child(Base):
    __tablename__ = "children"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    parent_id: Mapped[int] = mapped_column(ForeignKey("parents.id"))
    name: Mapped[str] = mapped_column(String)
    parent: Mapped[Parent] = relationship(back_populates="children")


class Author(Base):
    __tablename__ = "authors"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    name: Mapped[str] = mapped_column(String)
    books: Mapped[list["Book"]] = relationship(
        secondary=author_book,
        back_populates="authors",
    )


class Book(Base):
    __tablename__ = "books"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    title: Mapped[str] = mapped_column(String)
    authors: Mapped[list[Author]] = relationship(
        secondary=author_book,
        back_populates="books",
    )


@pytest.fixture
def relationship_engine(tmp_xlsx: str) -> Iterator[Engine]:
    engine = create_engine(f"excel:///{tmp_xlsx}")
    Base.metadata.create_all(engine)
    yield engine
    engine.dispose()


def test_one_to_many_relationship_round_trip(relationship_engine: Engine) -> None:
    with Session(relationship_engine) as session:
        session.add(
            Parent(
                id=1,
                name="parent-1",
                children=[
                    Child(id=10, name="child-1"),
                    Child(id=11, name="child-2"),
                ],
            )
        )
        session.commit()

    with Session(relationship_engine) as session:
        parent = (
            session.query(Parent)
            .options(joinedload(Parent.children))
            .filter(Parent.id == 1)
            .one()
        )
        assert [child.name for child in parent.children] == ["child-1", "child-2"]


@pytest.mark.xfail(
    reason="Lazy one-to-many loaders currently return empty collections with this backend.",
)
def test_one_to_many_lazy_loading_boundary(relationship_engine: Engine) -> None:
    with Session(relationship_engine) as session:
        session.add(Parent(id=2, name="parent-2", children=[Child(id=12, name="child-3")]))
        session.commit()

    with Session(relationship_engine) as session:
        parent = session.query(Parent).filter(Parent.id == 2).one()
        assert [child.name for child in parent.children] == ["child-3"]


def test_many_to_many_association_table_persists(relationship_engine: Engine) -> None:
    with Session(relationship_engine) as session:
        author = Author(id=1, name="Ada")
        author.books.append(Book(id=1, title="Spec"))
        session.add(author)
        session.commit()

    with relationship_engine.connect() as conn:
        links = conn.execute(select(author_book)).all()
        assert links == [(1, 1)]


@pytest.mark.xfail(
    raises=(sa_exc.CompileError, sa_exc.DBAPIError),
    reason="Many-to-many relationship loaders emit SQL not fully supported by excel-dbapi.",
)
def test_many_to_many_relationship_loading_boundary(relationship_engine: Engine) -> None:
    with Session(relationship_engine) as session:
        author = Author(id=2, name="Grace")
        author.books.append(Book(id=2, title="Parser"))
        session.add(author)
        session.commit()

    with Session(relationship_engine) as session:
        loaded = session.query(Author).filter(Author.id == 2).one()
        assert [book.title for book in loaded.books] == ["Parser"]
