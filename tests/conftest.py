"""Shared test fixtures for sqlalchemy-excel."""

from __future__ import annotations

import enum
from datetime import date, datetime

import pytest
from sqlalchemy import ForeignKey, String, Text, create_engine
from sqlalchemy.orm import (
    DeclarativeBase,
    Mapped,
    Session,
    mapped_column,
    relationship,
)

# --- Sample ORM Models ---


class Base(DeclarativeBase):
    pass


class Department(Base):
    __tablename__ = "departments"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100), unique=True)
    code: Mapped[str] = mapped_column(String(10), unique=True)

    employees: Mapped[list[Employee]] = relationship(back_populates="department")


class EmployeeStatus(enum.Enum):
    ACTIVE = "active"
    INACTIVE = "inactive"
    ON_LEAVE = "on_leave"


class Employee(Base):
    __tablename__ = "employees"

    id: Mapped[int] = mapped_column(primary_key=True)
    email: Mapped[str] = mapped_column(String(255), unique=True)
    first_name: Mapped[str] = mapped_column(String(100))
    last_name: Mapped[str] = mapped_column(String(100))
    status: Mapped[EmployeeStatus] = mapped_column(default=EmployeeStatus.ACTIVE)
    salary: Mapped[float | None] = mapped_column(default=None)
    hire_date: Mapped[date] = mapped_column(default=date.today)
    department_id: Mapped[int | None] = mapped_column(
        ForeignKey("departments.id"), default=None
    )
    notes: Mapped[str | None] = mapped_column(Text, default=None)

    department: Mapped[Department | None] = relationship(back_populates="employees")


class Product(Base):
    __tablename__ = "products"

    id: Mapped[int] = mapped_column(primary_key=True)
    sku: Mapped[str] = mapped_column(String(50), unique=True)
    name: Mapped[str] = mapped_column(String(200))
    price: Mapped[float] = mapped_column()
    quantity: Mapped[int] = mapped_column(default=0)
    is_active: Mapped[bool] = mapped_column(default=True)
    created_at: Mapped[datetime] = mapped_column(default=datetime.now)


# --- Simple model for basic testing ---


class SimpleUser(Base):
    __tablename__ = "simple_users"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100))
    email: Mapped[str] = mapped_column(String(255))
    age: Mapped[int | None] = mapped_column(default=None)


# --- Fixtures ---


@pytest.fixture
def engine():
    """Create an in-memory SQLite engine with all tables."""
    eng = create_engine("sqlite:///:memory:")
    Base.metadata.create_all(eng)
    return eng


@pytest.fixture
def session(engine):
    """Create a SQLAlchemy session."""
    with Session(engine) as sess:
        yield sess


@pytest.fixture
def sample_departments(session):
    """Create sample department records."""
    departments = [
        Department(id=1, name="Engineering", code="ENG"),
        Department(id=2, name="Marketing", code="MKT"),
        Department(id=3, name="Sales", code="SLS"),
    ]
    session.add_all(departments)
    session.commit()
    return departments


@pytest.fixture
def sample_employees(session, sample_departments):
    """Create sample employee records."""
    employees = [
        Employee(
            id=1,
            email="alice@example.com",
            first_name="Alice",
            last_name="Smith",
            status=EmployeeStatus.ACTIVE,
            salary=85000.0,
            hire_date=date(2023, 1, 15),
            department_id=1,
        ),
        Employee(
            id=2,
            email="bob@example.com",
            first_name="Bob",
            last_name="Jones",
            status=EmployeeStatus.ACTIVE,
            salary=92000.0,
            hire_date=date(2022, 6, 1),
            department_id=1,
        ),
        Employee(
            id=3,
            email="carol@example.com",
            first_name="Carol",
            last_name="Williams",
            status=EmployeeStatus.ON_LEAVE,
            salary=78000.0,
            hire_date=date(2024, 3, 10),
            department_id=2,
        ),
    ]
    session.add_all(employees)
    session.commit()
    return employees


@pytest.fixture
def tmp_xlsx(tmp_path):
    """Return a factory for temporary xlsx file paths."""

    def _factory(name: str = "test.xlsx") -> str:
        return str(tmp_path / name)

    return _factory
