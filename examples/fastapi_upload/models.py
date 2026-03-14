"""SQLAlchemy ORM models for the FastAPI upload example.

This module defines sample models used to demonstrate
sqlalchemy-excel's template generation, validation, and import features.
"""

from __future__ import annotations

import enum
from datetime import date  # noqa: TC003

from sqlalchemy import ForeignKey, String
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


class Base(DeclarativeBase):
    """Shared declarative base for all models."""


class Department(enum.Enum):
    """Department enum for dropdown validation in Excel templates."""

    engineering = "Engineering"
    marketing = "Marketing"
    sales = "Sales"
    hr = "Human Resources"
    finance = "Finance"


class Team(Base):
    """Team model — demonstrates foreign key relationships."""

    __tablename__ = "teams"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100))
    department: Mapped[Department] = mapped_column()

    employees: Mapped[list[Employee]] = relationship(back_populates="team")


class Employee(Base):
    """Employee model — the primary model for this example.

    Demonstrates:
    - Various column types (int, str, date, enum, nullable)
    - Foreign key references
    - Column length constraints
    - Default values
    - Column documentation via ``doc``
    """

    __tablename__ = "employees"

    id: Mapped[int] = mapped_column(primary_key=True)
    first_name: Mapped[str] = mapped_column(
        String(50),
        doc="Employee's first/given name",
    )
    last_name: Mapped[str] = mapped_column(
        String(50),
        doc="Employee's last/family name",
    )
    email: Mapped[str] = mapped_column(
        String(255),
        unique=True,
        doc="Corporate email address",
    )
    hire_date: Mapped[date] = mapped_column(
        doc="Date the employee was hired",
    )
    department: Mapped[Department] = mapped_column(
        doc="Department assignment",
    )
    team_id: Mapped[int | None] = mapped_column(
        ForeignKey("teams.id"),
        default=None,
        doc="Optional team assignment",
    )
    salary: Mapped[float | None] = mapped_column(
        default=None,
        doc="Annual salary (optional, for HR use)",
    )
    is_active: Mapped[bool] = mapped_column(
        default=True,
        doc="Whether the employee is currently active",
    )

    team: Mapped[Team | None] = relationship(back_populates="employees")
