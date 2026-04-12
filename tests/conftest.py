"""Shared test fixtures."""

from __future__ import annotations

import pytest
from sqlalchemy import create_engine
from sqlalchemy.orm import Session


@pytest.fixture
def tmp_xlsx(tmp_path):
    """Return a path to a temporary .xlsx file."""
    return str(tmp_path / "test.xlsx")


@pytest.fixture
def engine(tmp_xlsx):
    """Create a SQLAlchemy engine pointing to a temporary Excel file."""
    eng = create_engine(f"excel:///{tmp_xlsx}")
    yield eng
    eng.dispose()


@pytest.fixture
def session(engine):
    """Create a SQLAlchemy session."""
    with Session(engine) as sess:
        yield sess
