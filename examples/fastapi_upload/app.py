"""FastAPI application demonstrating sqlalchemy-excel integration.

Run with:
    pip install sqlalchemy-excel[fastapi]
    uvicorn examples.fastapi_upload.app:app --reload

Endpoints:
    GET  /                          — Welcome page
    GET  /employees/template        — Download Excel template
    POST /employees/validate        — Validate an uploaded Excel file
    POST /employees/import          — Validate and import to database
    GET  /employees/health          — Health check
    GET  /employees/export          — Export all employees to Excel
"""

from __future__ import annotations

from typing import Any

from fastapi import FastAPI
from fastapi.responses import Response
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session

from examples.fastapi_upload.models import Base, Employee
from sqlalchemy_excel import ExcelExporter, ExcelMapping
from sqlalchemy_excel.integrations.fastapi import create_import_router

# ---------------------------------------------------------------------------
# Database setup (in-memory SQLite for demonstration)
# ---------------------------------------------------------------------------

DATABASE_URL = "sqlite:///./example.db"

engine = create_engine(DATABASE_URL, echo=False)
Base.metadata.create_all(engine)


def get_session():  # type: ignore[no-untyped-def]
    """FastAPI dependency that yields a SQLAlchemy session."""
    with Session(engine) as session:
        yield session


# ---------------------------------------------------------------------------
# FastAPI application
# ---------------------------------------------------------------------------

app = FastAPI(
    title="sqlalchemy-excel Demo",
    description="Demonstrates Excel template generation, validation, and import.",
    version="0.1.0",
)


# ---------------------------------------------------------------------------
# Auto-generated endpoints via create_import_router
# ---------------------------------------------------------------------------

employee_router = create_import_router(
    Employee,
    prefix="/employees",
    tags=["employees"],
    session_dependency=get_session,
)
app.include_router(employee_router)


# ---------------------------------------------------------------------------
# Additional custom endpoints
# ---------------------------------------------------------------------------


@app.get("/")
def root() -> dict[str, str]:
    """Welcome endpoint with usage instructions."""
    return {
        "message": "sqlalchemy-excel FastAPI demo",
        "docs": "/docs",
        "template": "/employees/template",
        "health": "/employees/health",
    }


@app.get("/employees/export", tags=["employees"])
def export_employees() -> Response:
    """Export all employees to an Excel file."""
    mapping = ExcelMapping.from_model(Employee)
    exporter = ExcelExporter([mapping])

    with Session(engine) as session:
        employees = list(session.execute(select(Employee)).scalars().all())

    content: Any = exporter.export(employees)

    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="employees_export.xlsx"'},
    )
