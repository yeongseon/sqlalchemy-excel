from __future__ import annotations

import importlib
from io import BytesIO

import pytest
from fastapi import FastAPI
from fastapi.testclient import TestClient
from openpyxl import Workbook
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session, sessionmaker
from sqlalchemy.pool import StaticPool

from sqlalchemy_excel.integrations.fastapi import create_import_router

_conftest = importlib.import_module("tests.conftest")
Base = _conftest.Base
SimpleUser = _conftest.SimpleUser


@pytest.fixture
def client() -> tuple[TestClient, sessionmaker[Session]]:
    engine = create_engine(
        "sqlite://",
        connect_args={"check_same_thread": False},
        poolclass=StaticPool,
    )
    Base.metadata.create_all(engine)
    session_local = sessionmaker(bind=engine, autoflush=False, autocommit=False)

    def get_session():
        with session_local() as session:
            yield session

    app = FastAPI()
    app.include_router(
        create_import_router(
            SimpleUser,
            prefix="/users",
            session_dependency=get_session,
            max_upload_bytes=1024 * 1024,
        )
    )
    return TestClient(app), session_local


def _xlsx_bytes(rows: list[list[object]]) -> bytes:
    workbook = Workbook()
    ws = workbook.active
    assert ws is not None
    ws.title = SimpleUser.__tablename__
    ws.append(["id", "name", "email", "age"])
    for row in rows:
        ws.append(row)

    bio = BytesIO()
    workbook.save(bio)
    return bio.getvalue()


def test_router_health(client: tuple[TestClient, sessionmaker[Session]]) -> None:
    test_client, _ = client
    response = test_client.get("/users/health")
    assert response.status_code == 200
    assert response.json()["status"] == "ok"


def test_router_template_download(client: tuple[TestClient, sessionmaker[Session]]) -> None:
    test_client, _ = client
    response = test_client.get("/users/template")
    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def test_router_validate_and_import_success(
    client: tuple[TestClient, sessionmaker[Session]],
) -> None:
    test_client, session_local = client
    payload = _xlsx_bytes([[1, "Alice", "alice@example.com", 30]])

    validate = test_client.post(
        "/users/validate",
        files={
            "file": (
                "users.xlsx",
                payload,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert validate.status_code == 200
    assert validate.json()["total_rows"] == 1

    imported = test_client.post(
        "/users/import",
        files={
            "file": (
                "users.xlsx",
                payload,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert imported.status_code == 200
    assert imported.json()["inserted"] == 1

    with session_local() as session:
        count = session.execute(select(SimpleUser)).scalars().all()
        assert len(count) == 1


def test_router_rejects_unsupported_content_type(
    client: tuple[TestClient, sessionmaker[Session]],
) -> None:
    test_client, _ = client
    response = test_client.post(
        "/users/validate",
        files={"file": ("users.txt", b"not-xlsx", "text/plain")},
    )
    assert response.status_code == 415


def test_router_rejects_oversized_payload() -> None:
    app = FastAPI()

    app.include_router(
        create_import_router(SimpleUser, prefix="/users", max_upload_bytes=10)
    )
    test_client = TestClient(app)

    response = test_client.post(
        "/users/validate",
        files={
            "file": (
                "users.xlsx",
                b"12345678910",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        },
    )
    assert response.status_code == 413
