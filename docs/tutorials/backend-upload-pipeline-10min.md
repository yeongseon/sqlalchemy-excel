# 10-Minute Backend Upload Pipeline (FastAPI + SQLAlchemy)

This tutorial shows a minimal production-ready flow:
1. Download template
2. Upload and validate Excel
3. Import into DB

## Install

```bash
pip install "sqlalchemy-excel[fastapi]" sqlalchemy
```

## App

```python
from fastapi import Depends, FastAPI
from sqlalchemy import create_engine
from sqlalchemy.orm import Session, sessionmaker

from sqlalchemy_excel.integrations.fastapi import create_import_router
from tests.conftest import Base, SimpleUser  # replace with your own model

engine = create_engine("sqlite:///./app.db")
Base.metadata.create_all(engine)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)


def get_session():
    with SessionLocal() as session:
        yield session


app = FastAPI()
app.include_router(
    create_import_router(
        SimpleUser,
        prefix="/users",
        tags=["users"],
        session_dependency=get_session,
        max_upload_bytes=5 * 1024 * 1024,
        allowed_content_types={
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "application/octet-stream",
        },
    )
)
```

## Test quickly

- `GET /users/template`
- `POST /users/validate` (multipart file)
- `POST /users/import` (multipart file)

## Production defaults

- Restrict upload content type to XLSX
- Set explicit max upload size
- Keep formula-sanitized exports enabled (default)
- Validate before import, import in transactions
