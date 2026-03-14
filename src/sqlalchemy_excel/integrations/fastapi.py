"""FastAPI integration for sqlalchemy-excel.

Provides router factory for creating upload/template endpoints.
Requires: pip install sqlalchemy-excel[fastapi]
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

from sqlalchemy_excel._compat import import_optional

if TYPE_CHECKING:
    from fastapi import APIRouter


def create_import_router(
    model: type,
    *,
    prefix: str = "",
    tags: list[str] | None = None,
    session_dependency: Any = None,
) -> APIRouter:
    """Create a FastAPI router with Excel template and import endpoints.

    This factory generates four endpoints:
    - GET {prefix}/template — Download Excel template
    - POST {prefix}/validate — Validate an uploaded Excel file
    - POST {prefix}/import — Validate and import to database
    - GET {prefix}/health — Health check

    Args:
        model: SQLAlchemy ORM model class.
        prefix: URL prefix for all endpoints.
        tags: OpenAPI tags for the endpoints.
        session_dependency: FastAPI Depends callable that yields a Session.

    Returns:
        Configured FastAPI APIRouter.

    Raises:
        ImportError: If fastapi is not installed.

    Example:
        >>> from myapp.models import User
        >>> from myapp.database import get_session
        >>> router = create_import_router(
        ...     User,
        ...     prefix="/users",
        ...     session_dependency=get_session,
        ... )
        >>> app.include_router(router)
    """
    fastapi = import_optional("fastapi", "fastapi")
    APIRouter = fastapi.APIRouter  # noqa: N806
    Depends = fastapi.Depends  # noqa: N806
    File = fastapi.File  # noqa: N806
    UploadFile = fastapi.UploadFile  # noqa: N806
    Response = fastapi.Response  # noqa: N806
    HTTPException = fastapi.HTTPException  # noqa: N806
    _ = UploadFile

    from sqlalchemy_excel.mapping import ExcelMapping
    from sqlalchemy_excel.template import (
        ExcelTemplate,  # pyright: ignore[reportMissingImports]
    )
    from sqlalchemy_excel.validation import ExcelValidator

    router = APIRouter(prefix=prefix, tags=tags or [])
    mapping = ExcelMapping.from_model(model)

    @router.get("/template")  # type: ignore[misc]
    def download_template() -> Any:
        """Download an Excel template for this model."""
        tpl = ExcelTemplate([mapping], include_sample_data=True)
        content = tpl.to_bytes()
        filename = f"{mapping.sheet_name}_template.xlsx"
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    @router.post("/validate")  # type: ignore[misc]
    async def validate_upload(
        file: Any = File(...),  # noqa: B008
    ) -> dict[str, Any]:
        """Validate an uploaded Excel file against the model schema."""
        validator = ExcelValidator([mapping])
        content = await file.read()

        from io import BytesIO

        report = validator.validate(BytesIO(content))
        return report.to_dict()

    if session_dependency is not None:

        @router.post("/import")  # type: ignore[misc]
        async def import_upload(
            file: Any = File(...),  # noqa: B008
            session: Any = Depends(session_dependency),  # noqa: B008
        ) -> dict[str, Any]:
            """Validate and import an uploaded Excel file to the database."""
            from io import BytesIO

            from sqlalchemy_excel.load import ExcelImporter

            content = await file.read()
            source = BytesIO(content)

            # Validate first
            validator = ExcelValidator([mapping])
            report = validator.validate(source)
            if report.has_errors:
                raise HTTPException(
                    status_code=422,
                    detail={
                        "message": "Validation failed",
                        "report": report.to_dict(),
                    },
                )

            # Import
            source.seek(0)
            importer = ExcelImporter([mapping], session=session)
            result = importer.insert(source)
            session.commit()

            return {
                "message": "Import successful",
                "inserted": result.inserted,
                "updated": result.updated,
                "skipped": result.skipped,
                "duration_ms": result.duration_ms,
            }

    @router.get("/health")  # type: ignore[misc]
    def health() -> dict[str, str]:
        """Health check endpoint."""
        return {"status": "ok", "model": model.__name__}

    return router  # type: ignore[no-any-return]
