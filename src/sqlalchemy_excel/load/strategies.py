"""Database load strategies for insert, upsert, and dry-run modes."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TYPE_CHECKING, Protocol, TypeVar

from sqlalchemy import select
from sqlalchemy.exc import IntegrityError

if TYPE_CHECKING:
    from collections.abc import Iterable

    from sqlalchemy.orm import DeclarativeBase, Session

_T = TypeVar("_T")


def _chunk(iterable: Iterable[_T], size: int) -> Iterable[list[_T]]:
    """Yield lists containing at most ``size`` items.

    Args:
        iterable: Source iterable.
        size: Maximum chunk size.

    Yields:
        Lists of items from ``iterable``.

    Raises:
        ValueError: If ``size`` is less than 1.
    """

    if size < 1:
        raise ValueError("Chunk size must be at least 1")

    batch: list[_T] = []
    for item in iterable:
        batch.append(item)
        if len(batch) >= size:
            yield batch
            batch = []

    if batch:
        yield batch


@dataclass(slots=True)
class ImportResult:
    """Result object returned by load strategy execution.

    Attributes:
        inserted: Number of inserted rows.
        updated: Number of updated rows.
        skipped: Number of skipped rows.
        failed: Number of failed rows.
        errors: Human-readable error messages collected during import.
        duration_ms: Total operation duration in milliseconds.
    """

    inserted: int = 0
    updated: int = 0
    skipped: int = 0
    failed: int = 0
    errors: list[str] = field(default_factory=list)
    duration_ms: float = 0.0

    @property
    def total(self) -> int:
        """Return total number of processed rows."""

        return self.inserted + self.updated + self.skipped + self.failed

    def summary(self) -> str:
        """Return a concise human-readable summary of the import result."""

        return (
            f"inserted={self.inserted}, updated={self.updated}, "
            f"skipped={self.skipped}, failed={self.failed}, total={self.total}, "
            f"errors={len(self.errors)}, duration_ms={self.duration_ms:.2f}"
        )


class LoadStrategy(Protocol):
    """Protocol for database load strategy implementations."""

    def execute(
        self,
        session: Session,
        model_class: type[DeclarativeBase],
        rows: Iterable[dict[str, object]],
        key_columns: list[str],
        batch_size: int,
    ) -> ImportResult:
        """Execute the load strategy and return aggregate import results.

        Args:
            session: Active SQLAlchemy session.
            model_class: Target ORM model class.
            rows: Rows to load.
            key_columns: Key columns used for upsert modes.
            batch_size: Maximum number of rows processed per batch.

        Returns:
            Aggregated import result.
        """

        ...


class InsertStrategy:
    """Insert-only strategy using batched ORM object creation."""

    def execute(
        self,
        session: Session,
        model_class: type[DeclarativeBase],
        rows: Iterable[dict[str, object]],
        key_columns: list[str],
        batch_size: int,
    ) -> ImportResult:
        """Insert all rows as new ORM objects.

        Args:
            session: Active SQLAlchemy session.
            model_class: Target ORM model class.
            rows: Rows to insert.
            key_columns: Unused by insert strategy.
            batch_size: Batch size for flush boundaries.

        Returns:
            Insert execution result.
        """

        del key_columns

        result = ImportResult()
        for batch in _chunk(rows, batch_size):
            savepoint = session.begin_nested()
            try:
                objects = [model_class(**row) for row in batch]
                session.add_all(objects)
                session.flush()
                result.inserted += len(objects)
            except IntegrityError as exc:
                result.failed += len(batch)
                result.errors.append(str(exc))
                if savepoint.is_active:
                    savepoint.rollback()
            except Exception as exc:
                result.failed += len(batch)
                result.errors.append(str(exc))
                if savepoint.is_active:
                    savepoint.rollback()
            else:
                if savepoint.is_active:
                    savepoint.commit()

        return result


class UpsertStrategy:
    """Upsert strategy that updates existing rows and inserts new rows."""

    def execute(
        self,
        session: Session,
        model_class: type[DeclarativeBase],
        rows: Iterable[dict[str, object]],
        key_columns: list[str],
        batch_size: int,
    ) -> ImportResult:
        """Upsert rows using key-based lookup per row.

        Args:
            session: Active SQLAlchemy session.
            model_class: Target ORM model class.
            rows: Rows to upsert.
            key_columns: Key columns used to find existing records.
            batch_size: Batch size for flush boundaries.

        Returns:
            Upsert execution result.

        Raises:
            ValueError: If ``key_columns`` is empty.
        """

        if not key_columns:
            raise ValueError("Upsert strategy requires at least one key column")

        result = ImportResult()
        for batch in _chunk(rows, batch_size):
            savepoint = session.begin_nested()
            try:
                for row in batch:
                    key_filter = {key: row[key] for key in key_columns}
                    existing = session.execute(
                        select(model_class).filter_by(**key_filter)
                    ).scalar_one_or_none()
                    if existing is None:
                        session.add(model_class(**row))
                        result.inserted += 1
                        continue

                    for column_name, value in row.items():
                        setattr(existing, column_name, value)
                    result.updated += 1

                session.flush()
            except IntegrityError as exc:
                if savepoint.is_active:
                    savepoint.rollback()

                self._recover_failed_batch(
                    session=session,
                    model_class=model_class,
                    batch=batch,
                    key_columns=key_columns,
                    result=result,
                    initial_error=exc,
                )
            except Exception as exc:
                result.failed += len(batch)
                result.errors.append(str(exc))
                if savepoint.is_active:
                    savepoint.rollback()
            else:
                if savepoint.is_active:
                    savepoint.commit()

        return result

    def _recover_failed_batch(
        self,
        *,
        session: Session,
        model_class: type[DeclarativeBase],
        batch: list[dict[str, object]],
        key_columns: list[str],
        result: ImportResult,
        initial_error: IntegrityError,
    ) -> None:
        """Recover a failed upsert batch by retrying rows individually.

        Args:
            session: Active SQLAlchemy session.
            model_class: Target ORM model class.
            batch: Failed batch.
            key_columns: Key columns for existing-row lookup.
            result: Mutable aggregate result.
            initial_error: IntegrityError raised by batch flush.
        """

        result.errors.append(str(initial_error))

        inserted_in_batch = 0
        updated_in_batch = 0
        for row in batch:
            key_filter = {key: row[key] for key in key_columns}
            existing = session.execute(
                select(model_class).filter_by(**key_filter)
            ).scalar_one_or_none()
            if existing is None:
                inserted_in_batch += 1
            else:
                updated_in_batch += 1

        result.inserted -= inserted_in_batch
        result.updated -= updated_in_batch

        for row in batch:
            savepoint = session.begin_nested()
            try:
                key_filter = {key: row[key] for key in key_columns}
                existing = session.execute(
                    select(model_class).filter_by(**key_filter)
                ).scalar_one_or_none()
                if existing is None:
                    session.add(model_class(**row))
                    result.inserted += 1
                else:
                    for column_name, value in row.items():
                        setattr(existing, column_name, value)
                    result.updated += 1
                session.flush()
            except Exception as exc:
                result.failed += 1
                result.errors.append(str(exc))
                if savepoint.is_active:
                    savepoint.rollback()
            else:
                if savepoint.is_active:
                    savepoint.commit()


class DryRunStrategy:
    """Dry-run strategy that validates insertability without persisting data."""

    def execute(
        self,
        session: Session,
        model_class: type[DeclarativeBase],
        rows: Iterable[dict[str, object]],
        key_columns: list[str],
        batch_size: int,
    ) -> ImportResult:
        """Execute insert-like loading within savepoints and roll everything back.

        Args:
            session: Active SQLAlchemy session.
            model_class: Target ORM model class.
            rows: Rows to test for insertion.
            key_columns: Unused by dry-run strategy.
            batch_size: Batch size for flush boundaries.

        Returns:
            Dry-run execution result with non-persistent counts.
        """

        del key_columns

        result = ImportResult()
        for batch in _chunk(rows, batch_size):
            savepoint = session.begin_nested()
            try:
                objects = [model_class(**row) for row in batch]
                session.add_all(objects)
                session.flush()
                result.inserted += len(objects)
            except IntegrityError as exc:
                result.failed += len(batch)
                result.errors.append(str(exc))
            except Exception as exc:
                result.failed += len(batch)
                result.errors.append(str(exc))
            finally:
                if savepoint.is_active:
                    savepoint.rollback()

        return result
