"""Excel import orchestration for insert, upsert, and dry-run operations."""

from __future__ import annotations

from collections.abc import Iterable, Sequence
from time import perf_counter
from typing import TYPE_CHECKING, Protocol, cast

from sqlalchemy_excel.exceptions import ImportError_
from sqlalchemy_excel.load.strategies import (
    DryRunStrategy,
    ImportResult,
    InsertStrategy,
    LoadStrategy,
    UpsertStrategy,
)

if TYPE_CHECKING:
    from sqlalchemy.orm import Session

    from sqlalchemy_excel._types import FileSource
    from sqlalchemy_excel.mapping import ExcelMapping

RowDict = dict[str, object]


class _ValidationReportLike(Protocol):
    """Protocol for validation reports consumed by importer."""

    has_errors: bool
    errors: Sequence[object]
    invalid_rows: int


class _ValidatorLike(Protocol):
    """Protocol for validator implementation consumed by importer."""

    def validate(self, source: FileSource) -> _ValidationReportLike:
        """Validate source and return report object."""

        ...


class _ReaderResultLike(Protocol):
    """Protocol for reader result objects exposing row iterables."""

    rows: Iterable[RowDict]


class _ReaderLike(Protocol):
    """Protocol for reader implementation consumed by importer."""

    def read(
        self,
        source: FileSource,
        sheet_name: str | None = None,
    ) -> _ReaderResultLike | Iterable[RowDict]:
        """Read an Excel source and return row data."""

        ...


class ExcelImporter:
    """Import validated Excel rows into a database via SQLAlchemy ORM.

    Args:
        mappings: Model-to-sheet mappings used for reading and loading rows.
        session: Active SQLAlchemy session. Transaction boundaries are managed by caller.
    """

    def __init__(self, mappings: list[ExcelMapping], session: Session) -> None:
        if not mappings:
            raise ImportError_("At least one ExcelMapping is required")

        self._mappings: list[ExcelMapping] = mappings
        self._session: Session = session

    def insert(
        self,
        source: FileSource,
        *,
        batch_size: int = 1000,
        validate: bool = True,
    ) -> ImportResult:
        """Insert rows from Excel source into target models.

        Args:
            source: Excel file source.
            batch_size: Batch size used by insert strategy.
            validate: Whether to validate source before importing.

        Returns:
            Aggregate import result.
        """

        return self._run(
            source=source,
            strategy=InsertStrategy(),
            batch_size=batch_size,
            validate=validate,
        )

    def upsert(
        self,
        source: FileSource,
        *,
        batch_size: int = 1000,
        validate: bool = True,
    ) -> ImportResult:
        """Upsert rows from Excel source into target models.

        Args:
            source: Excel file source.
            batch_size: Batch size used by upsert strategy.
            validate: Whether to validate source before importing.

        Returns:
            Aggregate import result.
        """

        return self._run(
            source=source,
            strategy=UpsertStrategy(),
            batch_size=batch_size,
            validate=validate,
        )

    def dry_run(self, source: FileSource, *, validate: bool = True) -> ImportResult:
        """Validate and simulate inserts without persisting any data.

        Args:
            source: Excel file source.
            validate: Whether to validate source before loading.

        Returns:
            Aggregate dry-run result.
        """

        return self._run(
            source=source,
            strategy=DryRunStrategy(),
            batch_size=1000,
            validate=validate,
        )

    def _run(
        self,
        *,
        source: FileSource,
        strategy: LoadStrategy,
        batch_size: int,
        validate: bool,
    ) -> ImportResult:
        """Run an import strategy end-to-end.

        Args:
            source: Excel file source.
            strategy: Load strategy to execute.
            batch_size: Batch size for strategy execution.
            validate: Whether to validate before loading.

        Returns:
            Aggregate import result including duration.
        """

        start = perf_counter()
        result = ImportResult()

        if validate:
            validation_failure = self._validate_source(source)
            if validation_failure is not None:
                validation_failure.duration_ms = (perf_counter() - start) * 1000
                return validation_failure

        reader = self._create_reader()

        for mapping in self._mappings:
            self._reset_source_cursor(source)

            try:
                read_result = reader.read(source, sheet_name=mapping.sheet_name)
                rows = self._extract_rows_for_mapping(read_result, mapping)
                partial = strategy.execute(
                    session=self._session,
                    model_class=mapping.model_class,
                    rows=rows,
                    key_columns=mapping.key_columns,
                    batch_size=batch_size,
                )
                self._merge_result(result, partial)
            except Exception as exc:
                result.failed += 1
                result.errors.append(str(exc))

        result.duration_ms = (perf_counter() - start) * 1000
        return result

    def _validate_source(self, source: FileSource) -> ImportResult | None:
        """Validate source using ``ExcelValidator`` if available.

        Args:
            source: Excel file source.

        Returns:
            ``None`` when validation passes, otherwise an ``ImportResult`` containing
            validation errors.
        """

        validator = self._create_validator()
        self._reset_source_cursor(source)
        report = validator.validate(source)

        if not report.has_errors:
            return None

        errors = [str(error) for error in report.errors]
        invalid_rows = report.invalid_rows
        return ImportResult(
            inserted=0,
            updated=0,
            skipped=0,
            failed=invalid_rows,
            errors=errors,
            duration_ms=0.0,
        )

    def _create_validator(self) -> _ValidatorLike:
        """Create an ``ExcelValidator`` instance for importer mappings.

        Returns:
            Validator instance.

        Raises:
            ImportError_: If the validation module is unavailable.
        """

        try:
            from sqlalchemy_excel.validation.engine import ExcelValidator
        except ModuleNotFoundError as exc:
            raise ImportError_(
                "ExcelValidator is unavailable. "
                + "Implement sqlalchemy_excel.validation.engine first."
            ) from exc

        validator = ExcelValidator(self._mappings)
        return cast("_ValidatorLike", cast("object", validator))

    def _create_reader(self) -> _ReaderLike:
        """Create an ``OpenpyxlReader`` instance for reading workbook rows.

        Returns:
            Reader instance with a ``read`` method.

        Raises:
            ImportError_: If the reader module is unavailable.
        """

        try:
            from sqlalchemy_excel.reader.openpyxl_reader import OpenpyxlReader
        except ModuleNotFoundError as exc:
            raise ImportError_(
                "OpenpyxlReader is unavailable. "
                + "Implement sqlalchemy_excel.reader.openpyxl_reader first."
            ) from exc

        reader = OpenpyxlReader()
        return cast("_ReaderLike", cast("object", reader))

    def _extract_rows_for_mapping(
        self,
        read_result: _ReaderResultLike | Iterable[RowDict],
        mapping: ExcelMapping,
    ) -> list[RowDict]:
        """Extract model-aligned rows from a reader result.

        Args:
            read_result: Reader return value (ReaderResult or iterable of dicts).
            mapping: Mapping used to align headers to model columns.

        Returns:
            List of dict rows keyed by model column names.
        """

        raw_rows: Iterable[RowDict]
        if isinstance(read_result, Iterable) and not hasattr(read_result, "rows"):
            raw_rows = read_result
        else:
            raw_rows = cast("_ReaderResultLike", read_result).rows

        prepared_rows: list[RowDict] = []
        for raw_row in raw_rows:
            prepared_rows.append(self._align_row(raw_row, mapping))
        return prepared_rows

    @staticmethod
    def _align_row(row: RowDict, mapping: ExcelMapping) -> RowDict:
        """Align a raw reader row with ORM column names from ``ExcelMapping``.

        Args:
            row: Raw row dictionary from reader backend.
            mapping: Mapping providing column names and Excel headers.

        Returns:
            Row dictionary keyed by ORM column names.
        """

        aligned: RowDict = {}
        for column in mapping.columns:
            if column.name in row:
                aligned[column.name] = row[column.name]
            else:
                aligned[column.name] = row.get(column.excel_header)
        return aligned

    @staticmethod
    def _merge_result(target: ImportResult, partial: ImportResult) -> None:
        """Merge partial strategy result into aggregate result.

        Args:
            target: Aggregate result to update.
            partial: Partial result from one strategy execution.
        """

        target.inserted += partial.inserted
        target.updated += partial.updated
        target.skipped += partial.skipped
        target.failed += partial.failed
        target.errors.extend(partial.errors)

    @staticmethod
    def _reset_source_cursor(source: FileSource) -> None:
        """Reset file-like source cursor to the beginning when possible.

        Args:
            source: File path or binary stream.
        """

        seek = getattr(source, "seek", None)
        if callable(seek):
            _ = seek(0)
