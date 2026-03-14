"""Validation orchestration for Excel sources."""

from __future__ import annotations

from typing import TYPE_CHECKING, BinaryIO, Protocol, cast

from sqlalchemy_excel.reader.base import normalize_header
from sqlalchemy_excel.reader.excel_dbapi_reader import ExcelDbapiReader
from sqlalchemy_excel.validation.pydantic_backend import PydanticBackend
from sqlalchemy_excel.validation.report import CellError, ValidationReport

if TYPE_CHECKING:
    from collections.abc import Iterable, Mapping
    from pathlib import Path

    from sqlalchemy_excel.mapping import ExcelMapping


class _ReaderResult(Protocol):
    headers: list[str]
    rows: Iterable[Mapping[str, object]]
    total_rows: int | None


class _ReaderProtocol(Protocol):
    def read(
        self,
        source: str | Path | BinaryIO,
        sheet_name: str | None = None,
        header_row: int = 1,
    ) -> _ReaderResult: ...


class ExcelValidator:
    """Orchestrate row-level validation for Excel files."""

    def __init__(
        self, mappings: list[ExcelMapping], *, backend: str = "pydantic"
    ) -> None:
        """Initialize validator with mapping metadata and backend selection.

        Args:
            mappings: Mapping definitions, one per worksheet schema.
            backend: Validation backend identifier. Currently supports ``pydantic``.
        """

        if not mappings:
            raise ValueError("At least one mapping is required")
        if backend != "pydantic":
            raise ValueError("Only 'pydantic' backend is currently supported")

        self._mappings: list[ExcelMapping] = mappings
        self._backends: dict[str, PydanticBackend] = {
            mapping.sheet_name: PydanticBackend(mapping) for mapping in mappings
        }
        self._reader: _ReaderProtocol = _build_reader()

    def validate(
        self,
        source: str | Path | BinaryIO,
        *,
        sheet_name: str | None = None,
        max_errors: int | None = None,
        stop_on_first_error: bool = False,
    ) -> ValidationReport:
        """Validate Excel rows and return a structured report.

        Args:
            source: Input Excel file path or binary stream.
            sheet_name: Optional explicit worksheet name.
            max_errors: Optional cap on the number of collected errors.
            stop_on_first_error: Stop processing after first invalid row.

        Returns:
            A ``ValidationReport`` containing row counts and cell-level errors.
        """

        validate_all_mappings = sheet_name is None and len(self._mappings) > 1
        if validate_all_mappings:
            mappings_to_validate = self._mappings
        else:
            mappings_to_validate = [_select_mapping(self._mappings, sheet_name)]

        errors: list[CellError] = []
        total_rows = 0
        invalid_rows = 0

        for mapping in mappings_to_validate:
            _reset_source_cursor(source)
            selected_sheet = mapping.sheet_name if validate_all_mappings else sheet_name
            reader_result = self._reader.read(
                source, sheet_name=selected_sheet, header_row=1
            )
            backend = self._backends[mapping.sheet_name]
            header_map = _build_header_map(mapping, reader_result.headers)

            processed_rows = 0
            mapping_invalid_rows = 0
            for offset, raw_row in enumerate(reader_result.rows, start=2):
                processed_rows += 1
                row_data = _remap_row(raw_row, header_map, mapping)
                row_errors = backend.validate_row(row_data, offset)
                if not row_errors:
                    continue

                mapping_invalid_rows += 1
                errors.extend(row_errors)

                if max_errors is not None and len(errors) >= max_errors:
                    errors = errors[:max_errors]
                    break
                if stop_on_first_error:
                    break

            mapping_total_rows = (
                reader_result.total_rows
                if reader_result.total_rows is not None
                else processed_rows
            )
            total_rows += mapping_total_rows
            invalid_rows += mapping_invalid_rows

            if stop_on_first_error and mapping_invalid_rows > 0:
                break
            if max_errors is not None and len(errors) >= max_errors:
                break

        valid_rows = max(total_rows - invalid_rows, 0)
        return ValidationReport(
            errors=errors,
            total_rows=total_rows,
            valid_rows=valid_rows,
            invalid_rows=invalid_rows,
        )


def _select_mapping(
    mappings: list[ExcelMapping], sheet_name: str | None
) -> ExcelMapping:
    if sheet_name is None:
        return mappings[0]

    for mapping in mappings:
        if mapping.sheet_name == sheet_name:
            return mapping

    available = [mapping.sheet_name for mapping in mappings]
    raise ValueError(
        f"No mapping found for sheet '{sheet_name}'. Available: {available}"
    )


def _build_reader() -> _ReaderProtocol:
    return cast("_ReaderProtocol", cast("object", ExcelDbapiReader(read_only=True)))


def _build_header_map(mapping: ExcelMapping, headers: list[str]) -> dict[str, str]:
    normalized_to_column: dict[str, str] = {}
    for column in mapping.columns:
        normalized_to_column[normalize_header(column.name)] = column.name
        normalized_to_column[normalize_header(column.excel_header)] = column.name

    return {
        header: normalized_to_column.get(normalize_header(header), header)
        for header in headers
        if header
    }


def _remap_row(
    raw_row: Mapping[str, object],
    header_map: dict[str, str],
    mapping: ExcelMapping,
) -> dict[str, object]:
    remapped: dict[str, object] = {column.name: None for column in mapping.columns}
    for header, value in raw_row.items():
        target_column = header_map.get(header)
        if target_column in remapped:
            remapped[target_column] = value
    return remapped


def _reset_source_cursor(source: str | Path | BinaryIO) -> None:
    seek = getattr(source, "seek", None)
    if callable(seek):
        _ = seek(0)
