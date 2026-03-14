"""Pydantic v2-backed row validator for Excel data."""

from __future__ import annotations

import enum
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING, Literal, cast

from pydantic import BaseModel, Field, create_model
from pydantic import ValidationError as PydanticValidationError

from sqlalchemy_excel.validation.report import CellError

if TYPE_CHECKING:
    from sqlalchemy_excel.mapping import ColumnMapping, ExcelMapping

_TYPE_ERROR_CODES = {
    "int_parsing",
    "int_type",
    "float_parsing",
    "float_type",
    "decimal_parsing",
    "decimal_type",
    "bool_parsing",
    "bool_type",
    "date_from_datetime_parsing",
    "date_parsing",
    "date_type",
    "datetime_parsing",
    "datetime_type",
    "string_type",
}


class PydanticBackend:
    """Validate rows using a dynamic Pydantic model generated from ``ExcelMapping``."""

    def __init__(self, mapping: ExcelMapping) -> None:
        """Initialize backend and compile a row-validation model.

        Args:
            mapping: Schema metadata extracted from a SQLAlchemy model.
        """

        self._mapping: ExcelMapping = mapping
        self._columns: dict[str, ColumnMapping] = {
            column.name: column for column in mapping.columns
        }
        self._model: type[BaseModel] = self._create_pydantic_model(mapping)

    def validate_row(self, row: dict[str, object], row_number: int) -> list[CellError]:
        """Validate one row and return all cell errors for that row.

        Args:
            row: Raw row keyed by canonical column name.
            row_number: Excel row number (1-based, including header row offset).

        Returns:
            A list of ``CellError`` values. Empty when the row is valid.
        """

        prepared_row = {
            column.name: self._coerce_value(row.get(column.name), column)
            for column in self._mapping.columns
        }

        try:
            _ = self._model.model_validate(prepared_row)
        except PydanticValidationError as exc:
            errors: list[CellError] = []
            for issue in exc.errors():
                loc = issue.get("loc", ())
                column_name = str(loc[0]) if loc else "<row>"
                column = self._columns.get(column_name)
                value = row.get(column_name)
                expected_type = _expected_type(column)
                error_type = str(issue.get("type", ""))
                error_code = _map_error_code(
                    error_type=error_type, value=value, column=column
                )
                errors.append(
                    CellError(
                        row=row_number,
                        column=column_name,
                        value=value,
                        expected_type=expected_type,
                        message=str(issue.get("msg", "Validation failed")),
                        error_code=error_code,
                    )
                )
            return errors

        return []

    def _create_pydantic_model(self, mapping: ExcelMapping) -> type[BaseModel]:
        """Create a Pydantic model from column metadata."""

        field_definitions: dict[str, tuple[object, object]] = {}

        for column in mapping.columns:
            field_type = _field_type_for_column(column)
            if column.max_length is not None:
                if column.nullable:
                    field_info: object = cast(
                        "object",
                        Field(default=None, max_length=column.max_length),
                    )
                else:
                    field_info = cast(
                        "object",
                        Field(default=..., max_length=column.max_length),
                    )
            else:
                if column.nullable:
                    field_info = cast("object", Field(default=None))
                else:
                    field_info = cast("object", Field(default=...))

            field_definitions[column.name] = (field_type, field_info)

        model_name = f"{mapping.model_class.__name__}Validator"
        return cast(
            "type[BaseModel]",
            create_model(model_name, **field_definitions),  # type: ignore[call-overload]  # pyright: ignore[reportCallIssue,reportArgumentType]
        )

    def _coerce_value(self, value: object, column: ColumnMapping) -> object:
        """Apply lightweight coercion before strict Pydantic validation."""

        if value == "":
            value = None

        if value is None:
            return None

        if column.enum_values is not None:
            if isinstance(value, enum.Enum):
                return cast("object", value.value)
            return str(value)

        target_type = column.python_type

        try:
            if target_type is bool and isinstance(value, str):
                normalized = value.strip().lower()
                if normalized in {"true", "1", "yes", "y", "on"}:
                    return True
                if normalized in {"false", "0", "no", "n", "off"}:
                    return False
                return value

            if target_type is int and isinstance(value, str):
                return int(value.strip())

            if target_type is float and isinstance(value, str):
                return float(value.strip())

            if target_type is Decimal and isinstance(value, str):
                return Decimal(value.strip())

            if target_type is date and isinstance(value, str):
                return date.fromisoformat(value.strip())

            if target_type is datetime and isinstance(value, str):
                return datetime.fromisoformat(value.strip())

            if target_type is str:
                return str(value)
        except (TypeError, ValueError, InvalidOperation):
            return value

        return value


def _field_type_for_column(column: ColumnMapping) -> object:
    """Resolve dynamic Pydantic field type for a column mapping."""

    if column.enum_values:
        field_type = cast("object", Literal[tuple(column.enum_values)])
    else:
        field_type = column.python_type

    if column.nullable:
        return cast("object", field_type | None)  # type: ignore[operator]  # pyright: ignore[reportOperatorIssue]
    return field_type


def _expected_type(column: ColumnMapping | None) -> str:
    """Return human-readable expected type text for error reports."""

    if column is None:
        return "unknown"

    if column.enum_values:
        expected = f"one of {column.enum_values}"
    else:
        expected = column.python_type.__name__

    if column.max_length is not None:
        expected = f"{expected} (max length {column.max_length})"
    if column.nullable:
        expected = f"{expected} | None"
    return expected


def _map_error_code(
    *, error_type: str, value: object, column: ColumnMapping | None
) -> str:
    """Map Pydantic error details to project-level error codes."""

    if value is None and column is not None and not column.nullable:
        return "null_error"
    if error_type == "missing":
        return "null_error"
    if error_type.startswith("string_too_long"):
        return "length_error"
    if error_type == "literal_error":
        return "enum_error"
    if error_type in _TYPE_ERROR_CODES or error_type.endswith("_parsing"):
        return "type_error"
    return "constraint_error"
