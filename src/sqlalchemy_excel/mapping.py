from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from typing import TYPE_CHECKING, cast

from sqlalchemy import inspect as sa_inspect
from sqlalchemy.exc import NoInspectionAvailable
from sqlalchemy.sql.sqltypes import (
    Boolean,
    Date,
    DateTime,
    Float,
    Integer,
    Numeric,
    String,
    Text,
)
from sqlalchemy.sql.sqltypes import (
    Enum as SAEnum,
)

from sqlalchemy_excel.exceptions import MappingError

if TYPE_CHECKING:
    from sqlalchemy.orm import DeclarativeBase
    from sqlalchemy.sql.schema import Column
    from sqlalchemy.types import TypeEngine

    from sqlalchemy_excel._types import ColumnName

_TYPE_MAP: tuple[tuple[type[object], type[object]], ...] = (
    (Integer, int),
    (String, str),
    (Text, str),
    (Float, float),
    (Boolean, bool),
    (Date, date),
    (DateTime, datetime),
)


@dataclass(frozen=True)
class ColumnMapping:
    """Mapping metadata for a single ORM column.

    Attributes:
        name: ORM column name.
        excel_header: Display name in the Excel header.
        python_type: Inferred Python type used for validation/coercion.
        sqla_type: SQLAlchemy type object for the column.
        nullable: Whether the column accepts NULL values.
        primary_key: Whether the column is part of the primary key.
        has_default: Whether the column defines Python/server default.
        default_value: Static default value if available.
        enum_values: Enum value options for dropdown validation.
        max_length: Maximum length for constrained string types.
        description: Optional description from column doc/comment.
        foreign_key: Referenced ``table.column`` for foreign keys.
    """

    name: str
    excel_header: str
    python_type: type[object]
    sqla_type: TypeEngine[object]
    nullable: bool
    primary_key: bool
    has_default: bool
    default_value: object | None = None
    enum_values: list[str] | None = None
    max_length: int | None = None
    description: str | None = None
    foreign_key: str | None = None


@dataclass(frozen=True)
class ExcelMapping:
    """Complete mapping from an ORM model to Excel structure.

    Attributes:
        model_class: SQLAlchemy ORM model class.
        sheet_name: Target worksheet name.
        columns: Ordered column mappings.
        key_columns: Column names used as logical keys for upsert behavior.
    """

    model_class: type[DeclarativeBase]
    sheet_name: str
    columns: list[ColumnMapping]
    key_columns: list[str] = field(default_factory=list)

    @classmethod
    def from_model(
        cls,
        model: type[DeclarativeBase],
        *,
        sheet_name: str | None = None,
        key_columns: list[ColumnName] | None = None,
        include: list[ColumnName] | None = None,
        exclude: list[ColumnName] | None = None,
        header_map: dict[ColumnName, str] | None = None,
    ) -> ExcelMapping:
        """Create an ``ExcelMapping`` by introspecting a SQLAlchemy model.

        Args:
            model: SQLAlchemy ORM declarative model class.
            sheet_name: Optional worksheet name. Defaults to ``__tablename__``.
            key_columns: Optional key columns. Defaults to primary keys.
            include: Optional allowlist of column names.
            exclude: Optional denylist of column names.
            header_map: Optional overrides for generated Excel headers.

        Returns:
            A fully-populated ``ExcelMapping`` instance.

        Raises:
            MappingError: If introspection fails or invalid options are provided.
        """

        if include is not None and exclude is not None:
            raise MappingError("'include' and 'exclude' are mutually exclusive")

        try:
            mapper = sa_inspect(model)
        except NoInspectionAvailable as exc:
            raise MappingError(f"Could not inspect model {model!r}") from exc

        all_columns = list(mapper.columns)
        if not all_columns:
            raise MappingError(f"Model {model.__name__!r} has no mapped columns")

        include_set: set[ColumnName] | None = (
            set(include) if include is not None else None
        )
        exclude_set: set[ColumnName] = set(exclude) if exclude is not None else set()
        header_overrides = header_map or {}

        filtered_columns: list[Column[object]] = []
        for column in all_columns:
            if include_set is not None and column.name not in include_set:
                continue
            if column.name in exclude_set:
                continue
            filtered_columns.append(column)

        if not filtered_columns:
            raise MappingError(
                f"No columns selected for model {model.__name__!r} after filtering"
            )

        mappings = [
            _column_to_mapping(column, header_overrides=header_overrides)
            for column in filtered_columns
        ]

        selected_names = {column.name for column in filtered_columns}
        resolved_keys: list[str]
        if key_columns is None:
            resolved_keys = [column.name for column in filtered_columns if column.primary_key]
        else:
            resolved_keys = list(key_columns)

        missing_keys = [name for name in resolved_keys if name not in selected_names]
        if missing_keys:
            missing = ", ".join(missing_keys)
            raise MappingError(f"Unknown key columns for model {model.__name__!r}: {missing}")

        resolved_sheet_name = sheet_name
        if resolved_sheet_name is None:
            table_name = getattr(model, "__tablename__", None)
            if isinstance(table_name, str):
                resolved_sheet_name = table_name
            else:
                resolved_sheet_name = model.__name__

        return cls(
            model_class=model,
            sheet_name=resolved_sheet_name,
            columns=mappings,
            key_columns=resolved_keys,
        )


def _column_to_mapping(
    column: Column[object],
    *,
    header_overrides: dict[ColumnName, str],
) -> ColumnMapping:
    sqla_type = column.type
    python_type = _python_type_for_sqla_type(sqla_type)
    has_default, default_value = _extract_default(column)

    enum_values: list[str] | None = None
    if isinstance(sqla_type, SAEnum):
        enum_values = _extract_enum_values(sqla_type)

    max_length: int | None = None
    if isinstance(sqla_type, String):
        max_length = cast("int | None", getattr(sqla_type, "length", None))

    fk = next(iter(column.foreign_keys), None)
    foreign_key = cast("str", fk.target_fullname) if fk is not None else None

    description = column.doc or column.comment

    header = header_overrides.get(column.name, _default_excel_header(column.name))

    return ColumnMapping(
        name=column.name,
        excel_header=header,
        python_type=python_type,
        sqla_type=sqla_type,
        nullable=bool(column.nullable),
        primary_key=column.primary_key,
        has_default=has_default,
        default_value=default_value,
        enum_values=enum_values,
        max_length=max_length,
        description=description,
        foreign_key=foreign_key,
    )


def _python_type_for_sqla_type(sqla_type: TypeEngine[object]) -> type[object]:
    if isinstance(sqla_type, Numeric):
        as_decimal = cast("bool", getattr(cast("object", sqla_type), "asdecimal", False))
        return Decimal if as_decimal else float

    for sa_type, py_type in _TYPE_MAP:
        if isinstance(sqla_type, sa_type):
            return py_type

    return str


def _extract_default(column: Column[object]) -> tuple[bool, object | None]:
    default_obj = column.default
    if default_obj is not None:
        value_obj: object | None = getattr(default_obj, "arg", None)
        if callable(value_obj):
            return True, None
        return True, value_obj

    server_default = column.server_default
    if server_default is not None:
        server_value: object | None = getattr(server_default, "arg", None)
        return True, None if server_value is None else str(server_value)

    return False, None


def _extract_enum_values(enum_type: SAEnum) -> list[str]:
    enum_class = enum_type.enum_class
    if enum_class is not None:
        return [str(cast("object", member.value)) for member in enum_class]

    values = cast("list[object]", enum_type.enums)
    return [str(value) for value in values]


def _default_excel_header(column_name: str) -> str:
    return column_name.replace("_", " ").title()
