from __future__ import annotations

# pyright: reportImplicitRelativeImport=none, reportMissingImports=none, reportUnknownVariableType=none, reportUnknownMemberType=none, reportUnknownParameterType=none, reportUnknownArgumentType=none, reportAny=none
import pytest
from tests.conftest import (
    Base,
    Department,
    Employee,
    EmployeeStatus,
    Product,
    SimpleUser,
)

from sqlalchemy_excel.exceptions import MappingError
from sqlalchemy_excel.mapping import ExcelMapping


def _get_column(mapping: ExcelMapping, name: str):
    return next(column for column in mapping.columns if column.name == name)


def test_from_model_basic() -> None:
    mapping = ExcelMapping.from_model(SimpleUser)

    assert mapping.sheet_name == "simple_users"
    assert len(mapping.columns) == 4
    assert mapping.key_columns == ["id"]


def test_column_mapping_types() -> None:
    mapping = ExcelMapping.from_model(SimpleUser)

    assert _get_column(mapping, "id").python_type is int
    assert _get_column(mapping, "name").python_type is str
    assert _get_column(mapping, "email").python_type is str
    assert _get_column(mapping, "age").python_type is int


def test_column_mapping_nullable() -> None:
    mapping = ExcelMapping.from_model(SimpleUser)

    assert _get_column(mapping, "age").nullable is True
    assert _get_column(mapping, "name").nullable is False


def test_column_mapping_primary_key() -> None:
    mapping = ExcelMapping.from_model(SimpleUser)

    primary_key_columns = [
        column.name for column in mapping.columns if column.primary_key
    ]
    assert primary_key_columns == ["id"]


def test_column_mapping_max_length() -> None:
    mapping = ExcelMapping.from_model(SimpleUser)

    assert _get_column(mapping, "name").max_length == 100
    assert _get_column(mapping, "email").max_length == 255


def test_from_model_with_enum() -> None:
    mapping = ExcelMapping.from_model(Employee)

    status_column = _get_column(mapping, "status")
    expected_values = [member.value for member in EmployeeStatus]
    assert status_column.enum_values == expected_values


def test_from_model_with_foreign_key() -> None:
    mapping = ExcelMapping.from_model(Employee)

    department_id_column = _get_column(mapping, "department_id")
    assert department_id_column.foreign_key == f"{Department.__tablename__}.id"


def test_from_model_with_defaults() -> None:
    mapping = ExcelMapping.from_model(Employee)

    assert _get_column(mapping, "status").has_default is True
    assert _get_column(mapping, "hire_date").has_default is True


def test_excel_header_generation() -> None:
    mapping = ExcelMapping.from_model(Employee)

    assert _get_column(mapping, "first_name").excel_header == "First Name"


def test_include_columns() -> None:
    mapping = ExcelMapping.from_model(SimpleUser, include=["id", "name"])

    assert [column.name for column in mapping.columns] == ["id", "name"]
    assert len(mapping.columns) == 2


def test_exclude_columns() -> None:
    mapping = ExcelMapping.from_model(SimpleUser, exclude=["age"])

    column_names = [column.name for column in mapping.columns]
    assert len(mapping.columns) == 3
    assert "age" not in column_names


def test_include_exclude_mutually_exclusive() -> None:
    with pytest.raises(MappingError, match="mutually exclusive"):
        ExcelMapping.from_model(SimpleUser, include=["id"], exclude=["email"])


def test_header_map_override() -> None:
    mapping = ExcelMapping.from_model(SimpleUser, header_map={"name": "Full Name"})

    assert _get_column(mapping, "name").excel_header == "Full Name"


def test_custom_sheet_name() -> None:
    mapping = ExcelMapping.from_model(SimpleUser, sheet_name="Users")

    assert mapping.sheet_name == "Users"


def test_custom_key_columns() -> None:
    mapping = ExcelMapping.from_model(SimpleUser, key_columns=["email"])

    assert mapping.key_columns == ["email"]


def test_invalid_key_columns() -> None:
    with pytest.raises(MappingError, match="Unknown key columns"):
        ExcelMapping.from_model(SimpleUser, key_columns=["nonexistent"])


def test_non_model_raises_error() -> None:
    with pytest.raises(MappingError, match="Could not inspect model"):
        ExcelMapping.from_model(str)


def test_models_from_conftest_are_real_sqlalchemy_models() -> None:
    assert issubclass(SimpleUser, Base)
    assert issubclass(Employee, Base)
    assert issubclass(Department, Base)
    assert issubclass(Product, Base)
