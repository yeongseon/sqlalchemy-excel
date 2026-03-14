# TDD.md — sqlalchemy-excel Technical Design Document

> Version: 0.1.0-draft | Date: 2026-03-14

## 1. Implementation Plan

### Phase 1: Foundation (Week 1–2)

```
1. Project skeleton (pyproject.toml, src layout, CI)
2. Exception hierarchy (exceptions.py)
3. Type aliases (_types.py)
4. Compatibility helpers (_compat.py)
5. mapping.py (ORM introspection → ExcelMapping)
6. Unit tests for mapping
```

### Phase 2: Template & Reader (Week 3–4)

```
1. template.py (ExcelMapping → .xlsx)
2. reader/base.py (Protocol/ABC)
3. reader/openpyxl_reader.py
4. reader/pandas_reader.py (optional)
5. Unit tests for template and reader
```

### Phase 3: Validation (Week 5–6)

```
1. validation/report.py (CellError, ValidationReport)
2. validation/pydantic_backend.py
3. validation/engine.py (orchestrator)
4. validation/pandera_backend.py (optional)
5. Unit tests for validation
```

### Phase 4: Import & Export (Week 7–8)

```
1. load/strategies.py (Insert, Upsert)
2. load/importer.py (ExcelImporter)
3. export.py (ExcelExporter)
4. Integration tests (full pipeline)
```

### Phase 5: CLI & Integration (Week 9–10)

```
1. cli.py (Click commands)
2. integrations/fastapi.py (router factory)
3. __init__.py (public API re-exports)
4. End-to-end tests
5. Documentation, README, examples
```

## 2. Detailed Module Design

### 2.1 `exceptions.py`

```python
"""Custom exception hierarchy for sqlalchemy-excel."""
from __future__ import annotations


class SqlalchemyExcelError(Exception):
    """Base exception for all sqlalchemy-excel errors."""


class MappingError(SqlalchemyExcelError):
    """Raised when ORM model introspection fails.

    Examples:
        - Model has no mapped columns
        - Unsupported column type encountered
        - Ambiguous column mapping
    """


class TemplateError(SqlalchemyExcelError):
    """Raised when Excel template generation fails."""


class ReaderError(SqlalchemyExcelError):
    """Base exception for Excel file reading errors."""


class FileFormatError(ReaderError):
    """Raised when input file is not a valid .xlsx file."""


class SheetNotFoundError(ReaderError):
    """Raised when specified sheet name doesn't exist."""

    def __init__(self, sheet_name: str, available: list[str]) -> None:
        self.sheet_name = sheet_name
        self.available = available
        super().__init__(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {', '.join(available)}"
        )


class HeaderMismatchError(ReaderError):
    """Raised when Excel headers don't match expected columns."""

    def __init__(
        self,
        missing: list[str],
        extra: list[str],
    ) -> None:
        self.missing = missing
        self.extra = extra
        parts = []
        if missing:
            parts.append(f"Missing columns: {', '.join(missing)}")
        if extra:
            parts.append(f"Unexpected columns: {', '.join(extra)}")
        super().__init__(". ".join(parts))


class ValidationError(SqlalchemyExcelError):
    """Raised when data validation fails.

    Contains a ValidationReport with detailed errors.
    """

    def __init__(self, report: object) -> None:  # ValidationReport
        self.report = report
        super().__init__(str(report))


class ImportError_(SqlalchemyExcelError):
    """Raised when database import fails.

    Underscore suffix to avoid shadowing builtin ImportError.
    """


class DuplicateKeyError(ImportError_):
    """Raised when an insert violates a unique constraint."""


class ConstraintViolationError(ImportError_):
    """Raised when data violates a database constraint."""


class ExportError(SqlalchemyExcelError):
    """Raised when Excel export fails."""
```

### 2.2 `_types.py`

```python
"""Internal type aliases for sqlalchemy-excel."""
from __future__ import annotations

from os import PathLike
from typing import Any, BinaryIO, Union

# Path-like types accepted by file operations
FilePath = Union[str, PathLike[str]]

# Source types accepted by readers
FileSource = Union[str, PathLike[str], BinaryIO]

# A single row of data from Excel
RowDict = dict[str, Any]

# Column name
ColumnName = str
```

### 2.3 `_compat.py`

```python
"""Version compatibility and optional dependency helpers."""
from __future__ import annotations

import importlib
from typing import Any


def ensure_defusedxml() -> None:
    """Ensure defusedxml is installed for safe XML processing."""
    try:
        import defusedxml  # noqa: F401
    except ImportError as e:
        raise ImportError(
            "defusedxml is required for processing Excel files safely. "
            "Install it with: pip install defusedxml"
        ) from e


def import_optional(
    module_name: str,
    extra_name: str,
) -> Any:
    """Import an optional dependency, raising a helpful error if missing.

    Args:
        module_name: The module to import (e.g., "pandas").
        extra_name: The pip extra name (e.g., "pandas").

    Returns:
        The imported module.

    Raises:
        ImportError: With installation instructions.
    """
    try:
        return importlib.import_module(module_name)
    except ImportError as e:
        raise ImportError(
            f"{module_name} is required for this feature. "
            f"Install it with: pip install sqlalchemy-excel[{extra_name}]"
        ) from e
```

### 2.4 `mapping.py` — Core Design

```python
"""ORM model → ExcelMapping schema extraction."""
from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import Any, Sequence

from sqlalchemy import inspect as sa_inspect
from sqlalchemy.orm import DeclarativeBase
from sqlalchemy.types import (
    Boolean,
    Date,
    DateTime,
    Enum as SAEnum,
    Float,
    Integer,
    Numeric,
    String,
    Text,
    TypeEngine,
)

from sqlalchemy_excel.exceptions import MappingError


# SQLAlchemy type → Python type mapping
_TYPE_MAP: dict[type[TypeEngine], type] = {
    Integer: int,
    Float: float,
    Numeric: float,
    String: str,
    Text: str,
    Boolean: bool,
    Date: date,
    DateTime: datetime,
}


@dataclass(frozen=True)
class ColumnMapping:
    """Mapping for a single ORM column → Excel column."""
    name: str
    excel_header: str
    python_type: type
    sqla_type: TypeEngine
    nullable: bool
    primary_key: bool
    has_default: bool
    default_value: Any | None = None
    enum_values: list[str] | None = None
    max_length: int | None = None
    description: str | None = None
    foreign_key: str | None = None


@dataclass(frozen=True)
class ExcelMapping:
    """Complete mapping from an ORM model to Excel structure."""
    model_class: type
    sheet_name: str
    columns: list[ColumnMapping]
    key_columns: list[str] = field(default_factory=list)

    @classmethod
    def from_model(
        cls,
        model: type[DeclarativeBase],
        *,
        sheet_name: str | None = None,
        key_columns: list[str] | None = None,
        include: list[str] | None = None,
        exclude: list[str] | None = None,
        header_map: dict[str, str] | None = None,
    ) -> ExcelMapping:
        """Create ExcelMapping by introspecting a SQLAlchemy ORM model.

        Args:
            model: SQLAlchemy ORM model class.
            sheet_name: Sheet name in Excel. Defaults to table name.
            key_columns: Columns used as keys for upsert. Defaults to PKs.
            include: Whitelist of column names. Mutually exclusive with exclude.
            exclude: Blacklist of column names. Mutually exclusive with include.
            header_map: Override display headers {column_name: header_text}.

        Returns:
            ExcelMapping instance.

        Raises:
            MappingError: If model has no columns or introspection fails.
        """
        ...
```

**Introspection Strategy**:
1. `sa_inspect(model)` → `InstanceState` / mapper
2. Iterate `mapper.columns` for column metadata
3. Extract: name, type, nullable, primary_key, default, foreign_keys
4. Map SQLAlchemy types to Python types via `_TYPE_MAP`
5. Detect Enum columns → extract values for dropdown validation
6. Apply include/exclude filters
7. Apply header_map overrides

### 2.5 `template.py` — Template Generation

**Style Constants**:
```python
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
REQUIRED_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
```

**Template Generation Steps**:
1. Create workbook
2. For each ExcelMapping → create sheet
3. Write header row with styles
4. Add column comments (type, constraints)
5. Add DataValidation for enum columns
6. Set column widths
7. Add auto-filter
8. Optionally add sample data row
9. Lock header row (sheet protection optional)

### 2.6 `validation/pydantic_backend.py`

**Dynamic Model Generation**:
```python
def _create_pydantic_model(mapping: ExcelMapping) -> type[BaseModel]:
    """Dynamically create a Pydantic model from ExcelMapping.

    Generates field definitions with proper types, validators,
    and constraints based on the ORM column metadata.
    """
    field_definitions: dict[str, tuple[type, FieldInfo]] = {}

    for col in mapping.columns:
        field_type = col.python_type
        if col.nullable:
            field_type = field_type | None

        field_kwargs: dict[str, Any] = {}
        if col.max_length is not None:
            field_kwargs["max_length"] = col.max_length
        if col.enum_values is not None:
            # Use Literal type for enum validation
            ...

        field_definitions[col.name] = (
            field_type,
            Field(default=None if col.nullable else ..., **field_kwargs),
        )

    return create_model(
        f"{mapping.model_class.__name__}Validator",
        **field_definitions,
    )
```

### 2.7 `load/strategies.py`

**Insert Strategy**:
```python
class InsertStrategy:
    def execute(self, session, model_class, rows, key_columns, batch_size):
        inserted = 0
        for batch in _chunk(rows, batch_size):
            objects = [model_class(**row) for row in batch]
            session.add_all(objects)
            session.flush()
            inserted += len(objects)
        return ImportResult(inserted=inserted, updated=0, ...)
```

**Upsert Strategy**:
```python
class UpsertStrategy:
    def execute(self, session, model_class, rows, key_columns, batch_size):
        inserted = updated = 0
        for batch in _chunk(rows, batch_size):
            for row in batch:
                key_filter = {k: row[k] for k in key_columns}
                existing = session.query(model_class).filter_by(**key_filter).first()
                if existing:
                    for k, v in row.items():
                        setattr(existing, k, v)
                    updated += 1
                else:
                    session.add(model_class(**row))
                    inserted += 1
            session.flush()
        return ImportResult(inserted=inserted, updated=updated, ...)
```

## 3. pyproject.toml Configuration

```toml
[build-system]
requires = ["hatchling"]
build-backend = "hatchling.build"

[project]
name = "sqlalchemy-excel"
version = "0.1.0"
description = "SQLAlchemy model-driven Excel template generation, validation, and database import toolkit"
readme = "README.md"
license = {text = "MIT"}
requires-python = ">=3.10"
authors = [
    {name = "Yeongseon Choe"},
]
keywords = ["sqlalchemy", "excel", "template", "validation", "import", "openpyxl"]
classifiers = [
    "Development Status :: 3 - Alpha",
    "Intended Audience :: Developers",
    "License :: OSI Approved :: MIT License",
    "Programming Language :: Python :: 3",
    "Programming Language :: Python :: 3.10",
    "Programming Language :: Python :: 3.11",
    "Programming Language :: Python :: 3.12",
    "Programming Language :: Python :: 3.13",
    "Topic :: Database",
    "Topic :: Office/Business :: Financial :: Spreadsheet",
    "Typing :: Typed",
]

dependencies = [
    "sqlalchemy>=2.0",
    "openpyxl>=3.1",
    "pydantic>=2.0",
    "defusedxml>=0.7",
    "click>=8.0",
]

[project.optional-dependencies]
pandas = ["pandas>=2.0"]
pandera = ["pandera>=0.18", "pandas>=2.0"]
fastapi = ["fastapi>=0.100", "python-multipart>=0.0.5"]
dev = [
    "pytest>=8.0",
    "pytest-cov>=4.0",
    "hypothesis>=6.0",
    "ruff>=0.4",
    "mypy>=1.10",
    "httpx>=0.27",  # For FastAPI test client
]
all = [
    "sqlalchemy-excel[pandas]",
    "sqlalchemy-excel[pandera]",
    "sqlalchemy-excel[fastapi]",
    "sqlalchemy-excel[dev]",
]

[project.scripts]
sqlalchemy-excel = "sqlalchemy_excel.cli:cli"

[project.urls]
Homepage = "https://github.com/yeongseon/sqlalchemy-excel"
Documentation = "https://github.com/yeongseon/sqlalchemy-excel#readme"
Repository = "https://github.com/yeongseon/sqlalchemy-excel"
Issues = "https://github.com/yeongseon/sqlalchemy-excel/issues"

[tool.hatch.build.targets.wheel]
packages = ["src/sqlalchemy_excel"]

[tool.ruff]
target-version = "py310"
src = ["src"]
line-length = 88

[tool.ruff.lint]
select = [
    "E",    # pycodestyle errors
    "W",    # pycodestyle warnings
    "F",    # pyflakes
    "I",    # isort
    "N",    # pep8-naming
    "UP",   # pyupgrade
    "B",    # flake8-bugbear
    "SIM",  # flake8-simplify
    "TCH",  # flake8-type-checking
    "RUF",  # ruff-specific
]
ignore = ["E501"]  # line length handled by formatter

[tool.ruff.lint.isort]
known-first-party = ["sqlalchemy_excel"]

[tool.mypy]
python_version = "3.10"
strict = true
warn_return_any = true
warn_unused_configs = true

[[tool.mypy.overrides]]
module = "openpyxl.*"
ignore_missing_imports = true

[[tool.mypy.overrides]]
module = "defusedxml.*"
ignore_missing_imports = true

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "-ra -q"

[tool.coverage.run]
source = ["sqlalchemy_excel"]
branch = true

[tool.coverage.report]
exclude_lines = [
    "pragma: no cover",
    "if TYPE_CHECKING:",
    "if __name__",
    "@overload",
]
```

## 4. CI/CD Configuration

### GitHub Actions CI

```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:
    branches: [main]

permissions:
  contents: read

jobs:
  lint:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"
          cache: pip
      - run: pip install ruff mypy
      - run: ruff check src/ tests/
      - run: ruff format --check src/ tests/

  test:
    runs-on: ubuntu-latest
    strategy:
      fail-fast: false
      matrix:
        python-version: ["3.10", "3.11", "3.12", "3.13"]
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: ${{ matrix.python-version }}
          cache: pip
      - run: pip install -e ".[dev]"
      - run: pytest --cov=sqlalchemy_excel --cov-report=xml
      - uses: codecov/codecov-action@v4
        if: matrix.python-version == '3.12'
        with:
          file: ./coverage.xml
```

### GitHub Actions Publish

```yaml
name: Publish

on:
  push:
    tags:
      - "v*"

permissions:
  contents: read
  id-token: write

jobs:
  build-and-publish:
    runs-on: ubuntu-latest
    environment: pypi
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"
          cache: pip
      - run: pip install build
      - run: python -m build
      - uses: pypa/gh-action-pypi-publish@release/v1
```

## 5. Sample ORM Models for Testing

```python
"""Sample models used in tests and examples."""
from __future__ import annotations

import enum
from datetime import date, datetime

from sqlalchemy import ForeignKey, String, Text
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, relationship


class Base(DeclarativeBase):
    pass


class Department(Base):
    __tablename__ = "departments"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(100), unique=True)
    code: Mapped[str] = mapped_column(String(10), unique=True)

    employees: Mapped[list[Employee]] = relationship(back_populates="department")


class EmployeeStatus(enum.Enum):
    ACTIVE = "active"
    INACTIVE = "inactive"
    ON_LEAVE = "on_leave"


class Employee(Base):
    __tablename__ = "employees"

    id: Mapped[int] = mapped_column(primary_key=True)
    email: Mapped[str] = mapped_column(String(255), unique=True)
    first_name: Mapped[str] = mapped_column(String(100))
    last_name: Mapped[str] = mapped_column(String(100))
    status: Mapped[EmployeeStatus] = mapped_column(default=EmployeeStatus.ACTIVE)
    salary: Mapped[float | None] = mapped_column(default=None)
    hire_date: Mapped[date] = mapped_column(default=date.today)
    department_id: Mapped[int | None] = mapped_column(
        ForeignKey("departments.id"), default=None
    )
    notes: Mapped[str | None] = mapped_column(Text, default=None)

    department: Mapped[Department | None] = relationship(back_populates="employees")


class Product(Base):
    __tablename__ = "products"

    id: Mapped[int] = mapped_column(primary_key=True)
    sku: Mapped[str] = mapped_column(String(50), unique=True)
    name: Mapped[str] = mapped_column(String(200))
    price: Mapped[float] = mapped_column()
    quantity: Mapped[int] = mapped_column(default=0)
    is_active: Mapped[bool] = mapped_column(default=True)
    created_at: Mapped[datetime] = mapped_column(default=datetime.now)
```

## 6. Security Implementation Details

### Formula Injection Prevention

```python
# In reader or validation
_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

def sanitize_cell_value(value: Any) -> Any:
    """Prevent formula injection in cell values."""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return "'" + value  # Prefix with apostrophe
    return value
```

### File Size Validation

```python
DEFAULT_MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB

def validate_file_size(
    source: BinaryIO,
    max_size: int = DEFAULT_MAX_FILE_SIZE,
) -> None:
    """Check file size before processing."""
    source.seek(0, 2)  # Seek to end
    size = source.tell()
    source.seek(0)  # Reset
    if size > max_size:
        raise ReaderError(
            f"File size ({size:,} bytes) exceeds maximum "
            f"({max_size:,} bytes)"
        )
```

## 7. API Usage Examples

### Basic Usage

```python
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy_excel import (
    ExcelMapping,
    ExcelTemplate,
    ExcelValidator,
    ExcelImporter,
)
from myapp.models import User

# 1. Create mapping from ORM model
mapping = ExcelMapping.from_model(
    User,
    sheet_name="users",
    key_columns=["email"],
)

# 2. Generate template
template = ExcelTemplate([mapping])
template.save("users_template.xlsx")

# 3. Validate uploaded file
validator = ExcelValidator([mapping])
report = validator.validate("users_upload.xlsx")
if report.has_errors:
    report.to_excel("validation_errors.xlsx")
    print(report.summary())
    raise SystemExit(1)

# 4. Import to database
engine = create_engine("sqlite:///app.db")
with Session(engine) as session:
    importer = ExcelImporter([mapping], session=session)
    result = importer.upsert("users_upload.xlsx")
    session.commit()
    print(f"Inserted: {result.inserted}, Updated: {result.updated}")
```

### FastAPI Integration

```python
from fastapi import FastAPI
from sqlalchemy_excel.integrations.fastapi import create_import_router
from myapp.models import User
from myapp.database import get_session

app = FastAPI()

user_router = create_import_router(
    User,
    prefix="/users",
    tags=["users"],
    session_dependency=get_session,
)
app.include_router(user_router)
```

### CLI Usage

```bash
# Generate template
sqlalchemy-excel template --model myapp.models:User --output users.xlsx

# Validate upload
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx

# Import to DB
sqlalchemy-excel import --model myapp.models:User \
    --input upload.xlsx \
    --db sqlite:///app.db \
    --mode upsert \
    --dry-run

# Export from DB
sqlalchemy-excel export --model myapp.models:User \
    --db sqlite:///app.db \
    --output export.xlsx
```
