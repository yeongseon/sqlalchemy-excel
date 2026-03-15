# sqlalchemy-excel

[![PyPI version](https://img.shields.io/pypi/v/sqlalchemy-excel.svg)](https://pypi.org/project/sqlalchemy-excel/)
[![Python versions](https://img.shields.io/pypi/pyversions/sqlalchemy-excel.svg)](https://pypi.org/project/sqlalchemy-excel/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![CI](https://github.com/yeongseon/sqlalchemy-excel/actions/workflows/ci.yml/badge.svg)](https://github.com/yeongseon/sqlalchemy-excel/actions/workflows/ci.yml)

**SQLAlchemy model-driven Excel template generation, server-side validation, and database import toolkit.**

sqlalchemy-excel treats your SQLAlchemy ORM models as the **single source of truth** for Excel workflows. Define your models once, and let the library handle template generation, upload validation, data import, and query export — all with type safety, structured error reporting, and transaction control.

---

## Table of Contents

- [Key Features](#key-features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Core Concepts](#core-concepts)
  - [ExcelMapping — Schema Extraction](#excelmapping--schema-extraction)
  - [ExcelTemplate — Template Generation](#exceltemplate--template-generation)
  - [ExcelValidator — Upload Validation](#excelvalidator--upload-validation)
  - [ExcelImporter — Database Import](#excelimporter--database-import)
  - [ExcelExporter — Query Export](#excelexporter--query-export)
- [Validation and Error Reporting](#validation-and-error-reporting)
  - [CellError](#cellerror)
  - [ValidationReport](#validationreport)
- [Import Strategies and Results](#import-strategies-and-results)
  - [Insert Mode](#insert-mode)
  - [Upsert Mode](#upsert-mode)
  - [Dry Run Mode](#dry-run-mode)
  - [ImportResult](#importresult)
- [CLI Reference](#cli-reference)
  - [template](#template)
  - [validate](#validate)
  - [import](#import)
  - [export](#export)
  - [inspect](#inspect)
- [FastAPI Integration](#fastapi-integration)
- [Exception Hierarchy](#exception-hierarchy)
- [Security](#security)
- [Configuration Reference](#configuration-reference)
- [Development](#development)
- [License](#license)

---

## Key Features

- **Template Generation** — Create `.xlsx` files with correct headers, column types, dropdown validation, sample data, freeze panes, and auto-filters — all derived from your ORM models.
- **Server-Side Validation** — Parse uploaded Excel files and produce structured, row-level and column-level error reports using Pydantic v2.
- **Database Import** — Load validated data via insert, upsert, or dry-run strategies with batched execution, savepoint-based error recovery, and transaction safety.
- **Query Export** — Export SQLAlchemy query results to formatted Excel files with auto-sized columns, date formatting, and header styling.
- **CLI** — Five commands (`template`, `validate`, `import`, `export`, `inspect`) for scripting and automation.
- **FastAPI Integration** — Router factory that generates template download, upload validation, and import endpoints from a single model class.
- **Type Safe** — Fully typed with `mypy --strict`. No `# type: ignore` without specific error codes.

---

## Installation

### Core (required dependencies only)

```bash
pip install sqlalchemy-excel
```

This installs the core dependencies:

| Package | Version | Purpose |
|---------|---------|---------|
| SQLAlchemy | ≥ 2.0 | ORM schema introspection, database operations |
| openpyxl | ≥ 3.1 | Excel file reading and writing |
| Pydantic | ≥ 2.0 | Row-level type coercion and validation |
| defusedxml | ≥ 0.7 | Secure XML parsing for untrusted Excel files |
| Click | ≥ 8.0 | CLI interface |

### Optional Extras

```bash
# pandas-based reader (alternative to openpyxl reader)
pip install sqlalchemy-excel[pandas]

# Pandera for DataFrame-level cross-column validation
pip install sqlalchemy-excel[pandera]

# FastAPI integration (upload/template endpoints)
pip install sqlalchemy-excel[fastapi]

# Development tools (pytest, ruff, mypy, hypothesis)
pip install sqlalchemy-excel[dev]

# Everything
pip install sqlalchemy-excel[all]
```

### Development Installation

```bash
git clone https://github.com/yeongseon/sqlalchemy-excel.git
cd sqlalchemy-excel
pip install -e ".[all]"
```

---

## Quick Start

Define a SQLAlchemy model, generate a template, validate an upload, and import it — in under 20 lines:

```python
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column

from sqlalchemy_excel import (
    ExcelMapping,
    ExcelTemplate,
    ExcelValidator,
    ExcelImporter,
)


# 1. Define your model (SQLAlchemy 2.0 style)
class Base(DeclarativeBase):
    pass

class User(Base):
    __tablename__ = "users"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column()
    email: Mapped[str] = mapped_column()
    age: Mapped[int | None] = mapped_column(default=None)


# 2. Extract mapping from model
mapping = ExcelMapping.from_model(User)

# 3. Generate a downloadable Excel template
template = ExcelTemplate([mapping], include_sample_data=True)
template.save("users_template.xlsx")

# 4. Validate an uploaded file
validator = ExcelValidator([mapping])
report = validator.validate("users_upload.xlsx")

if report.has_errors:
    print(report.summary())
    for error in report.errors:
        print(f"  Row {error.row}, Column '{error.column}': {error.message}")
else:
    # 5. Import validated data
    engine = create_engine("sqlite:///app.db")
    Base.metadata.create_all(engine)

    with Session(engine) as session:
        importer = ExcelImporter([mapping], session=session)
        result = importer.insert("users_upload.xlsx")
        session.commit()
        print(result.summary())
```

---

## Core Concepts

### ExcelMapping — Schema Extraction

`ExcelMapping` introspects a SQLAlchemy ORM model and produces a structured representation of its columns, types, constraints, and relationships. This mapping drives all downstream operations (template generation, validation, import, export).

#### Creating a Mapping

```python
from sqlalchemy_excel import ExcelMapping

mapping = ExcelMapping.from_model(User)
```

#### `ExcelMapping.from_model()` — Full Signature

```python
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
) -> ExcelMapping
```

**Parameters:**

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `model` | `type[DeclarativeBase]` | *(required)* | SQLAlchemy ORM model class using 2.0-style `DeclarativeBase`. |
| `sheet_name` | `str \| None` | `None` | Override the worksheet name. Defaults to `__tablename__`. |
| `key_columns` | `list[str] \| None` | `None` | Columns used as logical keys for upsert operations. Defaults to primary key columns. |
| `include` | `list[str] \| None` | `None` | Allowlist of column names to include. Mutually exclusive with `exclude`. |
| `exclude` | `list[str] \| None` | `None` | Denylist of column names to exclude. Mutually exclusive with `include`. |
| `header_map` | `dict[str, str] \| None` | `None` | Override generated Excel headers. Keys are column names, values are display headers. |

**Returns:** `ExcelMapping` — A fully-populated mapping instance.

**Raises:** `MappingError` — If introspection fails, the model has no columns, or `include` and `exclude` are both provided.

#### Examples

```python
# Basic: all columns, default settings
mapping = ExcelMapping.from_model(User)

# Custom sheet name
mapping = ExcelMapping.from_model(User, sheet_name="Employee List")

# Only specific columns
mapping = ExcelMapping.from_model(User, include=["name", "email", "age"])

# Exclude auto-generated columns
mapping = ExcelMapping.from_model(User, exclude=["id", "created_at"])

# Custom headers for the Excel file
mapping = ExcelMapping.from_model(
    User,
    header_map={
        "name": "Full Name",
        "email": "Email Address",
    },
)

# Custom key columns for upsert (instead of primary key)
mapping = ExcelMapping.from_model(User, key_columns=["email"])
```

#### ColumnMapping — Per-Column Metadata

Each column in an `ExcelMapping` is represented as a frozen `ColumnMapping` dataclass:

```python
@dataclass(frozen=True)
class ColumnMapping:
    name: str                          # ORM column name
    excel_header: str                  # Display name in Excel header
    python_type: type                  # Inferred Python type (int, str, float, bool, date, datetime, Decimal)
    sqla_type: TypeEngine              # SQLAlchemy type object
    nullable: bool                     # Whether the column accepts NULL
    primary_key: bool                  # Whether the column is a primary key
    has_default: bool                  # Whether the column has a default value
    default_value: object | None       # Static default value (None if callable)
    enum_values: list[str] | None      # Enum options for dropdown validation
    max_length: int | None             # Maximum length for String columns
    description: str | None            # Column doc or comment
    foreign_key: str | None            # Referenced table.column for foreign keys
```

#### Type Inference

sqlalchemy-excel maps SQLAlchemy column types to Python types for validation:

| SQLAlchemy Type | Python Type |
|----------------|-------------|
| `Integer` | `int` |
| `String`, `Text` | `str` |
| `Float` | `float` |
| `Boolean` | `bool` |
| `Date` | `datetime.date` |
| `DateTime` | `datetime.datetime` |
| `Numeric(asdecimal=True)` | `decimal.Decimal` |
| `Numeric(asdecimal=False)` | `float` |
| `Enum` | `str` (with enum_values populated) |
| *(other)* | `str` (fallback) |

---

### ExcelTemplate — Template Generation

`ExcelTemplate` generates formatted `.xlsx` template files from one or more `ExcelMapping` instances. Templates include styled headers, data validation dropdowns, cell comments with type information, freeze panes, and auto-filters.

#### Creating and Saving Templates

```python
from sqlalchemy_excel import ExcelMapping, ExcelTemplate

mapping = ExcelMapping.from_model(User)
template = ExcelTemplate([mapping])

# Save to disk
template.save("users_template.xlsx")

# Get as bytes (for HTTP responses)
xlsx_bytes = template.to_bytes()

# Get as BytesIO stream
stream = template.to_bytesio()
```

#### Constructor

```python
ExcelTemplate(
    mappings: list[ExcelMapping],
    *,
    include_sample_data: bool = False,
)
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `mappings` | `list[ExcelMapping]` | *(required)* | One or more model mappings. Each mapping creates a separate worksheet. |
| `include_sample_data` | `bool` | `False` | If `True`, adds one row of representative sample data below the header. |

#### Methods

| Method | Returns | Description |
|--------|---------|-------------|
| `save(path)` | `None` | Write the template to a `.xlsx` file on disk. |
| `to_bytes()` | `bytes` | Render the template as XLSX bytes (useful for HTTP responses). |
| `to_bytesio()` | `BytesIO` | Render the template as an in-memory `BytesIO` stream. |

**Raises:** `TemplateError` — If workbook generation or saving fails.

#### Template Features

- **Styled headers** — Blue background (`#4472C4`), white bold font, centered alignment.
- **Required column highlighting** — Yellow background (`#FFF2CC`) for non-nullable columns without defaults.
- **Cell comments** — Each header cell includes a comment with type info, constraints, and descriptions.
- **Enum dropdowns** — Columns with `Enum` types get Excel data validation dropdowns.
- **Auto-filter** — Filter arrows on all header columns.
- **Freeze panes** — Header row is frozen for scrolling convenience.
- **Auto-width** — Column widths adjust to header and type hint lengths.
- **Sample data** — Optional representative row showing expected data formats.

#### Multi-Sheet Templates

```python
user_mapping = ExcelMapping.from_model(User)
order_mapping = ExcelMapping.from_model(Order)

# Creates one workbook with two sheets
template = ExcelTemplate([user_mapping, order_mapping])
template.save("data_template.xlsx")
```

---

### ExcelValidator — Upload Validation

`ExcelValidator` parses uploaded Excel files and validates each row against the ORM model's type and constraint definitions using Pydantic v2. It produces a structured `ValidationReport` with cell-level error details.

#### Basic Validation

```python
from sqlalchemy_excel import ExcelMapping, ExcelValidator

mapping = ExcelMapping.from_model(User)
validator = ExcelValidator([mapping])

# Validate from file path
report = validator.validate("upload.xlsx")

# Validate from BytesIO (e.g., from HTTP upload)
from io import BytesIO
report = validator.validate(BytesIO(file_bytes))

if report.has_errors:
    print(report.summary())
    # "Validated 100 rows: 95 valid, 5 invalid. 8 errors found."
```

#### Constructor

```python
ExcelValidator(
    mappings: list[ExcelMapping],
    *,
    backend: str = "pydantic",
)
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `mappings` | `list[ExcelMapping]` | *(required)* | Mapping definitions, one per expected worksheet schema. |
| `backend` | `str` | `"pydantic"` | Validation backend. Currently only `"pydantic"` is supported. |

#### `validate()` Method

```python
def validate(
    self,
    source: str | Path | BinaryIO,
    *,
    sheet_name: str | None = None,
    max_errors: int | None = None,
    stop_on_first_error: bool = False,
) -> ValidationReport
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `source` | `str \| Path \| BinaryIO` | *(required)* | Excel file path or binary stream. |
| `sheet_name` | `str \| None` | `None` | Explicit worksheet name. Defaults to the first mapping's sheet. |
| `max_errors` | `int \| None` | `None` | Maximum number of errors to collect before stopping. |
| `stop_on_first_error` | `bool` | `False` | Stop validation after the first invalid row. |

**Returns:** `ValidationReport` — Structured report with row counts and cell-level errors.

#### Validation Behavior

1. **Header matching** — Excel headers are normalized (lowercased, whitespace-collapsed) and matched to ORM column names or configured Excel headers. Extra columns are ignored; missing required columns produce errors.
2. **Type coercion** — Pydantic v2 attempts type coercion (e.g., `"42"` → `42` for integer columns). Values that cannot be coerced produce type errors.
3. **Nullability** — Non-nullable columns without defaults that receive `None` produce required-field errors.
4. **Enum validation** — Values not in the enum's allowed set produce validation errors.

> **Important**: openpyxl's `DataValidation` (dropdowns) is a client-side hint only — it is **not enforced** when files are edited programmatically. Server-side validation via `ExcelValidator` is always required.

---

### ExcelImporter — Database Import

`ExcelImporter` reads validated Excel data and loads it into a database through SQLAlchemy ORM operations. It supports insert, upsert, and dry-run modes with batched execution and savepoint-based error recovery.

#### Basic Import

```python
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy_excel import ExcelMapping, ExcelImporter

mapping = ExcelMapping.from_model(User)
engine = create_engine("sqlite:///app.db")

with Session(engine) as session:
    importer = ExcelImporter([mapping], session=session)

    # Insert new rows
    result = importer.insert("users.xlsx")
    session.commit()

    print(result.summary())
    # "inserted=50, updated=0, skipped=0, failed=0, total=50, errors=0, duration_ms=123.45"
```

#### Constructor

```python
ExcelImporter(
    mappings: list[ExcelMapping],
    session: Session,
)
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `mappings` | `list[ExcelMapping]` | Model-to-sheet mappings for reading and loading rows. |
| `session` | `Session` | Active SQLAlchemy session. Transaction boundaries are managed by the caller. |

**Raises:** `ImportError_` — If `mappings` is empty.

#### Methods

| Method | Description |
|--------|-------------|
| `insert(source, *, batch_size=1000, validate=True)` | Insert all rows as new records. |
| `upsert(source, *, batch_size=1000, validate=True)` | Update existing rows (by key columns) or insert new ones. |
| `dry_run(source, *, validate=True)` | Simulate import without persisting. All changes are rolled back. |

All methods accept these parameters:

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `source` | `str \| Path \| BinaryIO` | *(required)* | Excel file path or binary stream. |
| `batch_size` | `int` | `1000` | Number of rows processed per batch/flush. |
| `validate` | `bool` | `True` | Run `ExcelValidator` before importing. If validation fails, returns early with error details. |

All methods return an `ImportResult`.

> **Note**: The caller is responsible for calling `session.commit()` after a successful import. The importer uses `session.flush()` and savepoints internally but does not commit the outer transaction.

---

### ExcelExporter — Query Export

`ExcelExporter` exports SQLAlchemy query results to formatted Excel files with styled headers, auto-sized columns, date formatting, auto-filters, and frozen header rows.

#### Basic Export

```python
from sqlalchemy import create_engine, select
from sqlalchemy.orm import Session
from sqlalchemy_excel import ExcelMapping, ExcelExporter

mapping = ExcelMapping.from_model(User)
engine = create_engine("sqlite:///app.db")

with Session(engine) as session:
    users = list(session.execute(select(User)).scalars().all())

exporter = ExcelExporter([mapping])

# Save to file
exporter.export(users, "users_export.xlsx")

# Get as bytes (for HTTP responses)
xlsx_bytes = exporter.export(users)
```

#### Constructor

```python
ExcelExporter(mappings: list[ExcelMapping])
```

| Parameter | Type | Description |
|-----------|------|-------------|
| `mappings` | `list[ExcelMapping]` | One or more mappings defining the export structure. |

**Raises:** `ExportError` — If `mappings` is empty.

#### `export()` Method

```python
def export(
    self,
    rows: Sequence[Any],
    path: str | Path | None = None,
    *,
    sheet_name: str | None = None,
) -> bytes | None
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `rows` | `Sequence[Any]` | *(required)* | ORM model instances or dictionaries to export. |
| `path` | `str \| Path \| None` | `None` | File path to save. If `None`, returns the Excel file as `bytes`. |
| `sheet_name` | `str \| None` | `None` | Override the sheet name from the mapping. |

**Returns:** `bytes` if `path` is `None`, otherwise `None`.

**Raises:** `ExportError` — If workbook creation or saving fails.

#### Export Features

- **Header styling** — Blue background, white bold font, centered alignment.
- **Date formatting** — `datetime` values formatted as `YYYY-MM-DD HH:MM:SS`, `date` values as `YYYY-MM-DD`.
- **Auto-width** — Column widths auto-adjust to content (capped at 50 characters).
- **Auto-filter** — Filter arrows on all columns.
- **Freeze panes** — Header row frozen.
- **Dict support** — Rows can be ORM instances or plain dictionaries.

---

## Validation and Error Reporting

### CellError

Each validation error is represented as a frozen `CellError` dataclass:

```python
@dataclass(frozen=True)
class CellError:
    row: int              # Excel row number (1-based, includes header)
    column: str           # Column name or header
    value: Any            # The raw value that failed validation
    expected_type: str    # Human-readable expected type
    message: str          # Descriptive error message
    error_code: str       # Machine-readable error code
```

**Example:**

```python
for error in report.errors:
    print(f"Row {error.row}, Column '{error.column}': "
          f"{error.message} (got: {error.value!r}, "
          f"expected: {error.expected_type}, code: {error.error_code})")
# Row 5, Column 'age': Input should be a valid integer (got: 'abc', expected: int, code: int_parsing)
```

### ValidationReport

`ValidationReport` aggregates validation results with convenience methods:

```python
@dataclass
class ValidationReport:
    errors: list[CellError]     # All collected cell-level errors
    total_rows: int             # Total data rows validated
    valid_rows: int             # Rows with zero errors
    invalid_rows: int           # Rows with at least one error
```

#### Properties and Methods

| Member | Type | Description |
|--------|------|-------------|
| `has_errors` | `bool` (property) | `True` if any validation errors exist. |
| `summary()` | `str` | Human-readable summary string. |
| `to_dict()` | `dict` | JSON-serializable dictionary (for API responses). |
| `errors_by_row()` | `dict[int, list[CellError]]` | Group errors by Excel row number. |
| `to_excel(path)` | `None` | Export error report to an Excel file. |

#### Example: Processing Validation Results

```python
report = validator.validate("upload.xlsx")

# Quick check
if not report.has_errors:
    print("All rows valid!")

# Summary
print(report.summary())
# "Validated 100 rows: 95 valid, 5 invalid. 8 errors found."

# Iterate errors by row
for row_num, errors in report.errors_by_row().items():
    print(f"Row {row_num}:")
    for e in errors:
        print(f"  {e.column}: {e.message}")

# JSON serialization (for API responses)
import json
print(json.dumps(report.to_dict(), indent=2, default=str))

# Export error report to Excel
report.to_excel("validation_errors.xlsx")
```

---

## Import Strategies and Results

### Insert Mode

Inserts all rows as new records using batched `session.add_all()` with savepoint-protected flushes. If a batch fails (e.g., `IntegrityError`), it is rolled back and the error is recorded; other batches continue.

```python
result = importer.insert("data.xlsx", batch_size=500)
session.commit()
```

### Upsert Mode

For each row, looks up an existing record by the mapping's `key_columns`. If found, updates all columns; if not found, inserts a new record. Uses per-batch savepoints with automatic single-row retry on `IntegrityError`.

```python
# Use email as the upsert key instead of primary key
mapping = ExcelMapping.from_model(User, key_columns=["email"])
importer = ExcelImporter([mapping], session=session)

result = importer.upsert("data.xlsx")
session.commit()
```

> **Note**: Upsert requires at least one key column. If `key_columns` is not specified in `ExcelMapping.from_model()`, primary key columns are used by default.

### Dry Run Mode

Executes the full insert pipeline (reading, validation, ORM object creation, flush) but rolls back every batch via savepoints. No data is persisted. Useful for previewing what would happen.

```python
result = importer.dry_run("data.xlsx")
print(result.summary())
# "inserted=50, updated=0, skipped=0, failed=0, total=50, errors=0, duration_ms=89.12"
# (nothing actually written to DB)
```

### ImportResult

All import methods return an `ImportResult` dataclass:

```python
@dataclass
class ImportResult:
    inserted: int = 0           # Rows successfully inserted
    updated: int = 0            # Rows successfully updated (upsert only)
    skipped: int = 0            # Rows skipped
    failed: int = 0             # Rows that failed to import
    errors: list[str] = []      # Human-readable error messages
    duration_ms: float = 0.0    # Total operation duration in milliseconds
```

| Property/Method | Type | Description |
|-----------------|------|-------------|
| `total` | `int` (property) | `inserted + updated + skipped + failed` |
| `summary()` | `str` | Concise summary string. |

---

## CLI Reference

sqlalchemy-excel provides a CLI for scripting and automation. All commands use a `--model` option in `module.path:ClassName` format.

```bash
sqlalchemy-excel --version
sqlalchemy-excel --help
```

### template

Generate an Excel template from a SQLAlchemy model.

```bash
sqlalchemy-excel template \
    --model myapp.models:User \
    --output users_template.xlsx \
    --sample-data \
    --sheet-name "User Import"
```

| Option | Required | Default | Description |
|--------|----------|---------|-------------|
| `--model` | Yes | — | ORM model path (`module.path:ClassName`) |
| `--output` | No | `template.xlsx` | Output file path |
| `--sample-data` | No | `False` | Include a sample data row |
| `--sheet-name` | No | Table name | Override worksheet name |

### validate

Validate an Excel file against a model schema.

```bash
# Text summary (default)
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx

# JSON output
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx --format json

# Export error report to Excel
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx \
    --format excel --output errors.xlsx
```

| Option | Required | Default | Description |
|--------|----------|---------|-------------|
| `--model` | Yes | — | ORM model path |
| `--input` | Yes | — | Excel file to validate |
| `--format` | No | `text` | Output format: `text`, `json`, or `excel` |
| `--output` | No | `validation_report.xlsx` | Output path (for `excel` format) |

Exit code is `1` if validation errors are found, `0` otherwise.

### import

Import an Excel file into a database.

```bash
# Basic insert
sqlalchemy-excel import \
    --model myapp.models:User \
    --input users.xlsx \
    --db sqlite:///app.db

# Upsert with custom batch size
sqlalchemy-excel import \
    --model myapp.models:User \
    --input users.xlsx \
    --db postgresql://user:pass@localhost/mydb \
    --mode upsert \
    --batch-size 500

# Preview without persisting
sqlalchemy-excel import \
    --model myapp.models:User \
    --input users.xlsx \
    --db sqlite:///app.db \
    --dry-run
```

| Option | Required | Default | Description |
|--------|----------|---------|-------------|
| `--model` | Yes | — | ORM model path |
| `--input` | Yes | — | Excel file to import |
| `--db` | Yes | — | Database URL (SQLAlchemy format) |
| `--mode` | No | `insert` | Import mode: `insert` or `upsert` |
| `--dry-run` | No | `False` | Preview import without persisting |
| `--batch-size` | No | `1000` | Batch size for DB operations |

### export

Export database records to an Excel file.

```bash
sqlalchemy-excel export \
    --model myapp.models:User \
    --db sqlite:///app.db \
    --output users_export.xlsx
```

| Option | Required | Default | Description |
|--------|----------|---------|-------------|
| `--model` | Yes | — | ORM model path |
| `--db` | Yes | — | Database URL |
| `--output` | No | `export.xlsx` | Output file path |

### inspect

Inspect an Excel file's structure without requiring a model.

```bash
sqlalchemy-excel inspect --input mystery.xlsx
```

```
File: mystery.xlsx
Sheets: 2

  Sheet: users
  Headers: Id, Name, Email, Age
  Data rows: 150
  Columns:
    Id: int
    Name: str
    Email: str
    Age: int, NoneType

  Sheet: orders
  Headers: Order Id, User Id, Amount, Date
  Data rows: 500
  Columns:
    Order Id: int
    User Id: int
    Amount: float
    Date: datetime
```

| Option | Required | Default | Description |
|--------|----------|---------|-------------|
| `--input` | Yes | — | Excel file to inspect |

---

## FastAPI Integration

sqlalchemy-excel provides a router factory that generates template download, upload validation, and database import endpoints from a single model class.

### Installation

```bash
pip install sqlalchemy-excel[fastapi]
```

### Basic Setup

```python
from fastapi import FastAPI
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from sqlalchemy_excel.integrations.fastapi import create_import_router

# Your model and database setup
from myapp.models import User

engine = create_engine("sqlite:///app.db")

def get_session():
    with Session(engine) as session:
        yield session

app = FastAPI()

# Creates 4 endpoints under /users/
router = create_import_router(
    User,
    prefix="/users",
    tags=["users"],
    session_dependency=get_session,
)
app.include_router(router)
```

### Generated Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `{prefix}/template` | Download an Excel template with sample data. Returns `.xlsx` file. |
| `POST` | `{prefix}/validate` | Upload and validate an Excel file. Returns `ValidationReport` as JSON. |
| `POST` | `{prefix}/import` | Validate and import an uploaded Excel file. Returns import summary. Raises `422` if validation fails. |
| `GET` | `{prefix}/health` | Health check. Returns `{"status": "ok", "model": "User"}`. |

> **Note**: The `/import` endpoint is only created when `session_dependency` is provided.

### `create_import_router()` Signature

```python
def create_import_router(
    model: type,
    *,
    prefix: str = "",
    tags: list[str] | None = None,
    session_dependency: Any = None,
) -> APIRouter
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `model` | `type` | *(required)* | SQLAlchemy ORM model class. |
| `prefix` | `str` | `""` | URL prefix for all endpoints. |
| `tags` | `list[str] \| None` | `None` | OpenAPI tags for endpoint grouping. |
| `session_dependency` | `Any` | `None` | FastAPI `Depends` callable that yields a `Session`. Required for the `/import` endpoint. |

---

## Exception Hierarchy

All custom exceptions inherit from `SqlalchemyExcelError`:

```
SqlalchemyExcelError                  # Base exception
├── MappingError                      # ORM introspection failures
├── TemplateError                     # Template generation failures
├── ReaderError                       # Excel file reading failures
│   ├── FileFormatError               # Invalid .xlsx file
│   ├── SheetNotFoundError            # Sheet name not found
│   └── HeaderMismatchError           # Header/column mismatch
├── ValidationError                   # Data validation failures (wraps report)
├── ImportError_                      # Database import failures (underscore avoids builtin clash)
│   ├── DuplicateKeyError             # Unique constraint violation
│   └── ConstraintViolationError      # Other DB constraint violations
└── ExportError                       # Export failures
```

### Catching Errors

```python
from sqlalchemy_excel.exceptions import (
    SqlalchemyExcelError,
    MappingError,
    TemplateError,
    ValidationError,
    ImportError_,
)

try:
    mapping = ExcelMapping.from_model(SomeModel)
except MappingError as e:
    print(f"Failed to introspect model: {e}")

# Or catch all library errors
try:
    template.save("output.xlsx")
except SqlalchemyExcelError as e:
    print(f"sqlalchemy-excel error: {e}")
```

> **Note**: `ImportError_` has a trailing underscore to avoid shadowing Python's builtin `ImportError`.

---

## Security

sqlalchemy-excel is designed for processing untrusted Excel file uploads. The following security measures are built in:

### XML Attack Protection

[defusedxml](https://pypi.org/project/defusedxml/) is a **required dependency** and is verified at import time. It protects against:

- **XML Entity Expansion (Billion Laughs)** — Exponential entity expansion that can cause denial of service.
- **External Entity Injection (XXE)** — Entities that reference external resources.
- **DTD Retrieval** — External DTD loading that can leak internal data.

If `defusedxml` is not installed, importing `sqlalchemy_excel` raises an error immediately.

### Formula Injection

When generating templates or exporting data, be aware that Excel can execute formulas starting with `=`, `+`, `-`, `@`, `\t`, or `\r`. If user-supplied data could contain these prefixes, consider sanitizing values before export.

### File Size

For production deployments, configure maximum upload file sizes at the web framework level (e.g., FastAPI's `UploadFile` limits, nginx `client_max_body_size`) to prevent memory exhaustion from oversized uploads.

### Read-Only Mode

When reading untrusted files for validation, openpyxl is used in `read_only=True` mode, which:

- Reduces memory usage (streaming row access)
- Avoids loading embedded macros
- Prevents execution of Excel VBA code

---

## Configuration Reference

### ExcelTemplate Configuration

| Parameter | Default | Description |
|-----------|---------|-------------|
| `include_sample_data` | `False` | Add a sample data row showing expected formats. |

### ExcelValidator Configuration

| Parameter | Default | Description |
|-----------|---------|-------------|
| `backend` | `"pydantic"` | Validation backend. Only `"pydantic"` currently supported. |
| `max_errors` | `None` | Maximum errors to collect. `None` = unlimited. |
| `stop_on_first_error` | `False` | Stop after the first invalid row. |

### ExcelImporter Configuration

| Parameter | Default | Description |
|-----------|---------|-------------|
| `batch_size` | `1000` | Rows per batch for insert/upsert flush cycles. |
| `validate` | `True` | Run validation before importing. |

### ExcelMapping Configuration

| Parameter | Default | Description |
|-----------|---------|-------------|
| `sheet_name` | `__tablename__` | Worksheet name. |
| `key_columns` | Primary keys | Columns used for upsert key matching. |
| `include` | All columns | Allowlist of column names. |
| `exclude` | None | Denylist of column names. |
| `header_map` | Auto-generated | Custom Excel header names. |

---

## Development

### Setup

```bash
git clone https://github.com/yeongseon/sqlalchemy-excel.git
cd sqlalchemy-excel
pip install -e ".[all]"
```

### Running Tests

```bash
# Run all tests
pytest

# With coverage
pytest --cov=sqlalchemy_excel --cov-report=term-missing

# Run specific test file
pytest tests/unit/test_mapping.py -v
```

### Linting and Formatting

```bash
# Lint
ruff check .

# Auto-fix
ruff check . --fix

# Format
ruff format .
```

### Type Checking

```bash
mypy --strict src/
```

### Project Structure

```
src/sqlalchemy_excel/        # Library source code
├── __init__.py              # Public API re-exports
├── mapping.py               # ORM → ExcelMapping extraction
├── template.py              # Template generation
├── export.py                # Query result → Excel export
├── cli.py                   # Click CLI
├── exceptions.py            # Exception hierarchy
├── _types.py                # Internal type aliases
├── _compat.py               # Compatibility helpers
├── reader/                  # Excel file readers
│   ├── base.py              # Reader interface
│   └── openpyxl_reader.py   # Default openpyxl reader
├── validation/              # Validation engine
│   ├── report.py            # CellError, ValidationReport
│   ├── pydantic_backend.py  # Pydantic v2 backend
│   └── engine.py            # ExcelValidator orchestrator
├── load/                    # Database import
│   ├── strategies.py        # Insert/Upsert/DryRun strategies
│   └── importer.py          # ExcelImporter orchestrator
└── integrations/
    └── fastapi.py           # FastAPI router factory
```

---

## License

MIT License. Copyright (c) 2025 Yeongseon Choe.

See [LICENSE](LICENSE) for full text.


## Operational Guides

- [Release checklist](docs/release-checklist.md)
- [Security defaults checklist](docs/security-defaults-checklist.md)
- [10-minute backend upload pipeline tutorial](docs/tutorials/backend-upload-pipeline-10min.md)
- [Growth roadmap](docs/GROWTH_ROADMAP_2026.md)

