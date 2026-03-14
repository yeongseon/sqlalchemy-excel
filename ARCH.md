# ARCH.md — sqlalchemy-excel Architecture Document

> Version: 0.1.0-draft | Date: 2026-03-14

## 1. System Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    sqlalchemy-excel                         │
│                                                             │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌───────────┐  │
│  │ mapping  │→ │ template │  │  reader  │→ │validation │  │
│  │          │  │          │  │          │  │           │  │
│  │ ORM→     │  │ Mapping→ │  │ .xlsx→   │  │ Rows→     │  │
│  │ Schema   │  │ .xlsx    │  │ Rows     │  │ Report    │  │
│  └──────────┘  └──────────┘  └──────────┘  └─────┬─────┘  │
│       │                                          │         │
│       │              ┌──────────┐                │         │
│       └─────────────→│   load   │←───────────────┘         │
│                      │          │                           │
│                      │ Rows→DB  │                           │
│                      └──────────┘                           │
│                                                             │
│  ┌──────────┐  ┌──────────────┐  ┌────────┐               │
│  │  export  │  │ integrations │  │  cli   │               │
│  │          │  │              │  │        │               │
│  │ Query→   │  │ FastAPI      │  │ Click  │               │
│  │ .xlsx    │  │ router       │  │ cmds   │               │
│  └──────────┘  └──────────────┘  └────────┘               │
└─────────────────────────────────────────────────────────────┘
```

## 2. Module Architecture

### 2.1 Layer Diagram

```
┌──────────────────────────────────────────────┐
│              Public API Layer                │
│  __init__.py: ExcelMapping, ExcelTemplate,   │
│  ExcelValidator, ExcelImporter, ExcelExporter│
├──────────────────────────────────────────────┤
│           Interface Layer (CLI/Web)          │
│  cli.py (Click) | integrations/fastapi.py   │
├──────────────────────────────────────────────┤
│             Service Layer                    │
│  mapping.py | template.py | export.py       │
│  validation/engine.py | load/importer.py    │
├──────────────────────────────────────────────┤
│             Backend Layer                    │
│  reader/openpyxl_reader.py                  │
│  reader/pandas_reader.py (optional)         │
│  validation/pydantic_backend.py             │
│  validation/pandera_backend.py (optional)   │
│  load/strategies.py                         │
├──────────────────────────────────────────────┤
│           Infrastructure Layer               │
│  _types.py | _compat.py | exceptions.py     │
└──────────────────────────────────────────────┘
```

### 2.2 Dependency Rules

1. **Public API** → Service Layer only (never directly to Backend)
2. **Service Layer** → Backend Layer via abstractions (Protocol/ABC)
3. **Backend Layer** → External libraries (openpyxl, pydantic, pandas)
4. **Interface Layer** → Service Layer only
5. **Infrastructure** ← Used by all layers

### 2.3 Optional Dependencies

```
sqlalchemy-excel            (core: sqlalchemy, openpyxl, pydantic, defusedxml, click)
sqlalchemy-excel[pandas]    (+ pandas)
sqlalchemy-excel[pandera]   (+ pandera, pandas)
sqlalchemy-excel[fastapi]   (+ fastapi, python-multipart)
sqlalchemy-excel[dev]       (+ pytest, hypothesis, ruff, mypy, coverage)
sqlalchemy-excel[all]       (all of the above)
```

## 3. Module Specifications

### 3.1 `mapping.py` — ORM Schema Extraction

**Responsibility**: Introspect SQLAlchemy ORM model and produce an `ExcelMapping` dataclass.

**Key Types**:
```python
@dataclass(frozen=True)
class ColumnMapping:
    """Mapping for a single ORM column to Excel column."""
    name: str                    # ORM column name
    excel_header: str            # Display name in Excel header
    python_type: type            # Python type (str, int, float, etc.)
    sqla_type: TypeEngine        # SQLAlchemy type object
    nullable: bool
    primary_key: bool
    has_default: bool            # Server default or Python default
    default_value: Any | None
    enum_values: list[str] | None  # For Enum columns
    max_length: int | None       # For String(N)
    description: str | None      # From column doc/comment
    foreign_key: str | None      # Referenced table.column

@dataclass(frozen=True)
class ExcelMapping:
    """Complete mapping from ORM model to Excel structure."""
    model_class: type             # The ORM model class
    sheet_name: str
    columns: list[ColumnMapping]
    key_columns: list[str]        # For upsert operations
```

**Design Decisions**:
- Uses `sqlalchemy.inspect()` for model introspection
- Frozen dataclass for immutability (mapping shouldn't change after creation)
- `ExcelMapping.from_model(Model, ...)` class method as primary factory

### 3.2 `template.py` — Excel Template Generation

**Responsibility**: Convert `ExcelMapping` to a formatted .xlsx file.

**Key API**:
```python
class ExcelTemplate:
    def __init__(self, mappings: list[ExcelMapping]) -> None: ...
    def save(self, path: str | Path) -> None: ...
    def to_bytes(self) -> bytes: ...
    def to_bytesio(self) -> BytesIO: ...
```

**Design Decisions**:
- One sheet per `ExcelMapping`
- Header row: bold, colored background, auto-filter
- Column comments: type hint, nullable, constraints
- DataValidation for enum columns (dropdown lists)
- Optional sample data row
- Column width auto-fit based on header + type

### 3.3 `reader/` — Excel File Parsing

**Responsibility**: Read .xlsx files and produce structured row data.

**Architecture**: Strategy pattern with pluggable backends.

```python
# reader/base.py
class ReaderResult(NamedTuple):
    headers: list[str]
    rows: Iterable[dict[str, Any]]
    total_rows: int | None  # None if streaming

class BaseReader(Protocol):
    def read(
        self,
        source: str | Path | BinaryIO,
        sheet_name: str | None = None,
        header_row: int = 1,
    ) -> ReaderResult: ...

# reader/openpyxl_reader.py — default, always available
class OpenpyxlReader(BaseReader):
    def __init__(self, read_only: bool = False) -> None: ...

# reader/pandas_reader.py — optional
class PandasReader(BaseReader): ...
```

**Design Decisions**:
- Default: openpyxl (always available)
- `read_only=True` for files > threshold (configurable, default 10MB)
- Returns `Iterable[dict]` to support both eager and streaming modes
- Header normalization: strip whitespace, lowercase, spaces→underscores

### 3.4 `validation/` — Validation Engine

**Responsibility**: Validate parsed rows against schema, produce structured error reports.

**Architecture**: Pluggable validation backends.

```python
# validation/engine.py
class ExcelValidator:
    def __init__(
        self,
        mappings: list[ExcelMapping],
        backend: str = "pydantic",  # "pydantic" or "pandera"
    ) -> None: ...

    def validate(
        self,
        source: str | Path | BinaryIO,
        *,
        max_errors: int | None = None,
        stop_on_first_error: bool = False,
    ) -> ValidationReport: ...

# validation/report.py
@dataclass
class CellError:
    row: int              # 1-based row number in Excel
    column: str           # Column name
    value: Any            # Original cell value
    expected_type: str    # Expected type description
    message: str          # Human-readable error message
    error_code: str       # Machine-readable error code

@dataclass
class ValidationReport:
    errors: list[CellError]
    total_rows: int
    valid_rows: int
    invalid_rows: int

    @property
    def has_errors(self) -> bool: ...
    def summary(self) -> str: ...
    def to_dict(self) -> dict: ...
    def to_excel(self, path: str | Path) -> None: ...
```

**Validation Pipeline**:
```
Raw Cell Value
    → Type coercion (str→int, str→date, etc.)
    → Nullable check
    → Length/range check
    → Enum membership check
    → Custom validators (if registered)
    → CellError (if any step fails)
```

**Design Decisions**:
- Pydantic v2 as default backend (auto-generates model from ExcelMapping)
- Pandera as optional backend (requires pandas)
- All errors collected (not fail-fast by default)
- ValidationReport exportable to Excel (error rows highlighted)

### 3.5 `load/` — Database Import

**Responsibility**: Load validated data into database via SQLAlchemy Session.

**Architecture**: Strategy pattern for different load modes.

```python
# load/importer.py
class ExcelImporter:
    def __init__(
        self,
        mappings: list[ExcelMapping],
        session: Session,
    ) -> None: ...

    def insert(self, source, *, batch_size=1000) -> ImportResult: ...
    def upsert(self, source, *, batch_size=1000) -> ImportResult: ...
    def dry_run(self, source) -> ImportResult: ...

@dataclass
class ImportResult:
    inserted: int
    updated: int
    skipped: int
    failed: int
    errors: list[CellError]
    duration_ms: float
```

**Load Strategies** (`load/strategies.py`):
```python
class LoadStrategy(Protocol):
    def execute(
        self,
        session: Session,
        model_class: type,
        rows: Iterable[dict],
        key_columns: list[str],
        batch_size: int,
    ) -> ImportResult: ...

class InsertStrategy(LoadStrategy): ...   # session.add_all()
class UpsertStrategy(LoadStrategy): ...   # merge() or ON CONFLICT
```

**Design Decisions**:
- All operations through SQLAlchemy Session (never raw SQL for portability)
- Batch processing with configurable batch_size
- Transaction boundary managed by caller (commit/rollback outside importer)
- Dry-run validates and reports without touching DB

### 3.6 `export.py` — Query Result Export

**Responsibility**: Export SQLAlchemy query results to formatted Excel.

```python
class ExcelExporter:
    def __init__(self, mappings: list[ExcelMapping]) -> None: ...

    def export(
        self,
        rows: Sequence[Any],  # ORM instances or Row objects
        path: str | Path | None = None,
    ) -> bytes | None: ...  # Returns bytes if path is None
```

### 3.7 `integrations/fastapi.py` — FastAPI Router Factory

**Responsibility**: Provide ready-to-use FastAPI endpoints.

```python
def create_import_router(
    model: type,
    *,
    prefix: str = "",
    tags: list[str] | None = None,
    session_dependency: Any = None,  # FastAPI Depends
) -> APIRouter:
    """Create a FastAPI router with template download and file upload endpoints."""
```

**Endpoints Generated**:
- `GET /template` — Download Excel template
- `POST /upload` — Upload and validate Excel file
- `POST /import` — Upload, validate, and import to DB
- `GET /import/{job_id}` — Check import job status (async)

### 3.8 `cli.py` — CLI Interface

**Responsibility**: Click-based command-line interface.

```python
@click.group()
def cli(): ...

@cli.command()
@click.option("--model", required=True)
@click.option("--output", default="template.xlsx")
def template(model: str, output: str) -> None: ...

@cli.command()
@click.option("--model", required=True)
@click.option("--input", required=True)
@click.option("--format", type=click.Choice(["text", "json", "excel"]))
def validate(model: str, input: str, format: str) -> None: ...
```

## 4. Data Flow

### 4.1 Template Generation Flow

```
ORM Model
    → mapping.ExcelMapping.from_model()
    → ExcelMapping (dataclass)
    → template.ExcelTemplate([mapping])
    → openpyxl.Workbook
    → .xlsx file / BytesIO
```

### 4.2 Upload → Validate → Import Flow

```
.xlsx file (upload)
    → reader.OpenpyxlReader.read()
    → ReaderResult (headers + row iterator)
    → validation.ExcelValidator.validate()
    → ValidationReport
    → [if has_errors] → report.to_excel() → error report
    → [if clean] → load.ExcelImporter.upsert()
    → ImportResult (counts + stats)
```

### 4.3 Export Flow

```
SQLAlchemy Query
    → session.execute(select(Model))
    → Row objects
    → export.ExcelExporter.export(rows)
    → openpyxl.Workbook
    → .xlsx file / bytes
```

## 5. Error Handling Architecture

### Exception Hierarchy

```
SqlalchemyExcelError (base)
├── MappingError          # ORM introspection failures
├── TemplateError         # Template generation failures
├── ReaderError           # Excel parsing failures
│   ├── FileFormatError   # Not a valid xlsx
│   ├── SheetNotFoundError
│   └── HeaderMismatchError
├── ValidationError       # Data validation failures (wraps report)
├── ImportError_          # DB import failures (underscore to avoid builtin)
│   ├── DuplicateKeyError
│   └── ConstraintViolationError
└── ExportError           # Export failures
```

### Error Recovery Strategy

1. **Template generation errors**: Fail fast with clear message about unsupported type
2. **Reader errors**: Wrap openpyxl exceptions with file context (path, sheet, row)
3. **Validation errors**: Never raise — collect in ValidationReport
4. **Import errors**: Transaction rollback, partial results in ImportResult
5. **All errors**: Include actionable message ("expected X, got Y at row N, column C")

## 6. Security Architecture

### Threats & Mitigations

| Threat | Mitigation |
|--------|-----------|
| XML Bomb (billion laughs) | defusedxml required at import time |
| Formula injection | Escape cells starting with `=`, `+`, `-`, `@`, `\t`, `\r` |
| File size DoS | Configurable max_file_size (default 50MB) |
| Zip bomb | openpyxl read_only mode + file size check before extraction |
| Path traversal | Validate file paths, use tempfile for uploads |
| SQL injection | All DB operations via SQLAlchemy ORM (parameterized) |

### Security Initialization

```python
# _compat.py
def ensure_defusedxml() -> None:
    """Ensure defusedxml is installed. Called at module import time."""
    try:
        import defusedxml  # noqa: F401
    except ImportError:
        raise ImportError(
            "defusedxml is required for safe XML processing. "
            "Install it with: pip install defusedxml"
        )
```

## 7. Performance Considerations

### Memory Management

- **Small files (< 10MB)**: Load fully into memory (openpyxl normal mode)
- **Large files (≥ 10MB)**: Stream via openpyxl read_only mode
- **Validation**: Process rows in chunks (default chunk_size=1000)
- **DB import**: Batch insert/upsert (default batch_size=1000)

### Optimization Targets

| Operation | Target | Approach |
|-----------|--------|----------|
| Template gen | < 100ms | Pre-compute styles, batch cell writes |
| Parse 10K rows | < 3s | openpyxl optimized iteration |
| Validate 10K rows | < 1s | Pydantic v2 compiled validators |
| Import 10K rows | < 5s | Batch add_all(), bulk operations |
| Export 10K rows | < 3s | write_only mode for large exports |

## 8. Testing Architecture

```
tests/
├── conftest.py          # Shared fixtures
│   ├── in_memory_engine  # SQLite :memory:
│   ├── sample_models     # User, Product, Order models
│   └── temp_workbook     # Temporary .xlsx creation
├── unit/
│   ├── test_mapping.py   # ExcelMapping extraction
│   ├── test_template.py  # Template generation
│   ├── test_reader.py    # File parsing
│   ├── test_validation.py # Validation engine
│   ├── test_importer.py  # DB import
│   └── test_report.py    # ValidationReport
├── integration/
│   ├── test_end_to_end.py  # Full pipeline
│   └── test_fastapi.py    # FastAPI integration
└── fixtures/
    ├── sample_valid.xlsx
    ├── sample_invalid.xlsx
    └── sample_large.xlsx
```

### Test Invariants (Hypothesis)

1. **Round-trip**: `template → fill with valid data → validate → no errors`
2. **Idempotent validation**: `validate(data) == validate(data)` (same report)
3. **Export/Import**: `export(query) → import(file) → same data in DB`
