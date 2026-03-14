# TDD.md — sqlalchemy-excel Technical Design Document

> **Version**: 0.1.0  
> **Last Updated**: 2026-03-15  
> **Status**: MVP Complete (feature-complete)

---

## 1. Implementation Overview

sqlalchemy-excel v0.1.0 is fully implemented, tested, and ready for PyPI publication. This document describes the detailed technical design of each module, the excel-dbapi integration architecture, import/validation strategies, and comprehensive usage examples.

### 1.1 Development Phases (All Complete)

| Phase | Scope | Modules | Status |
|-------|-------|---------|--------|
| Phase 1: Foundation | Project skeleton, exceptions, types, compatibility | `pyproject.toml`, `exceptions.py`, `_types.py`, `_compat.py` | ✅ |
| Phase 2: Core Schema | ORM introspection → mapping extraction | `mapping.py` | ✅ |
| Phase 3: Template & Reader | Template generation, reader protocol + implementations | `template.py`, `reader/base.py`, `reader/openpyxl_reader.py` | ✅ |
| Phase 4: Validation | Pydantic-based validation engine + error reporting | `validation/engine.py`, `validation/pydantic_backend.py`, `validation/report.py` | ✅ |
| Phase 5: Import & Export | Database import strategies + query export | `load/strategies.py`, `load/importer.py`, `export.py` | ✅ |
| Phase 6: excel-dbapi Integration | Dual-channel session, SQL-based reader, full dependency wiring | `excelio/session.py`, `reader/excel_dbapi_reader.py` | ✅ |
| Phase 7: CLI & Integration | Click CLI, FastAPI router, public API | `cli.py`, `integrations/fastapi.py`, `__init__.py` | ✅ |

### 1.2 File Map

```
src/sqlalchemy_excel/
├── __init__.py              106 lines   Public API (lazy imports, defusedxml check)
├── mapping.py               269 lines   ORM → ExcelMapping extraction
├── template.py              272 lines   ExcelMapping → .xlsx template
├── export.py                179 lines   Query result → .xlsx export
├── cli.py                   354 lines   Click CLI (5 commands)
├── exceptions.py             88 lines   Full exception hierarchy
├── _types.py                 20 lines   FilePath, FileSource, RowDict, ColumnName
├── _compat.py                63 lines   defusedxml, sanitize_cell_value()
├── excelio/
│   ├── __init__.py            7 lines   Re-exports ExcelWorkbookSession
│   └── session.py           104 lines   Dual-channel wrapper
├── reader/
│   ├── __init__.py           13 lines   Re-exports
│   ├── base.py               61 lines   ReaderResult, BaseReader Protocol
│   ├── openpyxl_reader.py   238 lines   Legacy reader (compatibility)
│   └── excel_dbapi_reader.py 189 lines  Primary SQL-based reader
├── validation/
│   ├── __init__.py            8 lines   Re-exports
│   ├── engine.py            186 lines   ExcelValidator orchestrator
│   ├── pydantic_backend.py  222 lines   Dynamic Pydantic model generation
│   └── report.py            132 lines   CellError, ValidationReport
├── load/
│   ├── __init__.py            8 lines   Re-exports
│   ├── importer.py          347 lines   ExcelImporter orchestrator
│   └── strategies.py        373 lines   Insert/Upsert/DryRun strategies
└── integrations/
    ├── __init__.py            7 lines   Re-exports
    └── fastapi.py           143 lines   create_import_router() factory
                            ──────
                           2,655 lines total
```

---

## 2. Dependency Graph

### 2.1 Core Dependencies

```
sqlalchemy-excel
├── sqlalchemy >= 2.0         # ORM schema source of truth, DB operations
├── excel-dbapi >= 1.0        # Full dependency: all Excel I/O
├── openpyxl >= 3.1           # Cell-level formatting via workbook channel
├── pydantic >= 2.0           # Row-level type coercion and validation
├── defusedxml >= 0.7         # Security: XML attack prevention
└── click >= 8.0              # CLI interface
```

### 2.2 excel-dbapi as Full Dependency

`excel-dbapi >= 1.0` is listed in `pyproject.toml` under `dependencies` — it is **not optional**. All Excel read/write operations go through excel-dbapi:

```toml
dependencies = [
    "sqlalchemy>=2.0",
    "openpyxl>=3.1",
    "pydantic>=2.0",
    "defusedxml>=0.7",
    "click>=8.0",
    "excel-dbapi>=1.0",
]
```

### 2.3 Internal Module Dependency Flow

```
                    ┌──────────────────────────────────┐
                    │        __init__.py                │
                    │   (lazy imports, defusedxml)      │
                    └──────┬──────┬──────┬──────┬──────┘
                           │      │      │      │
                    ┌──────▼──┐ ┌─▼──────▼─┐ ┌──▼──────┐
                    │cli.py   │ │template.py│ │export.py│
                    └──┬──┬──┘ └─────┬─────┘ └────┬────┘
                       │  │          │             │
              ┌────────▼──▼──────┐   │       ┌─────▼──────┐
              │validation/engine │   │       │   excelio/  │
              └────────┬─────────┘   │       │  session.py │
                       │             │       └─────┬──────┘
              ┌────────▼─────────┐   │             │
              │ load/importer.py │   │             │
              └────────┬─────────┘   │             │
                       │             │             │
              ┌────────▼─────────┐ ┌─▼─────────┐  │
              │load/strategies.py│ │mapping.py  │  │
              └──────────────────┘ └────────────┘  │
                       │                           │
              ┌────────▼──────────────────────────▼┐
              │  reader/excel_dbapi_reader.py       │
              │  (SQL-based Excel reading)          │
              └────────────────┬───────────────────┘
                               │ excel_dbapi.connect()
                               ▼
              ┌──────────────────────────────────────┐
              │          excel-dbapi >= 1.0           │
              │  PEP 249 DB-API 2.0 Driver           │
              └──────────────────────────────────────┘
```

---

## 3. Detailed Module Design

### 3.1 Module: `__init__.py`

#### Purpose
Public API surface with lazy imports and security initialization.

#### Design Pattern
**Lazy Import** — `__getattr__`-based deferred loading keeps startup fast and avoids circular imports.

#### Implementation

```python
from sqlalchemy_excel._compat import ensure_defusedxml

# Security check at import time — raises ImportError if defusedxml missing
ensure_defusedxml()

# Lazy import map: symbol → (module_path, class_name)
_LAZY_IMPORTS: dict[str, tuple[str, str]] = {
    "ExcelMapping": ("sqlalchemy_excel.mapping", "ExcelMapping"),
    "ColumnMapping": ("sqlalchemy_excel.mapping", "ColumnMapping"),
    "ExcelTemplate": ("sqlalchemy_excel.template", "ExcelTemplate"),
    "ExcelWorkbookSession": ("sqlalchemy_excel.excelio.session", "ExcelWorkbookSession"),
    "ExcelValidator": ("sqlalchemy_excel.validation.engine", "ExcelValidator"),
    "ValidationReport": ("sqlalchemy_excel.validation.report", "ValidationReport"),
    "CellError": ("sqlalchemy_excel.validation.report", "CellError"),
    "ExcelImporter": ("sqlalchemy_excel.load.importer", "ExcelImporter"),
    "ImportResult": ("sqlalchemy_excel.load.strategies", "ImportResult"),
    "ExcelExporter": ("sqlalchemy_excel.export", "ExcelExporter"),
}

def __getattr__(name: str) -> object:
    if name in _LAZY_IMPORTS:
        module_path, class_name = _LAZY_IMPORTS[name]
        module = importlib.import_module(module_path)
        return getattr(module, class_name)
    raise AttributeError(f"module 'sqlalchemy_excel' has no attribute {name!r}")
```

#### Exports

`__all__` lists 21 symbols: 10 lazy-loaded classes + 11 exception classes (re-exported directly).

### 3.2 Module: `_compat.py`

#### Purpose
Version compatibility helpers, optional dependency imports, and cell value sanitization.

#### Implementation

```python
_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

def ensure_defusedxml() -> None:
    """Verify defusedxml is installed. Raises ImportError if missing."""
    try:
        import defusedxml  # noqa: F401
    except ImportError as e:
        raise ImportError(
            "defusedxml is required for processing Excel files safely. "
            "Install it with: pip install defusedxml"
        ) from e

def import_optional(module_name: str, extra_name: str) -> Any:
    """Import an optional dependency with a helpful installation message."""
    try:
        return importlib.import_module(module_name)
    except ImportError as e:
        raise ImportError(
            f"{module_name} is required for this feature. "
            f"Install it with: pip install sqlalchemy-excel[{extra_name}]"
        ) from e

def sanitize_cell_value(value: object) -> object:
    """Prevent formula injection by prefixing dangerous cell values with '."""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return f"'{value}"
    return value
```

Usage locations: `template.py` (sample data), `export.py` (exported values), `report.py` (error report cells).

### 3.3 Module: `_types.py`

#### Purpose
Internal type aliases shared across the codebase.

```python
from os import PathLike
from typing import Any, BinaryIO, Union

FilePath = Union[str, PathLike[str]]
FileSource = Union[str, PathLike[str], BinaryIO]
RowDict = dict[str, Any]
ColumnName = str
```

### 3.4 Module: `exceptions.py`

#### Purpose
Full exception hierarchy. All custom exceptions inherit from `SqlalchemyExcelError`.

#### Hierarchy

```
SqlalchemyExcelError (base)
├── MappingError                  # ORM introspection failures
├── TemplateError                 # Template generation failures
├── ReaderError                   # Excel file reading failures
│   ├── FileFormatError           # Invalid .xlsx file
│   ├── SheetNotFoundError        # Sheet name not found (stores sheet_name, available)
│   └── HeaderMismatchError       # Header/column mismatch (stores missing, extra)
├── ValidationError               # Data validation failures (wraps ValidationReport)
├── ImportError_                  # Database import failures (underscore avoids builtin)
│   ├── DuplicateKeyError         # Unique constraint violation
│   └── ConstraintViolationError  # Other DB constraint violations
└── ExportError                   # Export failures
```

#### Special Exception Attributes

```python
class SheetNotFoundError(ReaderError):
    def __init__(self, sheet_name: str, available: list[str]) -> None:
        self.sheet_name = sheet_name
        self.available = available
        super().__init__(
            f"Sheet '{sheet_name}' not found. "
            f"Available sheets: {', '.join(available)}"
        )

class HeaderMismatchError(ReaderError):
    def __init__(self, missing: list[str], extra: list[str]) -> None:
        self.missing = missing
        self.extra = extra
        # Builds descriptive message from missing and extra lists

class ValidationError(SqlalchemyExcelError):
    def __init__(self, report: object) -> None:
        self.report = report  # ValidationReport instance
        super().__init__(str(report))
```

### 3.5 Module: `excelio/session.py` — ExcelWorkbookSession

#### Purpose
Dual-channel wrapper around an excel-dbapi connection providing both SQL cursor access (data channel) and direct openpyxl workbook access (format channel).

#### Design Patterns
- **Factory Method** — `open()` classmethod abstracts away connection construction
- **Adapter** — Wraps excel-dbapi connection with a simplified interface
- **Context Manager** — `with` statement support for automatic cleanup

#### Implementation

```python
excel_dbapi: Any = importlib.import_module("excel_dbapi")

class ExcelWorkbookSession:
    def __init__(self, conn: Any) -> None:
        self._conn: Any = conn
        self._cursor: Any = conn.cursor()

    @classmethod
    def open(
        cls,
        path: str | Path,
        *,
        create: bool = False,
        data_only: bool = False,
    ) -> ExcelWorkbookSession:
        conn = excel_dbapi.connect(
            str(path),
            engine="openpyxl",
            autocommit=False,
            create=create,
            data_only=data_only,
        )
        return cls(conn)

    @property
    def conn(self) -> Any:      return self._conn
    @property
    def cursor(self) -> Any:    return self._cursor
    @property
    def workbook(self) -> Any:  return self._conn.workbook  # openpyxl Workbook

    def commit(self) -> None:   self._conn.commit()
    def rollback(self) -> None: self._conn.rollback()
    def close(self) -> None:    self._conn.close()

    def __enter__(self) -> ExcelWorkbookSession: return self
    def __exit__(self, exc_type, exc_val, exc_tb) -> None: self.close()
```

#### Dual-Channel Architecture

```
ExcelWorkbookSession
│
├── Data Channel (SQL)
│   └── self._cursor = self._conn.cursor()
│       ├── cursor.execute("SELECT * FROM Sheet1")
│       └── cursor.fetchall() → List[Tuple]
│
└── Format Channel (openpyxl)
    └── self._conn.workbook → openpyxl.Workbook
        ├── workbook[sheet_name] → Worksheet
        ├── cell.font, cell.fill, cell.alignment
        └── DataValidation, comments, freeze_panes
```

#### Connection Configuration by Use Case

| Use Case | engine | autocommit | create | data_only |
|----------|--------|-----------|--------|-----------|
| Template generation | `"openpyxl"` | `False` | `True` | `False` |
| File export | `"openpyxl"` | `False` | `True` | `False` |
| File reading/validation | `"openpyxl"` | `True` | `False` | `True` |
| File reading/import | `"openpyxl"` | `True` | `False` | `True` |

#### Design Decisions
- `engine="openpyxl"` fixed — format channel (workbook property) only supported by openpyxl engine
- `autocommit=False` — requires explicit `commit()` for transaction safety
- `create=True` — creates new workbook file (for template/export)
- `data_only=True` — reads cached formula values (for read-only operations)

### 3.6 Module: `reader/excel_dbapi_reader.py` — ExcelDbapiReader

#### Purpose
SQL-based Excel reader replacing direct openpyxl usage. Uses `SELECT * FROM SheetName` via excel-dbapi cursors. Primary reader for `ExcelValidator` and `ExcelImporter`.

#### Design Pattern
- **Template Method** — `read()` orchestrates resolve → validate → connect → query → normalize → return
- **Adapter** — Adapts excel-dbapi cursor results to `ReaderResult`

#### Implementation

```python
class ExcelDbapiReader:
    def __init__(
        self,
        *,
        read_only: bool = False,
        max_file_size: int = 50 * 1024 * 1024,  # 50MB
    ) -> None:
        self.read_only = read_only
        self.max_file_size = max_file_size

    def read(
        self,
        source: FileSource,
        sheet_name: str | None = None,
        header_row: int | None = None,
    ) -> ReaderResult:
        file_path, remove_after_read = self._resolve_source(source)
        self._validate_file_size_path(file_path)

        conn = excel_dbapi.connect(
            str(file_path),
            engine="openpyxl",
            autocommit=True,
            data_only=True,
        )
        try:
            # Sheet resolution
            available_sheets = list(conn.workbook.sheetnames)
            resolved_sheet = sheet_name or available_sheets[0]
            if resolved_sheet not in available_sheets:
                raise SheetNotFoundError(resolved_sheet, available_sheets)

            # SQL-based data reading (UNQUOTED table name!)
            cursor = conn.cursor()
            cursor.execute(f"SELECT * FROM {resolved_sheet}")

            # Header extraction + normalization
            headers = self._normalize_headers(
                [str(item[0]) for item in cursor.description]
            )

            # Row iteration with empty-row filtering
            rows = []
            for raw_row in cursor.fetchall():
                values = list(raw_row)
                if all(self._is_empty_cell(v) for v in values):
                    continue
                rows.append(dict(zip(headers, values, strict=True)))

            return ReaderResult(headers=headers, rows=rows, total_rows=len(rows))
        finally:
            conn.close()
            if remove_after_read:
                self._remove_temp_file(file_path)
```

#### `_resolve_source()` — BinaryIO → Temp File Conversion

excel-dbapi requires file paths. Binary streams are persisted to temp files:

```python
def _resolve_source(self, source: FileSource) -> tuple[str, bool]:
    if isinstance(source, (str, Path, os.PathLike)):
        return str(source), False  # Path as-is, no cleanup needed

    # BinaryIO: read content → write to temp file
    binary_source = source
    original_position = binary_source.tell()
    content = binary_source.read()
    binary_source.seek(original_position)  # Restore cursor

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp.write(content)
        tmp.flush()
        temp_path = tmp.name

    return temp_path, True  # Temp path, cleanup needed after read
```

#### Helper Methods

| Method | Purpose | Behavior |
|--------|---------|----------|
| `_validate_file_size_path(path)` | File size check | `os.path.getsize(path)`, raises `ReaderError` if > `max_file_size` |
| `_normalize_headers(raw)` | Header normalization | Applies `normalize_header()`, checks empty/duplicate |
| `_is_empty_cell(value)` | Empty cell detection | `None` or whitespace-only string |
| `_remove_temp_file(path)` | Temp file cleanup | `with suppress(OSError): os.unlink(path)` |

#### Error Handling

```python
except (InvalidFileException, BadZipFile):
    raise FileFormatError("Input is not a valid .xlsx file")
except (SheetNotFoundError, ReaderError, FileFormatError):
    raise  # Re-raise without wrapping
except Exception as exc:
    raise ReaderError(f"Failed to read Excel data: {exc}") from exc
finally:
    if conn is not None:
        conn.close()         # Always close connection
    if remove_after_read:
        _remove_temp_file(file_path)  # Clean up temp file
```

#### Design Decisions
- `autocommit=True` — read-only, no transaction management needed
- `data_only=True` — reads cached formula values, not formulas
- Table names **unquoted** — `SELECT * FROM Sheet1` (quotes cause lookup failure in excel-dbapi)
- `suppress(OSError)` pattern for safe temp file cleanup

### 3.7 Module: `reader/base.py` — Reader Protocol

#### Purpose
Shared types and protocol definition for reader implementations.

```python
@dataclass(slots=True)
class ReaderResult:
    headers: list[str]           # Normalized header names
    rows: Iterable[RowDict]      # Row dictionaries
    total_rows: int | None       # Row count (None for streaming)

class BaseReader(Protocol):
    def read(
        self,
        source: FileSource,
        sheet_name: str | None = None,
        header_row: int | None = None,
    ) -> ReaderResult: ...

def normalize_header(header: str) -> str:
    """strip → lowercase → spaces/special chars → underscores → deduplicate underscores"""
```

### 3.8 Module: `mapping.py` — ORM Schema Extraction

#### Purpose
Introspects SQLAlchemy ORM models and produces `ExcelMapping` dataclasses that drive all downstream operations.

#### Key Types

```python
@dataclass(frozen=True)
class ColumnMapping:
    name: str                          # ORM column name
    excel_header: str                  # Display text for Excel header
    python_type: type[object]          # Inferred Python type
    sqla_type: TypeEngine[object]      # SQLAlchemy type object
    nullable: bool                     # NULL allowed
    primary_key: bool                  # Primary key flag
    has_default: bool                  # Default value exists
    default_value: object | None       # Static default (None if callable)
    enum_values: list[str] | None      # Enum dropdown options
    max_length: int | None             # String(N) max length
    description: str | None            # Column doc/comment
    foreign_key: str | None            # Referenced table.column

@dataclass(frozen=True)
class ExcelMapping:
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
        key_columns: list[str] | None = None,
        include: list[str] | None = None,
        exclude: list[str] | None = None,
        header_map: dict[str, str] | None = None,
    ) -> ExcelMapping: ...
```

#### Type Mapping Table (`_TYPE_MAP`)

```python
_TYPE_MAP: tuple[tuple[type[TypeEngine], type], ...] = (
    (Integer, int),
    (Float, float),
    (String, str),
    (Text, str),
    (Boolean, bool),
    (Date, date),
    (DateTime, datetime),
)
```

#### `_python_type_for_sqla_type()` — Special Cases

```python
def _python_type_for_sqla_type(sqla_type: TypeEngine) -> type:
    # Numeric special handling
    if isinstance(sqla_type, Numeric):
        if getattr(sqla_type, "asdecimal", True):
            return Decimal
        return float

    # Tuple-based lookup (order matters)
    for sa_type, py_type in _TYPE_MAP:
        if isinstance(sqla_type, sa_type):
            return py_type

    return str  # Fallback
```

#### Introspection Strategy

```
1. sa_inspect(model) → mapper
2. mapper.columns iteration → column metadata extraction
3. _TYPE_MAP-based type mapping + Numeric/Enum special cases
4. _extract_enum_values() → enum_class.members or sqla enum.enums
5. include/exclude filter application (mutually exclusive → MappingError)
6. header_map override application
7. key_columns: if not specified → primary key columns auto-detected
8. _default_excel_header() → column_name.replace("_", " ").title()
```

### 3.9 Module: `template.py` — Template Generation

#### Purpose
Converts `ExcelMapping` to formatted `.xlsx` template files. Uses `ExcelWorkbookSession` for workbook creation.

#### excel-dbapi Integration

```python
class ExcelTemplate:
    def __init__(
        self,
        mappings: list[ExcelMapping],
        *,
        include_sample_data: bool = False,
    ) -> None: ...

    def save(self, path: str | Path) -> None:
        with ExcelWorkbookSession.open(path, create=True) as session:
            self._populate_workbook(session.workbook)
            session.commit()

    def to_bytesio(self) -> BytesIO:
        # 1. Create temp file
        # 2. self.save(temp_path) → ExcelWorkbookSession
        # 3. Read temp file → BytesIO
        # 4. Clean up temp file

    def to_bytes(self) -> bytes:
        return self.to_bytesio().getvalue()
```

#### Style Constants

```python
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
REQUIRED_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
HEADER_BORDER = Border(
    bottom=Side(style="thin", color="2F5496"),
    right=Side(style="thin", color="D6DCE4"),
)
```

#### Template Generation Pipeline

```
ExcelWorkbookSession.open(path, create=True)
    │
    ▼
session.workbook → openpyxl.Workbook
    │
    ▼
Delete all existing sheets
    │
    ▼
For each ExcelMapping:
    ├── workbook.create_sheet(title=sheet_name)
    ├── Header row (styled: font, fill, border, alignment)
    │   ├── Required columns → yellow background (FFF2CC)
    │   └── Optional columns → blue background (4472C4)
    ├── Cell comments (type, nullable, PK, FK, description)
    ├── Column widths auto-adjustment
    ├── Enum columns → DataValidation dropdown
    │   └── 255-char limit check → fallback to comment if exceeded
    ├── Auto-filter
    ├── Freeze panes (A2)
    └── Sample data row (optional)
    │
    ▼
session.commit() → atomic disk save (tempfile + os.replace)
```

#### DataValidation 255-Character Limit

```python
# If enum values fit within 255 characters
dv = DataValidation(type="list", formula1=f'"{joined_values}"')
ws.add_data_validation(dv)

# If exceeded → fall back to comment
comment = Comment(f"Allowed values: {joined_values}", "sqlalchemy-excel")
cell.comment = comment
```

### 3.10 Module: `validation/engine.py` — ExcelValidator

#### Purpose
Orchestrates Excel file validation using `ExcelDbapiReader` for data reading and `PydanticBackend` for row-level validation.

#### Implementation

```python
class ExcelValidator:
    def __init__(
        self,
        mappings: list[ExcelMapping],
        *,
        backend: str = "pydantic",
    ) -> None:
        self._mappings = mappings
        self._reader = ExcelDbapiReader(read_only=True)
        self._backends = {
            mapping.sheet_name: PydanticBackend(mapping)
            for mapping in mappings
        }

    def validate(
        self,
        source: str | Path | BinaryIO,
        *,
        sheet_name: str | None = None,
        max_errors: int | None = None,
        stop_on_first_error: bool = False,
    ) -> ValidationReport: ...
```

#### Validation Pipeline

```
.xlsx file (upload)
    │
    ▼
ExcelDbapiReader.read(source, sheet_name)
    │ → SELECT * FROM SheetName via excel-dbapi
    ▼
ReaderResult(headers, rows, total_rows)
    │
    ▼
_build_header_map() → header → column name mapping
    ├── normalize_header(column.name)
    └── normalize_header(column.excel_header)
    │
    ▼
For each row:
    ├── _remap_row() → remap headers to ORM column names
    ├── PydanticBackend.validate_row(row_data, row_number)
    │   ├── _coerce_value() → lightweight pre-Pydantic coercion
    │   ├── Pydantic model_validate() → strict type validation
    │   └── PydanticValidationError → CellError conversion
    ├── max_errors check → stop if exceeded
    └── stop_on_first_error check → stop on first invalid row
    │
    ▼
ValidationReport(errors, total_rows, valid_rows, invalid_rows)
```

#### Helper Functions

| Function | Purpose |
|----------|---------|
| `_select_mapping()` | Select appropriate mapping for sheet |
| `_build_header_map()` | Build normalized header → column name mapping |
| `_remap_row()` | Remap row dict keys from Excel headers to ORM column names |
| `_reset_source_cursor()` | Reset BinaryIO source position for re-reading |
| `_build_reader()` | Create `ExcelDbapiReader` instance |

### 3.11 Module: `validation/pydantic_backend.py` — Dynamic Model Generation

#### Purpose
Generates dynamic Pydantic v2 models from `ExcelMapping` for per-row validation.

#### Dynamic Model Construction

```python
def _create_pydantic_model(mapping: ExcelMapping) -> type[BaseModel]:
    field_definitions = {}
    for column in mapping.columns:
        field_type = _field_type_for_column(column)
        # Enum → Literal[values], nullable → type | None
        field_info = Field(
            default=None if column.nullable else ...,
            max_length=column.max_length,
        )
        field_definitions[column.name] = (field_type, field_info)

    return create_model(f"{model_name}Validator", **field_definitions)
```

#### `_field_type_for_column()` — Type Resolution

```python
def _field_type_for_column(column: ColumnMapping) -> type:
    if column.enum_values:
        return Literal[tuple(column.enum_values)]  # Enum → Literal
    base_type = column.python_type
    if column.nullable:
        return base_type | None
    return base_type
```

#### `_coerce_value()` — Pre-Pydantic Coercion

Lightweight type coercion before Pydantic validation:

```python
def _coerce_value(value: object, column: ColumnMapping) -> object:
    if value == "" or value is None:
        return None

    target = column.python_type
    if isinstance(value, str):
        if target is int:       return int(value)
        if target is float:     return float(value)
        if target is Decimal:   return Decimal(value)
        if target is date:      return date.fromisoformat(value)
        if target is datetime:  return datetime.fromisoformat(value)
        if target is bool:      return value.lower() in ("true", "yes", "1")

    if isinstance(value, enum.Enum):
        return value.value

    return value
```

#### `_map_error_code()` — Error Code Mapping

Maps Pydantic error types to project-specific error codes:

| Pydantic Error Type | Project Error Code |
|--------------------|-------------------|
| `missing`, `value_error.missing` | `null_error` |
| `int_parsing`, `float_parsing`, `*_type` | `type_error` |
| `string_too_long`, `*_length` | `length_error` |
| `literal_error`, `*_enum` | `enum_error` |
| *(other)* | `constraint_error` |

### 3.12 Module: `validation/report.py` — CellError & ValidationReport

#### Implementation

```python
@dataclass(frozen=True)
class CellError:
    row: int              # Excel row number (1-based, includes header offset)
    column: str           # Column name or header
    value: Any            # Raw cell value that failed validation
    expected_type: str    # Human-readable expected type description
    message: str          # Descriptive error message
    error_code: str       # Machine-readable: null_error, type_error, length_error,
                          #   enum_error, constraint_error

@dataclass
class ValidationReport:
    errors: list[CellError]
    total_rows: int
    valid_rows: int
    invalid_rows: int

    @property
    def has_errors(self) -> bool:
        return len(self.errors) > 0

    def summary(self) -> str:
        return (
            f"Validated {self.total_rows} rows: "
            f"{self.valid_rows} valid, {self.invalid_rows} invalid. "
            f"{len(self.errors)} errors found."
        )

    def to_dict(self) -> dict:       # JSON-serializable dictionary
    def errors_by_row(self) -> dict:  # Group errors by Excel row number
    def to_excel(self, path) -> None: # Export error report to Excel file
```

### 3.13 Module: `load/strategies.py` — Import Strategies

#### Purpose
Strategy pattern implementations for Insert, Upsert, and DryRun modes with savepoint-based error recovery.

#### Shared Utilities

```python
def _chunk(iterable: Iterable[_T], size: int) -> Iterable[list[_T]]:
    """Yield lists of at most `size` items."""
    if size < 1:
        raise ValueError("Chunk size must be at least 1")
    batch = []
    for item in iterable:
        batch.append(item)
        if len(batch) >= size:
            yield batch
            batch = []
    if batch:
        yield batch

def _build_key_filter(
    row: dict[str, object],
    key_columns: list[str],
) -> tuple[dict[str, object], list[str]]:
    """Extract key filter dict and list of missing keys from a row."""
    key_filter = {}
    missing_keys = []
    for key in key_columns:
        if key not in row:
            missing_keys.append(key)
            continue
        key_filter[key] = row[key]
    return key_filter, missing_keys
```

#### `ImportResult` Dataclass

```python
@dataclass(slots=True)
class ImportResult:
    inserted: int = 0
    updated: int = 0
    skipped: int = 0
    failed: int = 0
    errors: list[str] = field(default_factory=list)
    duration_ms: float = 0.0

    @property
    def total(self) -> int:
        return self.inserted + self.updated + self.skipped + self.failed

    def summary(self) -> str:
        return (
            f"inserted={self.inserted}, updated={self.updated}, "
            f"skipped={self.skipped}, failed={self.failed}, total={self.total}, "
            f"errors={len(self.errors)}, duration_ms={self.duration_ms:.2f}"
        )
```

#### `LoadStrategy` Protocol

```python
class LoadStrategy(Protocol):
    def execute(
        self,
        session: Session,
        model_class: type[DeclarativeBase],
        rows: Iterable[dict[str, object]],
        key_columns: list[str],
        batch_size: int,
    ) -> ImportResult: ...
```

#### InsertStrategy — Batched ORM Insert

```python
class InsertStrategy:
    def execute(self, session, model_class, rows, key_columns, batch_size):
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
            else:
                if savepoint.is_active:
                    savepoint.commit()
        return result
```

#### UpsertStrategy — Per-Row Lookup + Insert/Update

```python
class UpsertStrategy:
    def execute(self, session, model_class, rows, key_columns, batch_size):
        if not key_columns:
            raise ValueError("Upsert strategy requires at least one key column")

        result = ImportResult()
        for batch in _chunk(rows, batch_size):
            savepoint = session.begin_nested()
            try:
                for row in batch:
                    key_filter, missing_keys = _build_key_filter(row, key_columns)
                    if missing_keys:
                        result.failed += 1
                        result.errors.append(f"Missing upsert key column(s): ...")
                        continue

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
            except IntegrityError as exc:
                if savepoint.is_active:
                    savepoint.rollback()
                self._recover_failed_batch(...)
            else:
                if savepoint.is_active:
                    savepoint.commit()
        return result
```

#### `_recover_failed_batch()` — Per-Row Retry with Counter Correction

When a batch fails with `IntegrityError`, `UpsertStrategy` retries each row individually:

```python
def _recover_failed_batch(self, *, session, model_class, batch, key_columns, result, initial_error):
    result.errors.append(str(initial_error))

    # Step 1: Count rows already counted in the failed batch
    inserted_in_batch = 0
    updated_in_batch = 0
    for row in batch:
        key_filter, missing = _build_key_filter(row, key_columns)
        if missing: continue
        existing = session.execute(
            select(model_class).filter_by(**key_filter)
        ).scalar_one_or_none()
        if existing is None:
            inserted_in_batch += 1
        else:
            updated_in_batch += 1

    # Step 2: Subtract already-counted values (they were rolled back)
    result.inserted -= inserted_in_batch
    result.updated -= updated_in_batch

    # Step 3: Retry each row individually with its own savepoint
    for row in batch:
        savepoint = session.begin_nested()
        try:
            # Same upsert logic per row
            session.flush()
        except Exception as exc:
            result.failed += 1
            result.errors.append(str(exc))
            if savepoint.is_active:
                savepoint.rollback()
        else:
            if savepoint.is_active:
                savepoint.commit()
```

**Why counter correction?** The batch-level loop already incremented `result.inserted` and `result.updated` for rows before the `IntegrityError`. Since the entire batch was rolled back, those counts must be subtracted before re-processing.

#### DryRunStrategy — Non-Persistent Validation

```python
class DryRunStrategy:
    def execute(self, session, model_class, rows, key_columns, batch_size):
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
            finally:
                if savepoint.is_active:
                    savepoint.rollback()  # ALWAYS rollback — never persist
        return result
```

**Key difference**: `finally` block always rolls back the savepoint, regardless of success or failure. No data is persisted.

### 3.14 Module: `load/importer.py` — ExcelImporter

#### Purpose
Orchestrates Excel-to-database import using readers, validators, and load strategies.

#### Implementation

```python
class ExcelImporter:
    def __init__(
        self,
        mappings: list[ExcelMapping],
        session: Session,
    ) -> None:
        if not mappings:
            raise ImportError_("At least one mapping is required")
        self._mappings = mappings
        self._session = session

    def insert(self, source, *, batch_size=1000, validate=True) -> ImportResult:
        return self._run(InsertStrategy(), source, batch_size=batch_size, validate=validate)

    def upsert(self, source, *, batch_size=1000, validate=True) -> ImportResult:
        return self._run(UpsertStrategy(), source, batch_size=batch_size, validate=validate)

    def dry_run(self, source, *, validate=True) -> ImportResult:
        return self._run(DryRunStrategy(), source, batch_size=1000, validate=validate)
```

#### Import Pipeline (`_run` method)

```
Excel source
    │
    ▼
[validate=True] → ExcelValidator.validate(source)
    ├── has_errors → ImportResult(failed=invalid_rows, errors=...) early return
    └── Validation passed → continue
    │
    ▼
ExcelDbapiReader(read_only=True).read(source, sheet_name)
    │ → SELECT * FROM SheetName via excel-dbapi
    ▼
_extract_rows_for_mapping() → _align_row() applied
    ├── normalize_header() for header matching
    └── Map row keys to ORM column names
    │
    ▼
strategy.execute(session, model_class, rows, key_columns, batch_size)
    │
    ▼
ImportResult(inserted, updated, skipped, failed, errors, duration_ms)
```

#### Key Helper Methods

| Method | Purpose |
|--------|---------|
| `_create_reader()` | Creates `ExcelDbapiReader(read_only=True)` |
| `_create_validator()` | Creates `ExcelValidator(self._mappings)` |
| `_align_row(row, mapping)` | Normalize headers, match to column.name or column.excel_header |
| `_extract_rows_for_mapping()` | Handle both ReaderResult and raw Iterable |
| `_reset_source_cursor()` | Call seek(0) if available (for BinaryIO re-reading) |
| `_merge_result()` | Merge per-mapping results into aggregate ImportResult |

#### Design Decisions
- **Transaction boundary is caller's responsibility** — importer uses `session.flush()` + savepoints internally but never calls `session.commit()`
- **Timing** — `perf_counter` measures total operation duration → `duration_ms`
- **Validation gate** — optional pre-import validation prevents bad data from reaching the DB

### 3.15 Module: `export.py` — ExcelExporter

#### Purpose
Exports SQLAlchemy query results to formatted Excel files using `ExcelWorkbookSession`.

#### Implementation

```python
class ExcelExporter:
    def __init__(self, mappings: list[ExcelMapping]) -> None:
        if not mappings:
            raise ExportError("At least one mapping is required")
        self._mappings = mappings

    def export(
        self,
        rows: Sequence[Any],
        path: str | Path | None = None,
        *,
        sheet_name: str | None = None,
    ) -> bytes | None:
        if path is not None:
            with ExcelWorkbookSession.open(path, create=True) as session:
                self._populate_workbook(session.workbook, rows, sheet_name)
                session.commit()
            return None

        # path=None → temp file → return bytes
        ...
```

#### `_extract_value()` — Dual-Source Value Extraction

```python
def _extract_value(row: object, column_name: str) -> object:
    # ORM instances → getattr
    if hasattr(row, column_name):
        value = getattr(row, column_name)
    # Dicts → dict access
    elif isinstance(row, dict):
        value = row.get(column_name)
    else:
        return None

    # String sanitization for formula injection prevention
    if isinstance(value, str):
        return sanitize_cell_value(value)
    return value
```

#### Export Features
- Header styling (blue background, white bold font, centered)
- Date formatting (`YYYY-MM-DD HH:MM:SS` for datetime, `YYYY-MM-DD` for date)
- Auto-width column adjustment (capped at 50 characters)
- Auto-filter and freeze panes
- ORM instance and dict row support
- `sanitize_cell_value()` for formula injection prevention

### 3.16 Module: `cli.py` — CLI Interface

#### Purpose
Click-based CLI with 5 commands for scripting and automation.

#### Model Resolution

```python
def _resolve_model(dotpath: str) -> type:
    """Parse 'module.path:ClassName' and import the model class."""
    module_path, _, class_name = dotpath.partition(":")
    if not class_name:
        raise click.BadParameter("Expected format: module.path:ClassName")
    module = importlib.import_module(module_path)
    return getattr(module, class_name)
```

#### Commands

| Command | Options | Description |
|---------|---------|-------------|
| `template` | `--model`, `--output`, `--sample-data`, `--sheet-name` | Generate Excel template |
| `validate` | `--model`, `--input`, `--format`, `--output` | Validate uploaded file |
| `import` | `--model`, `--input`, `--db`, `--mode`, `--dry-run`, `--batch-size` | Import to database |
| `export` | `--model`, `--db`, `--output` | Export query results |
| `inspect` | `--input` | Inspect Excel file structure |

**Note**: `inspect` uses openpyxl directly (`read_only=True`, `data_only=True`) to show sheet structure without requiring a model.

### 3.17 Module: `integrations/fastapi.py` — FastAPI Router Factory

#### Purpose
Auto-generates REST endpoints from a single ORM model class.

```python
def create_import_router(
    model: type,
    *,
    prefix: str = "",
    tags: list[str] | None = None,
    session_dependency: Any = None,
) -> APIRouter:
```

#### Generated Endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET` | `{prefix}/template` | Download Excel template (with sample data) |
| `POST` | `{prefix}/validate` | Upload and validate → JSON report |
| `POST` | `{prefix}/import` | Validate + DB import (requires `session_dependency`) |
| `GET` | `{prefix}/health` | Health check: `{"status": "ok", "model": "ClassName"}` |

---

## 4. Integration Architecture: excel-dbapi

### 4.1 Dual-Channel Data Flow

```
┌──────────────────────────────────────────────────────────────┐
│  excel-dbapi Usage in sqlalchemy-excel                        │
│                                                               │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │  Format Channel (Write)                                  │ │
│  │                                                          │ │
│  │  ExcelWorkbookSession                                    │ │
│  │  ├── conn = excel_dbapi.connect(path, create=True)      │ │
│  │  ├── conn.workbook → openpyxl.Workbook                  │ │
│  │  │   ├── cell.font = Font(bold=True)                    │ │
│  │  │   ├── cell.fill = PatternFill(...)                   │ │
│  │  │   └── DataValidation, comments, freeze_panes         │ │
│  │  └── conn.commit() → atomic disk save                   │ │
│  │                                                          │ │
│  │  Used by: template.py, export.py                        │ │
│  └─────────────────────────────────────────────────────────┘ │
│                                                               │
│  ┌─────────────────────────────────────────────────────────┐ │
│  │  Data Channel (Read)                                     │ │
│  │                                                          │ │
│  │  ExcelDbapiReader                                        │ │
│  │  ├── conn = excel_dbapi.connect(path, data_only=True)   │ │
│  │  ├── cursor = conn.cursor()                              │ │
│  │  ├── cursor.execute("SELECT * FROM SheetName")           │ │
│  │  │   └── Table names UNQUOTED!                           │ │
│  │  ├── cursor.description → column metadata                │ │
│  │  └── cursor.fetchall() → List[Tuple]                    │ │
│  │                                                          │ │
│  │  Used by: validation/engine.py, load/importer.py        │ │
│  └─────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────┘
```

### 4.2 Contract Requirements from excel-dbapi

sqlalchemy-excel depends on these guaranteed interfaces:

| Interface | Guarantee |
|-----------|-----------|
| `connect(path, engine="openpyxl", create=True)` | Returns valid ExcelConnection |
| `conn.workbook` | Returns openpyxl Workbook (openpyxl engine only) |
| `conn.cursor()` | Returns PEP 249 Cursor |
| `cursor.execute("SELECT * FROM Sheet")` | Populates `description` and result set |
| `cursor.description` | 7-tuple format per PEP 249 |
| `cursor.fetchall()` | Returns `List[Tuple]` |
| `create=True` | Creates valid empty workbook if file missing |
| `conn.commit()` | Atomic save (tempfile + os.replace) |
| All exceptions | Follow PEP 249 hierarchy |

### 4.3 Exception Mapping

```
excel-dbapi Exception          → sqlalchemy-excel Exception
─────────────────────────────────────────────────────────
InvalidFileException, BadZipFile → FileFormatError
Sheet not in workbook.sheetnames → SheetNotFoundError
Connection failure               → ReaderError
Generic Exception                → ReaderError (wrapped)
```

---

## 5. pyproject.toml Configuration

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
authors = [{name = "Yeongseon Choe"}]
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
    "excel-dbapi>=1.0",
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
    "httpx>=0.27",
]
all = [
    "pandas>=2.0",
    "pandera>=0.18",
    "fastapi>=0.100",
    "python-multipart>=0.0.5",
    "pytest>=8.0",
    "pytest-cov>=4.0",
    "hypothesis>=6.0",
    "ruff>=0.4",
    "mypy>=1.10",
    "httpx>=0.27",
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
select = ["E", "W", "F", "I", "N", "UP", "B", "SIM", "TCH", "RUF"]
ignore = ["E501"]

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

[[tool.mypy.overrides]]
module = "excel_dbapi.*"
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

### Dependency Rationale

| Dependency | Why Required |
|-----------|-------------|
| `sqlalchemy>=2.0` | ORM schema introspection (2.0-style `Mapped[]`, `mapped_column()`), DB operations |
| `excel-dbapi>=1.0` | Full dependency: all Excel I/O (SQL reads, workbook creation, atomic saves) |
| `openpyxl>=3.1` | Cell-level formatting via `ExcelWorkbookSession.workbook` |
| `pydantic>=2.0` | Row-level validation with dynamic model generation |
| `defusedxml>=0.7` | Security: XML bomb / XXE prevention for untrusted uploads |
| `click>=8.0` | CLI interface with 5 commands |

---

## 6. Security Implementation

### 6.1 Threats & Mitigations

| Threat | Mitigation |
|--------|-----------|
| XML Bomb (billion laughs) | `defusedxml` required dependency, verified at import time |
| XXE (External Entity Injection) | `defusedxml` patches openpyxl's XML parsing |
| Formula injection (`=`, `+`, `-`, `@`, `\t`, `\r`) | `sanitize_cell_value()`: prefixes with `'` |
| File size DoS | `ExcelDbapiReader.max_file_size` (default 50MB) |
| Zip bomb | `_validate_file_size_path()` + openpyxl internal protection |
| Path traversal | File paths validated, uploads use tempfile |
| SQL injection | All DB operations via SQLAlchemy ORM (parameterized) |

### 6.2 Security Initialization

```python
# __init__.py — executed at package import time
from sqlalchemy_excel._compat import ensure_defusedxml
ensure_defusedxml()  # Raises ImportError immediately if defusedxml missing
```

### 6.3 Cell Value Sanitization

```python
# _compat.py
_FORMULA_PREFIXES = ("=", "+", "-", "@", "\t", "\r")

def sanitize_cell_value(value: object) -> object:
    """Prevent formula injection by prefixing dangerous values with apostrophe."""
    if isinstance(value, str) and value.startswith(_FORMULA_PREFIXES):
        return f"'{value}"
    return value
```

Used in: `template.py` (sample data), `export.py` (exported values), `report.py` (error report).

### 6.4 File Size Validation

```python
# reader/excel_dbapi_reader.py
def _validate_file_size_path(self, path: str) -> None:
    size = os.path.getsize(path)
    if size > self.max_file_size:
        raise ReaderError(
            f"File size ({size:,} bytes) exceeds maximum "
            f"({self.max_file_size:,} bytes)"
        )
```

---

## 7. API Usage Examples

### 7.1 Complete Workflow — Template → Validate → Import

```python
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column

from sqlalchemy_excel import (
    ExcelMapping,
    ExcelTemplate,
    ExcelValidator,
    ExcelImporter,
)


class Base(DeclarativeBase):
    pass

class User(Base):
    __tablename__ = "users"
    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column()
    email: Mapped[str] = mapped_column()
    age: Mapped[int | None] = mapped_column(default=None)


# 1. Create mapping from ORM model
mapping = ExcelMapping.from_model(User, key_columns=["email"])

# 2. Generate template
template = ExcelTemplate([mapping], include_sample_data=True)
template.save("users_template.xlsx")

# 3. Validate uploaded file
validator = ExcelValidator([mapping])
report = validator.validate("users_upload.xlsx")
if report.has_errors:
    print(report.summary())
    report.to_excel("validation_errors.xlsx")
    raise SystemExit(1)

# 4. Import to database
engine = create_engine("sqlite:///app.db")
Base.metadata.create_all(engine)
with Session(engine) as session:
    importer = ExcelImporter([mapping], session=session)
    result = importer.upsert("users_upload.xlsx")
    session.commit()
    print(result.summary())
```

### 7.2 Mapping Configuration Examples

```python
# Basic: all columns, default settings
mapping = ExcelMapping.from_model(User)

# Custom sheet name
mapping = ExcelMapping.from_model(User, sheet_name="Employee List")

# Include only specific columns
mapping = ExcelMapping.from_model(User, include=["name", "email", "age"])

# Exclude auto-generated columns
mapping = ExcelMapping.from_model(User, exclude=["id", "created_at"])

# Custom Excel headers
mapping = ExcelMapping.from_model(
    User,
    header_map={"name": "Full Name", "email": "Email Address"},
)

# Custom upsert key (instead of primary key)
mapping = ExcelMapping.from_model(User, key_columns=["email"])
```

### 7.3 Template Generation

```python
from sqlalchemy_excel import ExcelMapping, ExcelTemplate

mapping = ExcelMapping.from_model(User)

# Save to disk
template = ExcelTemplate([mapping])
template.save("template.xlsx")

# Get as bytes (for HTTP responses)
xlsx_bytes = template.to_bytes()

# Get as BytesIO stream
stream = template.to_bytesio()

# Multi-sheet template
user_mapping = ExcelMapping.from_model(User)
order_mapping = ExcelMapping.from_model(Order)
template = ExcelTemplate([user_mapping, order_mapping])
template.save("multi_sheet_template.xlsx")
```

### 7.4 Validation with Error Handling

```python
from sqlalchemy_excel import ExcelMapping, ExcelValidator

mapping = ExcelMapping.from_model(User)
validator = ExcelValidator([mapping])

# Validate from file path
report = validator.validate("upload.xlsx")

# Validate from BytesIO (e.g., HTTP upload)
from io import BytesIO
report = validator.validate(BytesIO(file_bytes))

if report.has_errors:
    # Summary
    print(report.summary())
    # "Validated 100 rows: 95 valid, 5 invalid. 8 errors found."

    # Iterate errors by row
    for row_num, errors in report.errors_by_row().items():
        print(f"Row {row_num}:")
        for e in errors:
            print(f"  {e.column}: {e.message} (code: {e.error_code})")

    # JSON serialization (for API responses)
    import json
    print(json.dumps(report.to_dict(), indent=2, default=str))

    # Export error report to Excel
    report.to_excel("validation_errors.xlsx")
```

### 7.5 Import Strategies

```python
from sqlalchemy import create_engine
from sqlalchemy.orm import Session
from sqlalchemy_excel import ExcelMapping, ExcelImporter

mapping = ExcelMapping.from_model(User, key_columns=["email"])
engine = create_engine("sqlite:///app.db")

with Session(engine) as session:
    importer = ExcelImporter([mapping], session=session)

    # Insert mode (new rows only)
    result = importer.insert("users.xlsx", batch_size=500)
    session.commit()

    # Upsert mode (update existing, insert new)
    result = importer.upsert("users.xlsx")
    session.commit()

    # Dry-run mode (simulate without persisting)
    result = importer.dry_run("users.xlsx")
    print(result.summary())
    # No session.commit() needed — changes already rolled back
```

### 7.6 Export

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

### 7.7 FastAPI Integration

```python
from fastapi import FastAPI
from sqlalchemy import create_engine
from sqlalchemy.orm import Session

from sqlalchemy_excel.integrations.fastapi import create_import_router
from myapp.models import User

engine = create_engine("sqlite:///app.db")

def get_session():
    with Session(engine) as session:
        yield session

app = FastAPI()

router = create_import_router(
    User,
    prefix="/users",
    tags=["users"],
    session_dependency=get_session,
)
app.include_router(router)

# Endpoints created:
# GET  /users/template  → Download Excel template
# POST /users/validate  → Upload and validate → JSON report
# POST /users/import    → Validate + import → result summary
# GET  /users/health    → {"status": "ok", "model": "User"}
```

### 7.8 CLI Usage

```bash
# Generate template
sqlalchemy-excel template --model myapp.models:User --output users.xlsx --sample-data

# Validate upload
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx --format json
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx \
    --format excel --output errors.xlsx

# Import to DB
sqlalchemy-excel import --model myapp.models:User \
    --input upload.xlsx \
    --db sqlite:///app.db \
    --mode upsert \
    --batch-size 500

# Dry-run import
sqlalchemy-excel import --model myapp.models:User \
    --input upload.xlsx \
    --db sqlite:///app.db \
    --dry-run

# Export from DB
sqlalchemy-excel export --model myapp.models:User \
    --db sqlite:///app.db \
    --output export.xlsx

# Inspect Excel file structure (no model needed)
sqlalchemy-excel inspect --input mystery.xlsx
```

---

## 8. Sample ORM Models for Testing

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

---

## 9. Testing Strategy

### 9.1 Test Structure

```
tests/
├── conftest.py              # Shared fixtures
│   ├── in_memory_engine     # SQLite :memory:
│   ├── sample_models        # User, Product, Order models
│   └── temp_workbook        # Temporary .xlsx creation helper
├── unit/
│   ├── test_mapping.py      # ExcelMapping extraction
│   ├── test_template.py     # Template generation
│   ├── test_reader.py       # Excel file parsing
│   ├── test_validation.py   # Validation engine
│   ├── test_importer.py     # DB import
│   ├── test_report.py       # ValidationReport
│   └── test_export.py       # Query export
├── integration/
│   └── test_end_to_end.py   # Full pipeline
└── fixtures/                # Test .xlsx files
```

### 9.2 Test Summary

| Category | Files | Count | Description |
|----------|-------|-------|-------------|
| Mapping tests | `test_mapping.py` | ~15 | Model introspection, type mapping, include/exclude, header_map |
| Template tests | `test_template.py` | ~15 | Workbook creation, styles, DataValidation, sample data |
| Reader tests | `test_reader.py` | ~18 | ExcelDbapiReader, BinaryIO, file size, headers, empty rows |
| Validation tests | `test_validation.py` | ~20 | Pydantic backend, type coercion, error codes, report |
| Importer tests | `test_importer.py` | ~20 | Insert/upsert/dry-run, batch processing, error recovery |
| Report tests | `test_report.py` | ~10 | CellError, ValidationReport, summary, to_dict, errors_by_row |
| Export tests | `test_export.py` | ~10 | File export, bytes export, date formatting, sanitization |
| Integration tests | `test_end_to_end.py` | ~9 | Full pipeline: template → fill → validate → import |
| **Total** | **8 files** | **117** | All passing ✅ |

### 9.3 Test Invariants (Hypothesis)

1. **Round-trip**: `template → fill with valid data → validate → no errors`
2. **Idempotent validation**: `validate(data) == validate(data)` (same report)
3. **Export/Import consistency**: `export(query) → import(file) → same data in DB`

### 9.4 Key Test Patterns

#### Mapping Tests

```python
def test_from_model_basic():
    mapping = ExcelMapping.from_model(User)
    assert mapping.sheet_name == "users"
    assert len(mapping.columns) > 0
    assert any(c.primary_key for c in mapping.columns)

def test_from_model_include_exclude_mutual():
    with pytest.raises(MappingError):
        ExcelMapping.from_model(User, include=["name"], exclude=["id"])

def test_type_mapping_numeric_decimal():
    # Numeric(asdecimal=True) → Decimal
    mapping = ExcelMapping.from_model(ModelWithDecimal)
    col = next(c for c in mapping.columns if c.name == "price")
    assert col.python_type is Decimal
```

#### Import Strategy Tests

```python
def test_insert_integrity_error(session, mapping):
    """IntegrityError in a batch → rollback + record failed."""
    importer = ExcelImporter([mapping], session=session)
    result = importer.insert("duplicate_keys.xlsx")
    assert result.failed > 0
    assert len(result.errors) > 0

def test_dry_run_no_persist(session, mapping):
    """DryRun counts rows but does not persist any data."""
    importer = ExcelImporter([mapping], session=session)
    result = importer.dry_run("valid_data.xlsx")
    assert result.inserted > 0
    # Verify no data actually in DB
    count = session.query(User).count()
    assert count == 0
```

---

## 10. CI/CD Configuration

### 10.1 GitHub Actions CI

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
      - run: pip install ruff
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

### 10.2 GitHub Actions Publish

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

---

## 11. Known Issues and Workarounds

| Issue | Workaround |
|-------|-----------|
| excel-dbapi quoted table names fail | Use unquoted: `Sheet1` not `"Sheet1"` |
| BinaryIO requires temp file for excel-dbapi | `_resolve_source()` handles conversion + cleanup |
| DataValidation is client-side only | Server-side `ExcelValidator` is always required |
| Large files (>50MB) rejected by default | Configure `max_file_size` on `ExcelDbapiReader` |
| `pandas` optional but some tests may need it | Install with `pip install sqlalchemy-excel[all]` |
| openpyxl `read_only=True` not used via excel-dbapi | excel-dbapi manages openpyxl mode internally |

---

## 12. Interoperability with excel-dbapi

### 12.1 Integration Points

```python
# ExcelWorkbookSession wraps excel-dbapi for write operations
from excel_dbapi import connect

class ExcelWorkbookSession:
    @classmethod
    def open(cls, path, *, create=False, data_only=False):
        conn = connect(
            str(path),
            engine="openpyxl",
            autocommit=False,
            create=create,
            data_only=data_only,
        )
        return cls(conn)

    @property
    def workbook(self):
        return self._conn.workbook  # openpyxl Workbook

    def commit(self):
        self._conn.commit()  # Atomic save
```

```python
# ExcelDbapiReader wraps excel-dbapi for read operations
class ExcelDbapiReader:
    def read(self, source, sheet_name=None):
        conn = connect(str(file_path), engine="openpyxl", autocommit=True, data_only=True)
        cursor = conn.cursor()
        cursor.execute(f"SELECT * FROM {sheet_name}")
        headers = [desc[0] for desc in cursor.description]
        rows = cursor.fetchall()
        return ReaderResult(headers, [dict(zip(headers, row)) for row in rows])
```

### 12.2 Dependency Specification

```toml
# In sqlalchemy-excel's pyproject.toml
[project]
dependencies = [
    "excel-dbapi>=1.0",
    ...
]
```

### 12.3 excel-dbapi API Contract

excel-dbapi guarantees these interfaces for sqlalchemy-excel:

| Interface | Guarantee |
|-----------|-----------|
| `connect(path, engine="openpyxl", create=True)` | Returns valid ExcelConnection |
| `conn.workbook` | Returns openpyxl Workbook (openpyxl engine only) |
| `conn.cursor()` | Returns PEP 249 Cursor |
| `cursor.execute("SELECT * FROM Sheet")` | Populates `description` and result set |
| `cursor.description` | 7-tuple format per PEP 249 |
| `cursor.fetchall()` | Returns `List[Tuple]` |
| `create=True` | Creates valid empty workbook if file missing |
| All exceptions | Follow PEP 249 hierarchy |

---

## 13. Future Technical Design (v0.2.x)

### 13.1 Pandera Integration (Optional Extra)

```python
# Planned: DataFrame-level cross-column validation
from sqlalchemy_excel import ExcelValidator

validator = ExcelValidator([mapping], backend="pandera")
report = validator.validate("upload.xlsx")
```

### 13.2 Streaming Reader for Large Files

```python
# Planned: Streaming row access for files > 50MB
reader = ExcelDbapiReader(streaming=True, max_file_size=500 * 1024 * 1024)
result = reader.read("large_file.xlsx")
for row in result.rows:  # Generator, not list
    process(row)
```

### 13.3 Async Support

```python
# Planned: Async import for web applications
async with AsyncSession(engine) as session:
    importer = AsyncExcelImporter([mapping], session=session)
    result = await importer.insert("upload.xlsx")
    await session.commit()
```
