# AGENTS.md — sqlalchemy-excel

> Project knowledge base for AI agents. Last updated: 2026-03-15.

## Project Identity

- **Name**: sqlalchemy-excel
- **Package**: `sqlalchemy-excel` (PyPI), import as `sqlalchemy_excel`
- **License**: MIT (Copyright 2025 Yeongseon Choe)
- **Repository**: https://github.com/yeongseon/sqlalchemy-excel
- **Python**: 3.10–3.13
- **Version**: 0.1.0
- **Stage**: MVP Complete

## One-Line Description

SQLAlchemy model-driven Excel template generation, server-side validation, and database import toolkit powered by excel-dbapi.

## What This Project Does

sqlalchemy-excel treats SQLAlchemy ORM models as the **single source of truth** for:

1. **Template Generation** — Create Excel files with correct headers, types, dropdowns, sample data
2. **Server Validation** — Parse uploaded Excel files and produce row/column-level error reports
3. **Database Import** — Load validated data via insert/upsert with transaction safety
4. **Export** — Query results → Excel with proper formatting

### Relationship with excel-dbapi

sqlalchemy-excel uses [excel-dbapi](https://github.com/yeongseon/excel-dbapi) (`excel-dbapi>=1.0`) as its **core Excel I/O layer** (full dependency, not optional). The integration works through two adapter modules:

- **`ExcelWorkbookSession`** (`excelio/session.py`) — Dual-channel wrapper around an excel-dbapi connection providing both SQL cursor access (data channel) and direct openpyxl workbook access (format channel). Used by `ExcelTemplate` and `ExcelExporter` for workbook creation.
- **`ExcelDbapiReader`** (`reader/excel_dbapi_reader.py`) — SQL-based Excel reader that replaces direct openpyxl usage. Uses `SELECT * FROM SheetName` via excel-dbapi cursors. Used by `ExcelValidator` and `ExcelImporter` for data reading.

## Tech Stack

| Layer | Technology | Role |
|-------|-----------|------|
| ORM | SQLAlchemy 2.0+ | Schema source of truth, DB operations |
| Excel I/O | excel-dbapi ≥ 1.0 | SQL-based Excel read/write, workbook access |
| Excel Format | openpyxl ≥ 3.1 | Cell-level formatting via `ExcelWorkbookSession.workbook` |
| Validation | Pydantic v2 | Row-level type coercion and error messages |
| Security | defusedxml ≥ 0.7 | Safe XML processing for untrusted Excel files |
| CLI | Click ≥ 8.0 | Command-line interface |
| Optional validation | Pandera | DataFrame-level cross-column validation (optional extra) |
| Web integration | FastAPI (optional) | Reference upload/import endpoints |
| Testing | pytest, Hypothesis | Unit/integration/property-based tests |
| Build | hatchling | Package build backend |
| Linting | Ruff | Code linting and formatting |
| Type checking | mypy (strict) | Static type analysis |
| CI/CD | GitHub Actions | Matrix testing (Python 3.10–3.13), PyPI publishing |

## Project Structure

```
sqlalchemy-excel/
├── pyproject.toml              # Package config, dependencies, entry points
├── README.md                   # Comprehensive documentation (~1,100 lines)
├── LICENSE
├── AGENTS.md                   # This file
├── PRD.md                      # Product requirements
├── ARCH.md                     # Architecture document
├── TDD.md                      # Technical design document
├── src/
│   └── sqlalchemy_excel/
│       ├── __init__.py          # Public API re-exports (lazy imports)
│       ├── mapping.py           # ORM model → ExcelMapping schema extraction
│       ├── template.py          # ExcelMapping → .xlsx template generation
│       ├── export.py            # Query result → Excel export
│       ├── exceptions.py        # Full exception hierarchy
│       ├── cli.py               # Click CLI entry point (5 commands)
│       ├── _types.py            # Internal type aliases (FilePath, FileSource, RowDict)
│       ├── _compat.py           # ensure_defusedxml, import_optional, sanitize_cell_value
│       ├── excelio/
│       │   ├── __init__.py      # Re-exports ExcelWorkbookSession
│       │   └── session.py       # ExcelWorkbookSession dual-channel wrapper
│       ├── reader/
│       │   ├── __init__.py
│       │   ├── base.py          # ReaderResult, BaseReader protocol, normalize_header()
│       │   ├── openpyxl_reader.py  # OpenpyxlReader (legacy, retained for compatibility)
│       │   └── excel_dbapi_reader.py  # ExcelDbapiReader (primary, SQL-based reader)
│       ├── validation/
│       │   ├── __init__.py
│       │   ├── engine.py        # ExcelValidator orchestrator (uses ExcelDbapiReader)
│       │   ├── pydantic_backend.py  # PydanticBackend dynamic model generation
│       │   └── report.py        # CellError, ValidationReport
│       ├── load/
│       │   ├── __init__.py
│       │   ├── importer.py      # ExcelImporter (uses ExcelDbapiReader)
│       │   └── strategies.py    # InsertStrategy, UpsertStrategy, DryRunStrategy, ImportResult
│       └── integrations/
│           ├── __init__.py
│           └── fastapi.py       # create_import_router() factory
├── tests/
│   ├── conftest.py              # Shared fixtures (in-memory SQLite, sample models)
│   ├── unit/
│   │   ├── test_mapping.py
│   │   ├── test_template.py
│   │   ├── test_reader.py
│   │   ├── test_validation.py
│   │   ├── test_importer.py
│   │   ├── test_report.py
│   │   └── test_export.py
│   └── integration/
│       └── test_end_to_end.py   # Template → fill → validate → import
├── examples/
│   └── fastapi_upload/
│       ├── app.py
│       └── models.py
└── .github/
    └── workflows/
        ├── ci.yml
        └── publish.yml
```

## Key Design Decisions

1. **src layout** — `src/sqlalchemy_excel/` prevents accidental import of uninstalled package
2. **excel-dbapi as full dependency** — All Excel I/O goes through excel-dbapi for unified SQL-based data access and openpyxl workbook formatting. Not optional — listed in core `dependencies`.
3. **Dual-channel architecture** — `ExcelWorkbookSession` provides both a DB-API cursor (data channel) and an openpyxl workbook (format channel). Template generation uses the workbook for styling while export and reading use SQL.
4. **ExcelDbapiReader replaces direct openpyxl reading** — `ExcelDbapiReader` uses `SELECT * FROM SheetName` via excel-dbapi cursors instead of direct openpyxl cell iteration. This unifies the data access layer.
5. **Pydantic v2 for validation** — Row-level coercion with structured error output (dynamic model generation from ExcelMapping)
6. **Pandera as optional** — DataFrame-level cross-column rules, installed via `pip install sqlalchemy-excel[pandera]`
7. **defusedxml required** — Security: openpyxl doesn't defend against XML attacks by default. Verified at import time.
8. **Lazy imports in `__init__.py`** — `__getattr__`-based lazy loading keeps import time fast and avoids circular imports
9. **hatchling build system** — Modern Python packaging with `src/` layout support
10. **BinaryIO → temp file for excel-dbapi** — `ExcelDbapiReader._resolve_source()` writes BinaryIO to a temp file because excel-dbapi requires file paths. Temp file is cleaned up after read.

## Coding Conventions

- **Type hints**: All public functions fully typed. Use `from __future__ import annotations`.
- **Docstrings**: Google style. All public classes/functions documented.
- **Imports**: `from __future__ import annotations` at top of every module.
- **Testing**: pytest with fixtures. Property-based tests via Hypothesis for round-trip invariants.
- **Linting**: Ruff (lint + format). Config in pyproject.toml.
- **Type checking**: mypy in strict mode (`strict = true` in pyproject.toml).
- **Error handling**: Never bare `except:`. Custom exceptions inherit from `SqlalchemyExcelError`.
- **No `as any` equivalent**: Never use `# type: ignore` without specific error code.
- **Formula injection prevention**: `_compat.sanitize_cell_value()` prefixes dangerous cell values with `'`.

## Exception Hierarchy

```python
class SqlalchemyExcelError(Exception): ...

class MappingError(SqlalchemyExcelError): ...       # ORM introspection failures
class TemplateError(SqlalchemyExcelError): ...      # Template generation failures
class ReaderError(SqlalchemyExcelError): ...        # Excel parsing failures
    class FileFormatError(ReaderError): ...         # Invalid .xlsx file
    class SheetNotFoundError(ReaderError): ...      # Sheet name not found
    class HeaderMismatchError(ReaderError): ...     # Header/column mismatch
class ValidationError(SqlalchemyExcelError): ...    # Data validation failures (wraps report)
class ImportError_(SqlalchemyExcelError): ...       # DB import failures (underscore to avoid builtin clash)
    class DuplicateKeyError(ImportError_): ...      # Unique constraint violation
    class ConstraintViolationError(ImportError_): ...  # Other DB constraint violations
class ExportError(SqlalchemyExcelError): ...        # Export failures
```

## Public API

```python
from sqlalchemy_excel import (
    # Core workflow classes
    ExcelMapping,           # ORM model → mapping config
    ColumnMapping,          # Per-column metadata dataclass
    ExcelTemplate,          # Generate downloadable template
    ExcelValidator,         # Validate uploaded file
    ExcelImporter,          # Import to database
    ExcelExporter,          # Export query results

    # excel-dbapi integration
    ExcelWorkbookSession,   # Dual-channel session (cursor + workbook)

    # Result/report types
    ValidationReport,       # Structured error report
    CellError,              # Single cell validation error
    ImportResult,           # Insert/upsert result counts

    # Exceptions (all re-exported)
    SqlalchemyExcelError,
    MappingError,
    TemplateError,
    ReaderError,
    FileFormatError,
    SheetNotFoundError,
    HeaderMismatchError,
    ValidationError,
    ImportError_,
    DuplicateKeyError,
    ConstraintViolationError,
    ExportError,
)
```

## CLI Commands

```bash
sqlalchemy-excel template --model myapp.models:User --output users.xlsx [--sample-data] [--sheet-name NAME]
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx [--format text|json|excel] [--output report.xlsx]
sqlalchemy-excel import --model myapp.models:User --input upload.xlsx --db sqlite:///app.db [--mode insert|upsert] [--dry-run] [--batch-size 1000]
sqlalchemy-excel export --model myapp.models:User --db sqlite:///app.db --output export.xlsx
sqlalchemy-excel inspect --input mystery.xlsx
```

## Testing Strategy

- **Unit tests**: Each module independently. 7 unit test files covering mapping, template, reader, validation, importer, report, and export.
- **Integration tests**: Full pipeline (template → fill → validate → import → verify DB state).
- **Property-based**: Hypothesis for round-trip (generate template → fill with valid data → validate → must pass).
- **CI matrix**: Python 3.10–3.13 via GitHub Actions.
- **Test count**: 117 tests, all passing.

## Environment Setup

```bash
# Development
pip install -e ".[dev]"

# With all optional extras
pip install -e ".[all]"

# Run tests
pytest

# Lint + format
ruff check . && ruff format .

# Type check
mypy --strict src/
```

## Important Notes for AI Agents

1. Package name is `sqlalchemy-excel` (hyphen), import name is `sqlalchemy_excel` (underscore)
2. SQLAlchemy 2.0+ only — use `Mapped[]`, `mapped_column()`, `DeclarativeBase`, NOT legacy `Column()`
3. openpyxl DataValidation can be *applied* to cells but is NOT *enforced* — server validation is always required
4. defusedxml must be installed for security when processing untrusted Excel files
5. `pandas` is optional — core functionality must work without it
6. FastAPI integration is optional — core library has zero web framework dependency
7. All DB operations go through SQLAlchemy Session — never raw SQL
8. ValidationReport must include: row number, column name, error message, original value, expected type
9. **excel-dbapi is a core dependency** — listed in `dependencies`, not optional. All Excel I/O goes through it.
10. **excel-dbapi uses unquoted table names** — `SELECT * FROM Sheet1` not `"Sheet1"`. Quotes cause lookup failure.
11. **ExcelDbapiReader handles BinaryIO by writing to temp file** — excel-dbapi requires file paths, so binary streams are persisted to temp files and cleaned up after reading.
12. **ExcelWorkbookSession.workbook** returns the openpyxl Workbook via `conn.workbook` property (openpyxl engine only in excel-dbapi).
