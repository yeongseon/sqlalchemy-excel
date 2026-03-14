# AGENTS.md — sqlalchemy-excel

> Project knowledge base for AI agents. Last updated: 2026-03-14.

## Project Identity

- **Name**: sqlalchemy-excel
- **Package**: `sqlalchemy-excel` (PyPI), import as `sqlalchemy_excel`
- **License**: MIT (Copyright 2025 Yeongseon Choe)
- **Repository**: https://github.com/yeongseon/sqlalchemy-excel
- **Python**: 3.10–3.13
- **Stage**: Pre-MVP (greenfield)

## One-Line Description

SQLAlchemy model-driven Excel template generation, server-side validation, and database import toolkit.

## What This Project Does

sqlalchemy-excel treats SQLAlchemy ORM models as the **single source of truth** for:

1. **Template Generation** — Create Excel files with correct headers, types, dropdowns, sample data
2. **Server Validation** — Parse uploaded Excel files and produce row/column-level error reports
3. **Database Import** — Load validated data via insert/upsert with transaction safety
4. **Export** — Query results → Excel with proper formatting

It is NOT (initially) a SQLAlchemy dialect. The "query Excel like a DB" feature is a long-term stretch goal.

## Tech Stack

| Layer | Technology | Role |
|-------|-----------|------|
| ORM | SQLAlchemy 2.0+ | Schema source of truth, DB operations |
| Validation | Pydantic v2 | Row-level type coercion and error messages |
| Excel I/O | openpyxl | Read/write xlsx, template formatting, data validation |
| Optional validation | Pandera | DataFrame-level cross-column validation (optional extra) |
| CLI | Click | Command-line interface |
| Web integration | FastAPI (optional) | Reference upload/import endpoints |
| Testing | pytest, Hypothesis | Unit/integration/property-based tests |
| CI/CD | GitHub Actions | Matrix testing, PyPI publishing |

## Project Structure

```
sqlalchemy-excel/
├── pyproject.toml              # Package config, dependencies, entry points
├── README.md
├── LICENSE
├── AGENTS.md                   # This file
├── PRD.md                      # Product requirements
├── ARCH.md                     # Architecture document
├── TDD.md                      # Technical design document
├── src/
│   └── sqlalchemy_excel/
│       ├── __init__.py          # Public API re-exports
│       ├── mapping.py           # ORM model → ExcelMapping schema extraction
│       ├── template.py          # ExcelMapping → .xlsx template generation
│       ├── reader/
│       │   ├── __init__.py
│       │   ├── base.py          # Abstract reader interface
│       │   ├── openpyxl_reader.py  # openpyxl-based reader (default)
│       │   └── pandas_reader.py    # pandas-based reader (optional)
│       ├── validation/
│       │   ├── __init__.py
│       │   ├── engine.py        # Validation orchestrator
│       │   ├── pydantic_backend.py  # Pydantic v2 validation backend
│       │   ├── pandera_backend.py   # Pandera validation backend (optional)
│       │   └── report.py        # ValidationReport with row/col errors
│       ├── load/
│       │   ├── __init__.py
│       │   ├── importer.py      # ExcelImporter (insert/upsert/dry-run)
│       │   └── strategies.py    # Insert, upsert, replace strategies
│       ├── export.py            # Query result → Excel export
│       ├── integrations/
│       │   ├── __init__.py
│       │   └── fastapi.py       # FastAPI router factory, upload helpers
│       ├── cli.py               # Click CLI entry point
│       ├── _types.py            # Internal type aliases
│       └── _compat.py           # Version compatibility helpers
├── tests/
│   ├── conftest.py              # Shared fixtures (in-memory SQLite, sample models)
│   ├── unit/
│   │   ├── test_mapping.py
│   │   ├── test_template.py
│   │   ├── test_validation.py
│   │   ├── test_importer.py
│   │   └── test_report.py
│   ├── integration/
│   │   ├── test_end_to_end.py   # Template → fill → validate → import
│   │   └── test_fastapi.py
│   └── fixtures/
│       ├── sample_valid.xlsx
│       ├── sample_invalid.xlsx
│       └── sample_large.xlsx
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
2. **openpyxl as core, pandas as optional** — Minimizes required dependencies; pandas users opt in
3. **Pydantic v2 for validation** — Row-level coercion with structured error output (ValidationError)
4. **Pandera as optional** — DataFrame-level cross-column rules, installed via `pip install sqlalchemy-excel[pandera]`
5. **No dialect in MVP** — SQLAlchemy dialect (create_engine("excel://...")) is a future feature, not MVP scope
6. **defusedxml required** — Security: openpyxl doesn't defend against XML attacks by default

## Coding Conventions

- **Type hints**: All public functions fully typed. Use `from __future__ import annotations`.
- **Docstrings**: Google style. All public classes/functions documented.
- **Imports**: `from __future__ import annotations` at top of every module.
- **Testing**: pytest with fixtures. Property-based tests via Hypothesis for round-trip invariants.
- **Linting**: Ruff (lint + format). Config in pyproject.toml.
- **Type checking**: mypy in strict mode.
- **Error handling**: Never bare `except:`. Custom exceptions inherit from `SqlalchemyExcelError`.
- **No `as any` equivalent**: Never use `# type: ignore` without specific error code.

## Exception Hierarchy

```python
class SqlalchemyExcelError(Exception): ...
class MappingError(SqlalchemyExcelError): ...       # ORM introspection failures
class TemplateError(SqlalchemyExcelError): ...      # Template generation failures
class ReaderError(SqlalchemyExcelError): ...        # Excel parsing failures
class ValidationError(SqlalchemyExcelError): ...    # Data validation failures (wraps report)
class ImportError_(SqlalchemyExcelError): ...       # DB import failures (underscore to avoid builtin clash)
class ExportError(SqlalchemyExcelError): ...        # Export failures
```

## Public API (MVP)

```python
from sqlalchemy_excel import (
    ExcelMapping,      # ORM model → mapping config
    ExcelTemplate,     # Generate downloadable template
    ExcelValidator,    # Validate uploaded file
    ExcelImporter,     # Import to database
    ExcelExporter,     # Export query results
    ValidationReport,  # Structured error report
)
```

## CLI Commands

```bash
sqlalchemy-excel template --model myapp.models:User --output users.xlsx
sqlalchemy-excel validate --model myapp.models:User --input upload.xlsx
sqlalchemy-excel import --model myapp.models:User --input upload.xlsx --db sqlite:///app.db
sqlalchemy-excel export --model myapp.models:User --db sqlite:///app.db --output export.xlsx
sqlalchemy-excel inspect --input mystery.xlsx
```

## Testing Strategy

- **Unit tests**: Each module independently. Mock openpyxl/SQLAlchemy where needed.
- **Integration tests**: Full pipeline (template → fill → validate → import → verify DB state).
- **Property-based**: Hypothesis for round-trip (generate template → fill with valid data → validate → must pass).
- **Fixture files**: Real .xlsx files in tests/fixtures/ for edge cases.
- **Security tests**: Malformed XML, oversized files, formula injection.
- **CI matrix**: Python 3.10–3.13, SQLAlchemy 2.0.x.

## Environment Setup

```bash
# Development
pip install -e ".[dev]"

# With all optional extras
pip install -e ".[dev,pandas,pandera,fastapi]"

# Run tests
pytest

# Lint + format
ruff check . && ruff format .

# Type check
mypy src/
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
