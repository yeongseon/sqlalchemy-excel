# Changelog

## [0.5.1] - 2026-04-13

### Added
- Compiler support for chained (multi-table) JOIN trees with recursive validation
- Compiler and end-to-end tests for chained INNER/LEFT JOIN compilation and execution
- Compiler and end-to-end tests for RIGHT JOIN shape compilation (`join(..., isouter=True)` with swapped sides)

### Changed
- `_validate_join_tree()` now validates JOIN subtrees recursively instead of rejecting chained JOINs
- JOIN ON-clause cross-source validation now resolves tables across each JOIN side, including nested JOIN nodes
- GROUP BY + JOIN guard now covers chained JOIN trees through JOIN-subtree validation flow

## [0.5.0] - 2026-04-13

### Added
- JOIN support: INNER JOIN and LEFT JOIN compile to valid SQL for excel-dbapi
- Compiler auto-detects JOIN context and emits table-qualified column names
- E2E tests: inner join, left join with NULL fill
- Compiler guard tests for JOIN compilation
- Direct `visit_subquery` guard tests for improved coverage
- `supports_multivalues_insert = True` dialect flag for multi-row INSERT support
- `visit_insert()` compiler override for `from_select()` (INSERT...SELECT) compilation

### Changed
- `visit_join()`: removed CompileError guard, delegates to base SQLAlchemy compiler
- `visit_column()`: conditionally includes table prefix based on JOIN context
- `_setup_select_stack()`: detects JOIN in FROM clause before column compilation
- `visit_function()`: regex updated to allow qualified column names (e.g., `a.id`)
- excel-dbapi dependency updated to >=0.4.0
- Test count: 152 → 155, coverage: 94% → 98%
- Version bumped to 0.5.0

## [0.4.0] - 2026-04-12

### Added
- HAVING guard via `_compose_select_body` override (previously `having_clause` was never called by SQLAlchemy)
- End-to-end tests: CRUD round-trip, ORM Session, inspector reflection, DDL lifecycle, rollback no-op
- Compiler guard tests for all unsupported SQL features
- Type compiler full coverage tests
- Reflection edge case tests

### Changed
- README restructured: limitations-first layout, Graph API moved to experimental section
- Test coverage: 80% → 98% (117 tests)

## [0.3.2] - 2026-04-12

### Fixed
- Install `graph` extras in CI to resolve `httpx` ModuleNotFoundError
- Use `pytest.importorskip("httpx")` in Graph dialect tests for graceful skip when extras not installed

## [0.3.1] - 2026-04-12

### Fixed
- Add explicit `supports_statement_cache = False` to `ExcelGraphDialect` to suppress SQLAlchemy caching warning
- Fix import ordering in test_graph_dialect.py for ruff I001 compliance

## [0.3.0] - 2026-04-12

### Added
- `ExcelGraphDialect` for remote Excel files via Microsoft Graph API
- `excel+graph:///drive_id/item_id` URL scheme support
- Entry point `excel.graph` for SQLAlchemy dialect resolution
- Optional dependency: `pip install sqlalchemy-excel[graph]`
- URL percent-decoding for drive/item IDs with special characters
- `readonly` query parameter forwarding to Graph backend
- Comprehensive Graph dialect tests with `httpx.MockTransport`
- `docs/` directory with USAGE.md, DEVELOPMENT.md, and ROADMAP.md

### Changed
- Version bumped to 0.3.0

## [0.2.2] - 2026-04-12

### Fixed
- Restored cast() calls needed for CI mypy and suppress redundant-cast locally
- Version bumped to 0.2.2

## [0.2.1] - 2026-04-12

### Added
- Project logo (modern minimalist SVG)
- Contributing guide, Code of Conduct, Security and Support policies
- Development tooling: Makefile, .editorconfig, pre-commit-config, codecov.yml, git-cliff config
- GitHub issue/PR templates and project management files
- py.typed marker for PEP 561 compliance
- twine check step in publish workflow

### Changed
- Classifier updated from Alpha to Beta
- Changelog URL added to project metadata

### Fixed
- Oracle review findings: rollback docs, absolute logo URLs, metadata alignment

## [0.2.0] - 2026-04-12

### Added
- Full dialect rewrite: ExcelCompiler, ExcelDDLCompiler, ExcelTypeCompiler, ExcelInspectionMixin
- Comprehensive README with ORM examples, type mapping table, schema inspection docs
- Test coverage reporting with Codecov CI integration
- IN, BETWEEN, LIKE operator tests for SQLAlchemy dialect

### Changed
- excel-dbapi dependency updated to >=0.2.0
- Version bumped to 0.2.0

### Fixed
- mypy strict incompatibility with SQLAlchemy dialect overrides (temporarily disabled then re-enabled)

## [0.1.0] - 2026-04-12

- Initial release
- SQLAlchemy 2.0 dialect for Excel files
- PEP 249 DB-API 2.0 driver via excel-dbapi
- SQL support: SELECT, INSERT, UPDATE, DELETE, CREATE TABLE, DROP TABLE
- WHERE clause with AND/OR, comparison operators, IS NULL, IS NOT NULL
- ORDER BY, LIMIT
- Type mapping: TEXT, INTEGER, FLOAT, BOOLEAN, DATE, DATETIME
- Reflection: get_table_names, get_columns, get_pk_constraint, has_table
- ORM support with DeclarativeBase
