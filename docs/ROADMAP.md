# Project Roadmap

> **Current version**: 0.4.0 · **Python**: 3.10+ · **Published**: [PyPI](https://pypi.org/project/sqlalchemy-excel/)

## Completed

### v0.1.0 — Initial Release

- SQLAlchemy 2.0 dialect implementation
- ORM support with `DeclarativeBase`
- Basic SQL: SELECT, INSERT, UPDATE, DELETE, CREATE TABLE, DROP TABLE
- WHERE clauses with comparison operators
- ORDER BY and LIMIT
- Type mapping for common SQLAlchemy types
- Schema inspection (`get_table_names`, `get_columns`, `has_table`)

### v0.2.x — Dialect Rewrite & Quality

- Complete dialect architecture rewrite: `ExcelCompiler`, `ExcelDDLCompiler`, `ExcelTypeCompiler`
- Enhanced operators: IN, BETWEEN, LIKE
- Codecov integration for coverage tracking
- mypy strict mode with full type safety
- PyPI Trusted Publisher (OIDC) for secure releases
- GitHub Actions CI/CD pipeline

### v0.3.x — Graph API & Remote Excel

- `ExcelGraphDialect` for remote Excel files via Microsoft Graph API
- `excel+graph:///drive_id/item_id` URL scheme
- Entry point `excel.graph` for SQLAlchemy dialect resolution
- Optional dependency: `pip install sqlalchemy-excel[graph]`
- URL percent-decoding for drive/item IDs with special characters
- `readonly` query parameter forwarding to Graph backend
- `supports_statement_cache = False` on both dialects

### v0.4.0 — Stabilization (Current)

- HAVING guard via `_compose_select_body` override (compile-time error instead of silent ignore)
- End-to-end tests: CRUD round-trip, ORM Session, inspector reflection, DDL lifecycle, rollback no-op
- Compiler guard tests for all unsupported SQL features
- Type compiler full coverage tests
- Reflection edge case tests
- Test coverage: **98% (117 tests)**
- README restructured: limitations-first layout, Graph API moved to experimental section

## Future

### Planned

- **DISTINCT**: Remove duplicate rows
- **OFFSET**: Pagination support (currently only LIMIT works)
- **Aggregate functions**: COUNT, SUM, AVG, MIN, MAX
- **GROUP BY**: Grouping with aggregate functions
- **Subqueries**: Nested SELECT statements
- **Multi-sheet JOIN**: INNER JOIN, LEFT JOIN across sheets
- **Async dialect**: `excel+aio://` with AsyncEngine support

> These features require changes in the underlying [excel-dbapi](https://github.com/yeongseon/excel-dbapi) query engine.

### Not Planned

These are explicitly out of scope:

- Full ACID transactions (Excel files don't support them)
- Concurrent write support (single-writer model by design)
- ALTER TABLE / schema migration
- Foreign key enforcement or index support
- Stored procedures or triggers

## Long-Term Vision

1. **Seamless SQLAlchemy experience** for Excel files
2. **Analysts and developers** use familiar SQL/ORM tools with Excel
3. **Bridge** ad-hoc Excel data and structured database workflows
4. **Cloud-native Excel** via Microsoft Graph API (experimental)

**Not a goal**: Replace real databases for production workloads.

## Versioning

sqlalchemy-excel follows [Semantic Versioning](https://semver.org/):

- **PATCH** (0.x.**y**): Bug fixes
- **MINOR** (0.**x**.0): New features, backward-compatible
- **MAJOR** (**x**.0.0): Breaking changes, stable API

**Current status**: Beta (0.x.x) — API may change before 1.0.0.

---

See [CHANGELOG.md](../CHANGELOG.md) for detailed release history.
