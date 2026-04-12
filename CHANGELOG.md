# Changelog

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
