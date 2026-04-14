# Project Roadmap

> **Current version**: 0.5.4 · **Python**: 3.10+ · **Published**: [PyPI](https://pypi.org/project/sqlalchemy-excel/)

## Completed

### Core SQL and ORM surface

- SQLAlchemy 2.x dialect implementation for local Excel files
- ORM support with `DeclarativeBase`, `Session`, and mapper reflection paths used in tests
- CRUD support: `SELECT`, `INSERT`, `UPDATE`, `DELETE`
- DDL support: `CREATE TABLE`, `DROP TABLE`
- Multi-row `INSERT ... VALUES` and `INSERT ... SELECT`
- Inspector support: `get_table_names`, `get_columns`, `has_table`

### Query capabilities available today

- Filtering/operators: `IN`, `BETWEEN`, `LIKE`
- Aggregation: `COUNT`, `SUM`, `AVG`, `MIN`, `MAX`
- `GROUP BY` and `HAVING`
- `DISTINCT` (single-table queries)
- Non-correlated subqueries in `WHERE ... IN` (for `SELECT`, `UPDATE`, `DELETE`; not supported with JOINs)
- Join surface: `INNER`, `LEFT`, `RIGHT` shape, `FULL OUTER`, `CROSS`, chained joins
- Compound queries: `UNION`, `UNION ALL`, `INTERSECT`, `EXCEPT`
- `CASE WHEN` expressions and arithmetic expressions
- UPSERT via `ON CONFLICT DO NOTHING / DO UPDATE`

### Platform and quality

- `ExcelGraphDialect` for Microsoft Graph-backed workbooks (`excel+graph:///drive_id/item_id`)
- Strict mypy, ruff linting/formatting, and CI on supported Python versions
- High test coverage across compiler, dialect, ORM, DML/DDL, reflection, and end-to-end flows

## In Progress / Next

- Expand documentation examples for advanced query shapes and ORM patterns
- Continue hardening edge cases around SQLAlchemy compilation and parser interoperability
- Improve release notes and compatibility guidance as excel-dbapi evolves

## Implemented Schema Migration Surface

- `ALTER TABLE` supports `ADD COLUMN`, `DROP COLUMN`, and `RENAME COLUMN`.

## Not Planned

- Full ACID transactions
- Concurrent multi-writer semantics
- Foreign key enforcement and index management
- Stored procedures and triggers

These constraints are inherited from the Excel file model and underlying driver behavior.

## Long-Term Vision

1. Keep SQLAlchemy usage intuitive for spreadsheet-backed data workflows.
2. Maintain compatibility with excel-dbapi capabilities as they expand.
3. Provide clear boundaries so teams can choose Excel vs. traditional databases intentionally.

See [CHANGELOG.md](../CHANGELOG.md) for release-by-release details.
