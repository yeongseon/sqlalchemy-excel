# Compatibility and Migration Guide

This document explains how `sqlalchemy-excel` versions pair with `excel-dbapi`,
which features are owned by the dialect vs. the DB-API driver, and how to
migrate safely between releases.

## Version Compatibility Matrix

| sqlalchemy-excel | excel-dbapi | SQLAlchemy | Python | Notes |
|---|---|---|---|---|
| `0.5.4` (current) | `>=0.4.1,<1.0` | `>=2.0` | `>=3.10` | Current release line; includes CASE, set ops, JOIN improvements |
| `0.5.0` to `0.5.3` | `>=0.4.0,<1.0` | `>=2.0` | `>=3.10` | JOINs, set operations, GROUP BY/HAVING + JOIN support |
| `0.3.0` to `0.4.x` | `>=0.2.0` | `>=2.0` | `>=3.10` | Graph dialect introduced in `0.3.0`; docs/test hardening in `0.4.x` |
| `0.2.0` to `0.2.2` | `>=0.2.0` | `>=2.0` | `>=3.10` | Full dialect rewrite and broader SQL coverage |
| `0.1.0` | Initial release line | `>=2.0` | `>=3.10` | First SQLAlchemy Excel dialect release |

Current baseline:

- `sqlalchemy-excel` version: `0.5.4`
- `excel-dbapi` requirement: `>=0.4.1,<1.0`
- SQLAlchemy requirement: `>=2.0`
- Python requirement: `>=3.10`

## Feature Parity by Ownership

Many features need coordination between this dialect (SQL compilation/
adaptation) and `excel-dbapi` (execution engine and workbook semantics).

| Feature Area | Requires Dialect + Driver | Primarily Driver-Side | Primarily Dialect-Side | Notes |
|---|---|---|---|---|
| Basic CRUD (`SELECT/INSERT/UPDATE/DELETE`) | Yes | No | No | End-to-end path across compiler and execution |
| DDL (`CREATE TABLE`, `DROP TABLE`) | Yes | No | No | Compiler emits SQL; driver applies workbook changes |
| Inspector/reflection (`get_table_names`, columns) | Yes | Yes | Partial | Driver exposes metadata; dialect maps into inspector API |
| Type mapping (`String`, `Integer`, `DateTime`, etc.) | Yes | Yes | Yes | Dialect compiles SQL types; driver stores/retrieves values |
| JOIN support (constrained equality ON clauses) | Yes | No | Yes | Dialect validates/compiles shape; driver executes SQL |
| Aggregates + `GROUP BY`/`HAVING` | Yes | No | Yes | Dialect enables and guards unsupported forms |
| Set operations (`UNION`, `INTERSECT`, `EXCEPT`) | Yes | No | Yes | Requires compiler support plus parser/execution support |
| Non-correlated subqueries (`WHERE ... IN (SELECT ...)`) | Yes | No | Yes | Correlated subqueries remain unsupported |
| Transaction controls (`commit`/`rollback`) | Yes | Yes | Partial | Driver semantics dominate behavior in practice |
| Graph dialect (`excel+graph`) | Yes | Yes | Yes | Dialect URL and options + remote backend behavior |

## Transaction and Rollback Semantics

`sqlalchemy-excel` does not provide full ACID transactional guarantees.

- `Session.commit()` persists changes to the workbook as expected.
- `Session.rollback()` is effectively a no-op for local workbook writes in
  standard usage; treat writes as durable once executed.
- Use a single-writer model. Concurrent writes to the same workbook are not
  supported and can lead to conflicts/corruption.
- If you need strict transaction isolation, rollback guarantees, and
  concurrent writers, use SQLite/PostgreSQL instead of Excel-backed storage.

Practical guidance:

- Stage risky transformations in a copy of the workbook.
- Use explicit backup/versioned workbook files before bulk updates.
- Keep write units small and validated.

## Migration Guide

### Before Any Upgrade

1. Pin current versions in your lock file/requirements.
2. Run your full test suite against the current baseline.
3. Back up workbook data used in integration tests and production flows.

### Upgrading to `0.5.4` (from `0.5.0` to `0.5.3`)

- Update dependency to `sqlalchemy-excel==0.5.4`.
- Ensure `excel-dbapi` resolves to `>=0.4.1,<1.0`.
- Validate CASE expression flows (`SELECT` and `UPDATE`) if your app uses
  conditional SQL logic.

### Upgrading from `0.4.x` to `0.5.x`

- Ensure `excel-dbapi>=0.4.1,<1.0` (minimum requirement increased in 0.5 line).
- Re-run query-heavy tests that involve JOINs, aggregates, and set operations.
- If you previously avoided JOIN-related aggregation due to earlier limits,
  verify behavior against your expected SQLAlchemy-generated SQL.

### Upgrading from `0.2.x` to `0.3.x`/`0.4.x`

- If adopting Graph support, install extras: `pip install sqlalchemy-excel[graph]`.
- For local Excel only, no Graph extras are required.
- Revalidate docs/usage assumptions around limitations and rollback behavior.

### Upgrading from `0.1.x`

- Treat as a major behavior jump: review CHANGELOG entries from `0.2.0` onward.
- Re-test ORM and Core paths, especially type handling and SQL compilation.

### Recommended Upgrade Procedure

1. Upgrade in a branch with pinned versions.
2. Run static checks (`ruff`, `mypy --strict`) and full tests.
3. Run a workbook round-trip smoke test (insert/update/delete/select).
4. Roll forward to production only after workbook backups are in place.

## Known Limitations

- No CTE support.
- No window functions (`OVER`).
- No `ALTER TABLE` migration primitives.
- No foreign key/index enforcement.
- No correlated subqueries.
- Constrained JOIN `ON` clauses (equality joins across sources).
- No concurrent multi-writer guarantees.
- `Session.rollback()` does not provide traditional database rollback semantics.

When these limitations are blockers, prefer a traditional database backend.
