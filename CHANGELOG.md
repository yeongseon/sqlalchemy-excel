# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.5.4] - 2026-04-13

### Added
- ALTER TABLE support (`ADD COLUMN`, `DROP COLUMN`, `RENAME COLUMN`) with metadata synchronization.
- UPSERT support via `INSERT ... ON CONFLICT DO NOTHING/DO UPDATE` with conflict target validation.
- Graph dialect support (`excel+graph:///drive_id/item_id`) for Microsoft Graph-backed workbooks.
- Compiler guards and expanded SQL support coverage across JOIN, compound queries, subqueries, aggregates, and CASE expressions.
- CI matrix coverage for Python 3.10-3.13 with strict type checking and high test coverage.

### Changed
- Compiler behavior now normalizes SQL execution text consistently before dispatch (`do_execute` / `do_execute_no_params`).
- Dependency baseline raised to `excel-dbapi>=0.4.1,<1.0` (and `excel-dbapi[graph]>=0.4.1,<1.0` for Graph extras).
- Documentation and development workflows were aligned for cross-repo delivery (dialect + driver) and repeatable releases.

### Fixed
- Graph optional dependency handling in tests and CI (graceful skip when extras are unavailable).
- JOIN/tree validation edge cases and aggregate/alias compilation edge behavior.
- Strict typing compatibility for SQLAlchemy override signatures and dialect internals.

[Unreleased]: https://github.com/yeongseon/sqlalchemy-excel/compare/v0.5.4...HEAD
[0.5.4]: https://github.com/yeongseon/sqlalchemy-excel/releases/tag/v0.5.4
