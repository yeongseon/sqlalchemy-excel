# Changelog

## 0.1.0 (2026-04-12)

- Initial release
- SQLAlchemy 2.0 dialect for Excel files
- PEP 249 DB-API 2.0 driver via excel-dbapi
- SQL support: SELECT, INSERT, UPDATE, DELETE, CREATE TABLE, DROP TABLE
- WHERE clause with AND/OR, comparison operators, IS NULL, IS NOT NULL
- ORDER BY, LIMIT
- Type mapping: TEXT, INTEGER, FLOAT, BOOLEAN, DATE, DATETIME
- Reflection: get_table_names, get_columns, get_pk_constraint, has_table
- ORM support with DeclarativeBase
