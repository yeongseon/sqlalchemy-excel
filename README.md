<p align="left">
  <img src="https://raw.githubusercontent.com/yeongseon/sqlalchemy-excel/main/logo.svg" alt="sqlalchemy-excel" width="48" height="48" align="middle" />
  <strong style="font-size: 2em;">sqlalchemy-excel</strong>
</p>

![CI](https://github.com/yeongseon/sqlalchemy-excel/actions/workflows/ci.yml/badge.svg)
[![codecov](https://codecov.io/gh/yeongseon/sqlalchemy-excel/branch/main/graph/badge.svg)](https://codecov.io/gh/yeongseon/sqlalchemy-excel)
[![PyPI](https://img.shields.io/pypi/v/sqlalchemy-excel.svg)](https://pypi.org/project/sqlalchemy-excel/)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Docs](https://img.shields.io/badge/docs-GitHub-blue.svg)](https://github.com/yeongseon/sqlalchemy-excel/tree/main/docs)

SQLAlchemy dialect for Excel files — use Excel worksheets as database tables.
This dialect supports CRUD, ORM mapping, aggregations, and constrained JOINs.

## Limitations (Read First)

Before writing any code, understand the dialect's capabilities and limits:

| Feature | Supported? |
|---------|-----------|
| SELECT with WHERE, ORDER BY, LIMIT, OFFSET | ✅ |
| INSERT, UPDATE, DELETE | ✅ |
| Multi-row INSERT VALUES | ✅ |
| INSERT ... SELECT (`from_select`) | ✅ |
| CREATE TABLE / DROP TABLE | ✅ |
| ORM with DeclarativeBase | ✅ |
| Schema inspection (tables, columns) | ✅ |
| IN, BETWEEN, LIKE operators | ✅ |
| DISTINCT | ✅ |
| GROUP BY / HAVING | ✅ |
| Aggregate functions (COUNT, SUM, AVG, MIN, MAX) | ✅ |
| Subqueries in WHERE ... IN | ✅ (non-correlated only) |
| INNER / LEFT JOIN (single, equality ON) | ✅ (constrained) |
| **Chained JOINs** (3+ tables) | ❌ |
| **FULL OUTER JOIN** | ❌ |
| **CTEs / UNION / INTERSECT / EXCEPT** | ❌ |
| **Window functions (OVER)** | ❌ |
| **ALTER TABLE** | ❌ |
| **Foreign keys / indexes** | ❌ |
| **Concurrent writes** | ❌ |
| **Session.rollback()** | No-op (data persists) |

If you need any of the ❌ features, use SQLite, PostgreSQL, or another full-featured database.

---

## Installation

```bash
pip install sqlalchemy-excel
```

`excel-dbapi` is automatically installed as a dependency.

## Quick Start

```python
from sqlalchemy import create_engine, Column, Integer, String
from sqlalchemy.orm import DeclarativeBase, Session, Mapped, mapped_column

engine = create_engine("excel:///data.xlsx")

class Base(DeclarativeBase):
    pass

class User(Base):
    __tablename__ = "Sheet1"
    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column()

Base.metadata.create_all(engine)

with Session(engine) as session:
    session.add(User(id=1, name="Alice"))
    session.commit()

with Session(engine) as session:
    users = session.query(User).all()
```

## URL Format

```python
# Relative path
engine = create_engine("excel:///data.xlsx")

# Absolute path (note four slashes)
engine = create_engine("excel:////home/user/data.xlsx")

# With engine options
engine = create_engine("excel:///data.xlsx", connect_args={"engine": "openpyxl"})
```

## Type Mapping

| SQLAlchemy Type | Excel Storage | Notes |
|---|---|---|
| `String`, `Text`, `VARCHAR`, `CHAR` | TEXT | All string types map to TEXT |
| `Integer`, `SmallInteger`, `BigInteger` | INTEGER | All integer types map to INTEGER |
| `Float`, `Numeric`, `Decimal` | FLOAT | All numeric types map to FLOAT |
| `Boolean` | BOOLEAN | |
| `Date` | DATE | |
| `DateTime`, `TIMESTAMP` | DATETIME | |
| `Time` | TEXT | Stored as text |
| `Uuid` | TEXT | Stored as text |

> BLOB, BINARY, JSON, and ARRAY types are not supported and will raise `CompileError`.

## ORM Examples

### Define a Model

```python
from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, Session, Mapped, mapped_column

engine = create_engine("excel:///data.xlsx")

class Base(DeclarativeBase):
    pass

class User(Base):
    __tablename__ = "users"
    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column()
    age: Mapped[int] = mapped_column()

Base.metadata.create_all(engine)
```

### Insert

```python
with Session(engine) as session:
    session.add(User(id=1, name="Alice", age=30))
    session.add(User(id=2, name="Bob", age=25))
    session.commit()
```

### Query with Filters

```python
from sqlalchemy import select

with Session(engine) as session:
    # Basic query
    users = session.query(User).all()

    # WHERE clause
    user = session.query(User).filter(User.name == "Alice").first()

    # IN operator
    stmt = select(User).where(User.name.in_(["Alice", "Bob"]))
    users = session.scalars(stmt).all()

    # BETWEEN operator
    stmt = select(User).where(User.age.between(25, 35))
    users = session.scalars(stmt).all()

    # LIKE operator
    stmt = select(User).where(User.name.like("A%"))
    users = session.scalars(stmt).all()

    # ORDER BY + LIMIT
    stmt = select(User).order_by(User.age.desc()).limit(5)
    users = session.scalars(stmt).all()
```

### Update and Delete

```python
with Session(engine) as session:
    user = session.query(User).filter(User.id == 1).first()
    if user:
        user.name = "Ann"
        session.commit()

with Session(engine) as session:
    user = session.query(User).filter(User.id == 2).first()
    if user:
        session.delete(user)
        session.commit()
```

## Core Usage

```python
from sqlalchemy import create_engine, text

engine = create_engine("excel:///data.xlsx")

with engine.connect() as conn:
    result = conn.execute(text("SELECT * FROM Sheet1"))
    for row in result:
        print(row)
```

```python
from sqlalchemy import Column, Integer, MetaData, String, Table, insert, select

metadata = MetaData()
source = Table(
    "source",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("name", String),
)
target = Table(
    "target",
    metadata,
    Column("id", Integer, primary_key=True),
    Column("name", String),
)
metadata.create_all(engine)

# Multi-row insert
with engine.connect() as conn:
    conn.execute(
        insert(source),
        [
            {"id": 1, "name": "Alice"},
            {"id": 2, "name": "Bob"},
        ],
    )
    conn.commit()

# INSERT ... SELECT
with engine.connect() as conn:
    conn.execute(
        target.insert().from_select(
            ["id", "name"],
            select(source.c.id, source.c.name),
        )
    )
    conn.commit()
```

## Schema Inspection

```python
from sqlalchemy import create_engine, inspect

engine = create_engine("excel:///data.xlsx")
inspector = inspect(engine)

# List all sheets (tables)
print(inspector.get_table_names())

# Get column info
print(inspector.get_columns("Sheet1"))

# Check if a sheet exists
print(inspector.has_table("Sheet1"))
```

---

## Experimental: Remote Excel via Microsoft Graph API

> **Status**: Experimental — API may change in future releases.

Access Excel files on OneDrive/SharePoint directly:

```bash
pip install sqlalchemy-excel[graph]
```

```python
from sqlalchemy import create_engine
from azure.identity import DefaultAzureCredential

engine = create_engine(
    "excel+graph:///drive_id/item_id",
    connect_args={"credential": DefaultAzureCredential()},
)

with engine.connect() as conn:
    result = conn.execute(text("SELECT * FROM Sheet1"))
    for row in result:
        print(row)
```

URL format: `excel+graph:///drive_id/item_id` where `drive_id` and `item_id` are Microsoft Graph resource identifiers.
Query parameters: `?readonly=false` to enable write operations.

---

## Related Projects

- [excel-dbapi](https://github.com/yeongseon/excel-dbapi) — The underlying PEP 249 DB-API 2.0 driver for Excel files.

## License

MIT
