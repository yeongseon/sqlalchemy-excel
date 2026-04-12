# Usage Guide

This guide covers how to use sqlalchemy-excel in your projects.

## Installation

```bash
pip install sqlalchemy-excel
```

The underlying driver `excel-dbapi` is automatically installed as a dependency.

## URL Format

sqlalchemy-excel uses the `excel://` URL scheme:

```python
from sqlalchemy import create_engine

# Relative path (relative to current working directory)
engine = create_engine("excel:///data.xlsx")

# Absolute path (note four slashes total: excel:// + //)
engine = create_engine("excel:////home/user/data.xlsx")
engine = create_engine("excel:////Users/alice/Documents/data.xlsx")

# With engine options (passed to connect_args)
engine = create_engine("excel:///data.xlsx", connect_args={"engine": "openpyxl"})
```

**Important**: Absolute paths require **four slashes** total (`excel:////absolute/path.xlsx`).

## Basic ORM Usage

### Define Models

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

# Create the table (sheet) if it doesn't exist
Base.metadata.create_all(engine)
```

### Use Sessions

```python
# Insert data
with Session(engine) as session:
    session.add(User(id=1, name="Alice", age=30))
    session.add(User(id=2, name="Bob", age=25))
    session.commit()

# Query data
with Session(engine) as session:
    users = session.query(User).all()
    for user in users:
        print(f"{user.id}: {user.name} ({user.age})")
```

## Core Usage

You can also use SQLAlchemy Core with `text()` queries:

```python
from sqlalchemy import create_engine, text

engine = create_engine("excel:///data.xlsx")

with engine.connect() as conn:
    # Execute a raw SQL query
    result = conn.execute(text("SELECT * FROM Sheet1"))
    for row in result:
        print(row)
```

> **Note**: excel-dbapi uses **qmark paramstyle** (`?`). When using `text()` queries,
> SQLAlchemy handles parameter translation automatically — you can use either `:name`
> style (SQLAlchemy translates it) or pass parameters directly. However, the underlying
> driver only understands `?` placeholders.

## Query Examples

### WHERE Clause

```python
from sqlalchemy import select

with Session(engine) as session:
    # Simple equality
    user = session.query(User).filter(User.name == "Alice").first()

    # Comparison operators
    stmt = select(User).where(User.age > 25)
    users = session.scalars(stmt).all()
```

### IN Operator

```python
with Session(engine) as session:
    stmt = select(User).where(User.name.in_(["Alice", "Bob", "Charlie"]))
    users = session.scalars(stmt).all()
```

### BETWEEN Operator

```python
with Session(engine) as session:
    # Find users with age between 25 and 35
    stmt = select(User).where(User.age.between(25, 35))
    users = session.scalars(stmt).all()
```

### LIKE Operator

```python
with Session(engine) as session:
    # Find users whose name starts with 'A'
    stmt = select(User).where(User.name.like("A%"))
    users = session.scalars(stmt).all()

    # Contains 'li'
    stmt = select(User).where(User.name.like("%li%"))
    users = session.scalars(stmt).all()
```

### ORDER BY and LIMIT

```python
with Session(engine) as session:
    # Order by age descending
    stmt = select(User).order_by(User.age.desc())
    users = session.scalars(stmt).all()

    # Get top 5 oldest users
    stmt = select(User).order_by(User.age.desc()).limit(5)
    users = session.scalars(stmt).all()
```

## Insert, Update, and Delete

### Insert

```python
with Session(engine) as session:
    new_user = User(id=3, name="Charlie", age=28)
    session.add(new_user)
    session.commit()
```

```python
from sqlalchemy import insert

# Multi-row insert (Core)
with engine.connect() as conn:
    conn.execute(
        insert(User.__table__),
        [
            {"id": 4, "name": "Dora", "age": 29},
            {"id": 5, "name": "Evan", "age": 41},
        ],
    )
    conn.commit()
```

```python
from sqlalchemy import Column, Integer, MetaData, String, Table, select

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

# INSERT ... SELECT (Core)
with engine.connect() as conn:
    conn.execute(
        target.insert().from_select(
            ["id", "name"],
            select(source.c.id, source.c.name).where(source.c.id >= 2),
        )
    )
    conn.commit()
```

### Update

```python
with Session(engine) as session:
    user = session.query(User).filter(User.id == 1).first()
    if user:
        user.name = "Ann"
        user.age = 31
        session.commit()
```

### Delete

```python
with Session(engine) as session:
    user = session.query(User).filter(User.id == 2).first()
    if user:
        session.delete(user)
        session.commit()
```

## Schema Inspection

SQLAlchemy's inspector API works with Excel files:

```python
from sqlalchemy import create_engine, inspect

engine = create_engine("excel:///data.xlsx")
inspector = inspect(engine)

# List all sheets (tables)
table_names = inspector.get_table_names()
print(f"Available sheets: {table_names}")

# Get columns for a specific sheet
columns = inspector.get_columns("users")
for col in columns:
    print(f"{col['name']}: {col['type']}")

# Check if a sheet exists
if inspector.has_table("users"):
    print("The 'users' sheet exists")
```

## Type Mapping

sqlalchemy-excel maps SQLAlchemy types to Excel storage types:

| SQLAlchemy Type | Excel Storage | Notes |
|-----------------|---------------|-------|
| `String`, `Text`, `VARCHAR`, `CHAR` | TEXT | All string types → TEXT |
| `Integer`, `SmallInteger`, `BigInteger` | INTEGER | All integer types → INTEGER |
| `Float`, `Numeric`, `Decimal` | FLOAT | All numeric types → FLOAT |
| `Boolean` | BOOLEAN | Stored as boolean |
| `Date` | DATE | Date without time |
| `DateTime`, `TIMESTAMP` | DATETIME | Date with time |
| `Time` | TEXT | Stored as text |
| `Uuid` | TEXT | Stored as text |

**Unsupported types**: BLOB, BINARY, JSON, ARRAY (will raise `CompileError`)

## Limitations

sqlalchemy-excel has some limitations due to the nature of Excel as a database:

- **Constrained JOIN support**: Only single INNER/LEFT JOIN with equality ON clause (`t1.col = t2.col`). No chained JOINs, FULL OUTER JOIN, OR/non-equality ON clauses, or non-column operands.
- **Non-correlated subqueries only**: Subqueries supported only in `WHERE ... IN (SELECT ...)`. No correlated, nested, or DML subqueries.
- **No CTEs, UNION, INTERSECT, EXCEPT**: Only simple SELECT/INSERT/UPDATE/DELETE.
- **No window functions**: `OVER` clause raises `CompileError`.
- **No ALTER TABLE**: Cannot modify table structure after creation.
- **No foreign keys or indexes**: Excel has no concept of these.
- **No concurrent writes**: Use a single-writer model.
- **Rollback is a no-op**: `Session.rollback()` does nothing — the underlying driver's rollback only works at the DB-API connection level with `autocommit=False`.

## Security

**Always use parameterized queries** to prevent SQL injection:

```python
# ✅ GOOD: ORM queries are automatically parameterized
with Session(engine) as session:
    user = session.query(User).filter(User.name == user_input).first()

# ✅ GOOD: Core text() with bound parameters
with engine.connect() as conn:
    result = conn.execute(
        text("SELECT * FROM users WHERE name = :name"),
        {"name": user_input}
    )

# ❌ BAD: String interpolation (vulnerable to SQL injection)
with engine.connect() as conn:
    result = conn.execute(
        text(f"SELECT * FROM users WHERE name = '{user_input}'")
    )
```
