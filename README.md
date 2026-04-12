# sqlalchemy-excel

SQLAlchemy dialect for Excel files — use Excel as a database.

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

## Installation

```bash
pip install sqlalchemy-excel
```

## URL Format

```python
# Relative path
engine = create_engine("excel:///data.xlsx")

# Absolute path (note four slashes)
engine = create_engine("excel:////home/user/data.xlsx")
```

## Features

- Full SQLAlchemy 2.0 dialect
- PEP 249 DB-API 2.0 compliant driver ([excel-dbapi](https://github.com/yeongseon/excel-dbapi))
- SELECT with WHERE, ORDER BY, LIMIT
- INSERT, UPDATE, DELETE
- CREATE TABLE / DROP TABLE
- ORM support with `DeclarativeBase`
- Type mapping: String, Integer, Float, Boolean, Date, DateTime

## Limitations

- No JOIN, GROUP BY, HAVING, DISTINCT, OFFSET
- No subqueries, CTEs, or aggregate functions
- No ALTER TABLE, foreign keys, or indexes
- Single-table operations only

## License

MIT
