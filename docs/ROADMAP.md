# Project Roadmap

This roadmap outlines the past achievements and future plans for sqlalchemy-excel.

## Completed Features

### v0.1.0 — Initial Release
- ✅ SQLAlchemy 2.0 dialect implementation
- ✅ ORM support with `DeclarativeBase`
- ✅ Basic SQL operations: SELECT, INSERT, UPDATE, DELETE
- ✅ WHERE clauses with comparison operators
- ✅ ORDER BY and LIMIT support
- ✅ Type mapping for common SQLAlchemy types
- ✅ Integration with excel-dbapi driver
- ✅ Schema inspection (`get_table_names`, `get_columns`, `has_table`)

### v0.2.x — Dialect Rewrite and Quality Improvements
- ✅ Complete dialect architecture rewrite
  - ExcelCompiler for SQL compilation
  - DDLCompiler for CREATE/DROP TABLE
  - TypeCompiler for type system
- ✅ Enhanced operator support: IN, BETWEEN, LIKE
- ✅ Codecov integration for coverage tracking
- ✅ mypy strict mode with full type safety
- ✅ Comprehensive test suite
- ✅ Improved documentation and examples
- ✅ GitHub Actions CI/CD pipeline
- ✅ PyPI Trusted Publisher (OIDC) for secure releases

## Planned Features

### High Priority

#### Remote Excel Access via Microsoft Graph API
- [ ] Implement `excel+graph://` URL scheme
- [ ] Support for OneDrive and SharePoint Excel files
- [ ] OAuth 2.0 authentication flow
- [ ] Read/write operations on cloud-stored Excel files
- [ ] Caching layer for remote files

**Use case**: Access and query Excel files stored in Microsoft 365 without downloading them locally.

```python
# Future API
engine = create_engine(
    "excel+graph:///sites/mysite/documents/data.xlsx",
    connect_args={
        "tenant_id": "...",
        "client_id": "...",
        "client_secret": "..."
    }
)
```

### Medium Priority

#### Advanced SQL Support
- [ ] **DISTINCT**: Remove duplicate rows
- [ ] **OFFSET**: Pagination support (currently only LIMIT works)
- [ ] **Aggregate functions**: COUNT, SUM, AVG, MIN, MAX
- [ ] **GROUP BY**: Grouping and aggregation
- [ ] **HAVING**: Filtering on aggregated data
- [ ] **Subqueries**: Nested SELECT statements
- [ ] **CTEs (Common Table Expressions)**: WITH clauses

**Status**: These features require significant changes to the excel-dbapi query engine. Aggregate functions are particularly complex due to Excel's storage model.

#### Multi-Table Operations
- [ ] **JOIN support**: INNER JOIN, LEFT JOIN, RIGHT JOIN
- [ ] Cross-sheet queries
- [ ] Foreign key awareness (metadata only, no enforcement)

**Challenge**: Excel has no native concept of relationships or joins. Implementation would require loading and joining data in memory.

#### Performance Optimization
- [ ] Lazy loading for large Excel files
- [ ] Column-level filtering (avoid loading entire rows)
- [ ] Query result caching
- [ ] Batch operation optimization
- [ ] Memory-efficient streaming for large datasets

**Target**: Support Excel files with 100K+ rows without excessive memory usage.

### Low Priority

#### Async Support
- [ ] Async dialect (`excel+aio://`)
- [ ] AsyncEngine and AsyncSession support
- [ ] Non-blocking I/O for file operations

**Note**: Requires asyncio-compatible openpyxl wrapper or alternative Excel library.

#### Additional Features
- [ ] Support for Excel formulas in queries
- [ ] Worksheet-level transactions (via temporary files)
- [ ] ALTER TABLE support (add/remove columns)
- [ ] Index simulation for faster lookups
- [ ] Excel template support (preserve formatting)
- [ ] Multiple sheet joins within same file

## Known Issues and Limitations

### Current Limitations (By Design)
- No transactional rollback (Excel files don't support ACID transactions)
- No concurrent writes (Excel file format limitations)
- Limited SQL feature set compared to traditional RDBMS
- Performance degrades with very large files (>50MB)

### Under Consideration
- **Alternative Excel engines**: Support for xlrd, xlwt, pyexcel in addition to openpyxl
- **CSV fallback**: Automatic conversion to CSV for read-only operations
- **SQLite hybrid mode**: Use SQLite as intermediate cache for complex queries

## Community Feedback

We welcome feedback on this roadmap! Please:
- 🌟 Star the repository if you find it useful
- 💬 Open an issue to suggest new features
- 🐛 Report bugs and edge cases
- 📝 Contribute to documentation improvements
- 🔀 Submit pull requests for planned features

**Priority is driven by community demand** — let us know what you need!

## Versioning Strategy

sqlalchemy-excel follows [Semantic Versioning](https://semver.org/):

- **MAJOR (1.0.0)**: Stable API, production-ready, breaking changes
- **MINOR (0.x.0)**: New features, backward-compatible
- **PATCH (0.0.x)**: Bug fixes, no new features

**Current status**: Beta (0.x.x) — API may change before 1.0.0.

## Long-Term Vision

The goal of sqlalchemy-excel is to:
1. Provide a **seamless SQLAlchemy experience** for Excel files
2. Enable **analysts and developers** to use familiar SQL tools with Excel
3. Bridge the gap between **ad-hoc Excel data** and **structured database workflows**
4. Support **cloud-native Excel** (Microsoft 365, Google Sheets in future)

**Not a goal**: Replace traditional databases for production workloads. Excel is great for prototyping, data analysis, and small-scale applications, but RDBMS should be used for critical systems.

## Contributing to the Roadmap

See [DEVELOPMENT.md](DEVELOPMENT.md) for contribution guidelines.

---

**Last updated**: 2024-01 (v0.2.2)
