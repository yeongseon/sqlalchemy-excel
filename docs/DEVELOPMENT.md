# Development Guide

This guide covers how to set up a development environment and contribute to sqlalchemy-excel.

## Getting Started

### Clone the Repository

```bash
git clone https://github.com/yeongseon/sqlalchemy-excel.git
cd sqlalchemy-excel
```

### Set Up Development Environment

```bash
make install
```

This creates a virtual environment (`.venv/`), installs all dependencies in editable mode, and sets up pre-commit hooks.

Or manually:

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install -e ".[dev]"
```

This installs:
- `sqlalchemy>=2.0`
- `excel-dbapi>=0.2.0`
- `pytest>=8.0`
- `pytest-cov>=4.0`
- `ruff>=0.4`
- `mypy>=1.10`

## Makefile Commands

| Command | Description |
|---------|-------------|
| `make install` | Bootstrap venv and install dev dependencies |
| `make format` | Format code with ruff |
| `make lint` | Run ruff linting checks |
| `make typecheck` | Run mypy strict type checking |
| `make test` | Run test suite with pytest |
| `make cov` | Run tests with coverage report (HTML + terminal) |
| `make check` | Run lint + typecheck |
| `make check-all` | Run lint + typecheck + tests |
| `make build` | Build distribution packages (sdist + wheel) |
| `make clean` | Remove build artifacts |
| `make clean-all` | Deep clean (caches, coverage, mypy cache) |

## Running Tests

```bash
# All tests
make test

# With coverage report
make cov

# Specific file
.venv/bin/python -m pytest tests/test_dialect.py -v

# Specific test
.venv/bin/python -m pytest tests/test_dialect.py::test_create_engine -v

# Verbose output
.venv/bin/python -m pytest tests/ -v
```

Test coverage target: **95%+** (currently 98%).

## Code Style

### Linting and Formatting

**ruff** handles both linting and formatting:

```bash
make format   # Auto-format
make lint     # Check lint rules
```

Configuration is in `pyproject.toml`:
- Target version: Python 3.10+
- Line length: 88 characters
- Enabled rules: pycodestyle, pyflakes, isort, pep8-naming, pyupgrade, flake8-bugbear, flake8-simplify, flake8-type-checking, ruff-specific

### Type Checking

**mypy** is configured in strict mode:

```bash
make typecheck
```

Configuration (`pyproject.toml`):
- `strict = true`
- `warn_return_any = true`
- `warn_unused_configs = true`

All code must pass strict type checking before merging.

## Project Structure

```
sqlalchemy-excel/
├── src/sqlalchemy_excel/     # Main source code
│   ├── __init__.py           # Package entry point, version
│   ├── dialect.py            # ExcelDialect, ExcelGraphDialect
│   ├── compiler.py           # ExcelCompiler (SQL compilation, HAVING guard)
│   ├── ddl.py                # ExcelDDLCompiler (CREATE/DROP TABLE)
│   ├── types.py              # ExcelTypeCompiler (type mappings)
│   ├── reflection.py         # ExcelInspectionMixin (schema inspection)
│   └── py.typed              # PEP 561 marker file
├── tests/                     # 117 tests (98% coverage)
│   ├── conftest.py           # Shared fixtures
│   ├── test_dialect.py
│   ├── test_compiler.py
│   ├── test_compiler_guards.py
│   ├── test_ddl.py
│   ├── test_dml.py
│   ├── test_e2e.py
│   ├── test_graph_dialect.py
│   ├── test_orm.py
│   ├── test_reflection.py
│   ├── test_reflection_full.py
│   ├── test_type_compiler_full.py
│   └── test_types.py
├── docs/
│   ├── USAGE.md              # Usage guide
│   ├── DEVELOPMENT.md        # This file
│   └── ROADMAP.md            # Project roadmap
├── pyproject.toml            # Project metadata (hatchling)
├── Makefile                  # Development commands
├── README.md
├── CHANGELOG.md
├── CONTRIBUTING.md
└── LICENSE
```

## Development Workflow

1. **Create a feature branch**:
   ```bash
   git checkout -b feature/my-feature
   ```

2. **Make changes** and add tests

3. **Format and lint**:
   ```bash
   make format
   make lint
   ```

4. **Run tests**:
   ```bash
   make test
   ```

5. **Commit changes**:
   ```bash
   git add .
   git commit -m "feat: add new feature"
   ```

6. **Push and create pull request**:
   ```bash
   git push origin feature/my-feature
   ```

## Release Process

sqlalchemy-excel uses **GitHub Releases** with **Trusted Publisher (OIDC)** for PyPI publishing. No API token required.

### Steps

1. Update `CHANGELOG.md` with new version entries.
2. Bump version in `pyproject.toml` and `src/sqlalchemy_excel/__init__.py`.
3. Commit and push:
   ```bash
   git add pyproject.toml src/sqlalchemy_excel/__init__.py CHANGELOG.md
   git commit -m "chore: bump version to X.Y.Z"
   git push origin main
   ```
4. Create a **GitHub Release** (via the GitHub UI or `gh release create vX.Y.Z`).
5. The `publish-pypi.yml` workflow triggers automatically, builds, validates, and publishes to PyPI.

### Verify

- PyPI: https://pypi.org/project/sqlalchemy-excel/
- GitHub Releases: https://github.com/yeongseon/sqlalchemy-excel/releases

## Continuous Integration

The CI pipeline (`.github/workflows/ci.yml`) runs on every push and pull request:

1. **Linting**: ruff
2. **Type checking**: mypy (strict mode)
3. **Testing**: pytest on Python 3.10, 3.11, 3.12, 3.13
4. **Coverage**: Upload to Codecov

All checks must pass before merging.

## Contributing Guidelines

- Write tests for all new features and bug fixes
- Maintain or improve code coverage (target: **95%+**)
- Follow the existing code style (enforced by ruff)
- Add type hints for all functions (enforced by mypy strict mode)
- Update documentation for user-facing changes
- Keep commits atomic with clear messages (`feat:`, `fix:`, `docs:`, `chore:`)

## Getting Help

- Open an issue: https://github.com/yeongseon/sqlalchemy-excel/issues
- Check existing discussions and issues
- Review the main README and USAGE guide
