# Development Guide

This guide covers how to set up a development environment and contribute to sqlalchemy-excel.

## Getting Started

### Clone the Repository

```bash
git clone https://github.com/yeongseon/sqlalchemy-excel.git
cd sqlalchemy-excel
```

### Set Up Virtual Environment

Create and activate a virtual environment:

```bash
# Using venv (Python 3.10+)
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate

# Or using virtualenv
virtualenv .venv
source .venv/bin/activate
```

### Install Development Dependencies

Install the package in editable mode with development dependencies:

```bash
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

The project includes a Makefile for common development tasks:

| Command | Description |
|---------|-------------|
| `make install` | Install the package in editable mode with dev dependencies |
| `make format` | Format code with ruff |
| `make lint` | Run linting checks with ruff and mypy |
| `make test` | Run tests with pytest |
| `make coverage` | Run tests with coverage report |
| `make build` | Build distribution packages |
| `make clean` | Remove build artifacts and cache files |

## Running Tests

Run the test suite using pytest:

```bash
# Run all tests
pytest

# Or use make
make test

# Run with coverage
pytest --cov=sqlalchemy_excel --cov-report=html
make coverage

# Run specific test file
pytest tests/test_dialect.py

# Run specific test
pytest tests/test_dialect.py::test_create_engine

# Run with verbose output
pytest -v
```

Test files are located in the `tests/` directory.

## Code Style

sqlalchemy-excel follows strict code quality standards:

### Linting and Formatting

**Ruff** is used for both linting and formatting:

```bash
# Format code (auto-fix)
ruff format .

# Check linting (without fixing)
ruff check .

# Auto-fix linting issues
ruff check --fix .

# Or use make
make format  # Format + auto-fix
make lint    # Check without fixing
```

Configuration is in `pyproject.toml`:
- Target version: Python 3.10+
- Line length: 88 characters
- Enabled rules: pycodestyle, pyflakes, isort, pep8-naming, pyupgrade, flake8-bugbear, flake8-simplify, flake8-type-checking, ruff-specific

### Type Checking

**mypy** is configured in strict mode:

```bash
mypy src/sqlalchemy_excel

# Or as part of make lint
make lint
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
│   ├── __init__.py           # Package entry point
│   ├── dialect.py            # ExcelDialect implementation
│   ├── compiler.py           # SQL compilation (ExcelCompiler, DDLCompiler, TypeCompiler)
│   ├── types.py              # Type mappings
│   └── py.typed              # PEP 561 marker file
├── tests/                     # Test suite
│   ├── test_dialect.py
│   ├── test_compiler.py
│   └── fixtures/             # Test data files
├── docs/                      # Documentation
│   ├── USAGE.md
│   ├── DEVELOPMENT.md
│   └── ROADMAP.md
├── pyproject.toml            # Project metadata and config
├── README.md                 # Main documentation
├── CHANGELOG.md              # Version history
├── LICENSE                   # MIT License
└── Makefile                  # Development commands
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

sqlalchemy-excel uses GitHub Actions for automated releases:

### 1. Update CHANGELOG.md

Document all changes in the changelog following the format:

```markdown
## [0.3.0] - 2024-01-15

### Added
- New feature X
- Support for Y

### Fixed
- Bug Z
```

### 2. Bump Version

Update the version in `pyproject.toml`:

```toml
[project]
version = "0.3.0"
```

### 3. Create Git Tag

```bash
git add pyproject.toml CHANGELOG.md
git commit -m "chore: bump version to 0.3.0"
git tag v0.3.0
git push origin main
git push origin v0.3.0
```

### 4. Automated Publishing

When you push a tag (`v*`), GitHub Actions automatically:
1. Runs all tests and linting
2. Builds distribution packages (`sdist` and `wheel`)
3. Publishes to PyPI using **Trusted Publisher (OIDC)**

**No API token needed** — the project uses PyPI's Trusted Publisher feature with OIDC authentication configured in GitHub Actions.

### 5. Verify Release

Check that the release appears on:
- PyPI: https://pypi.org/project/sqlalchemy-excel/
- GitHub Releases: https://github.com/yeongseon/sqlalchemy-excel/releases

## Continuous Integration

The CI pipeline (`.github/workflows/ci.yml`) runs on every push and pull request:

1. **Linting**: ruff + mypy
2. **Testing**: pytest on Python 3.10, 3.11, 3.12, 3.13
3. **Coverage**: Upload to Codecov

All checks must pass before merging.

## Contributing Guidelines

- Write tests for all new features and bug fixes
- Maintain or improve code coverage (target: 90%+)
- Follow the existing code style (enforced by ruff)
- Add type hints for all functions (enforced by mypy strict mode)
- Update documentation for user-facing changes
- Keep commits atomic and write clear commit messages

## Getting Help

- Open an issue: https://github.com/yeongseon/sqlalchemy-excel/issues
- Check existing discussions and issues
- Review the main README and USAGE guide
