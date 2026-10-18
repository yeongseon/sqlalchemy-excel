# Contributing Guide

We welcome contributions to the `sqlalchemy-excel` project.

## Issue Scope

- Explain the problem, proposed work, acceptance criteria, and priority using
  verified evidence and user impact in this repository.
- Apply reusable ideas to this repository's actual code, configuration, and
  conventions. Another project's policy or implementation is not evidence that
  the same requirement applies here.
- Remove copied project names, issue links, process labels, and release
  assumptions that are unrelated to the requested work.
- Keep relevant technical evidence and source attribution, including actual
  dependencies, integrations, upstream defects, and official documentation.
  Explain the relationship and the work owned by this repository. Apply the
  same local-evidence requirement to explicitly requested comparisons.
- Before creating or editing an issue, search for duplicates and review every
  project name and repository link in the title and body. The issue must be
  understandable without following an unrelated project's issue or policy.
- After saving, fetch the issue again and verify the saved title, body, and
  labels before reporting completion. Apply these checks to issues created
  through the API or CLI as well as through the issue forms.


## Branch Strategy

Use GitHub Flow and branch from `main`.

Recommended branch prefixes:

- `feat/` for new features
- `fix/` for bug fixes
- `docs/` for documentation-only changes
- `chore/` for tooling and maintenance
- `ci/` for workflow updates

## Development Workflow

1. Create a branch from `main`.
   ```bash
   git checkout main
   git pull origin main
   git checkout -b feat/your-feature-name
   ```
2. Write code and tests.
3. Run the local quality gate.
   ```bash
   make check-all
   ```
4. Push and create a pull request.
   ```bash
   git push origin feat/your-feature-name
   ```

## Project Commands

```bash
make format      # Format code with ruff
make lint        # Lint with ruff
make typecheck   # Type check with mypy
make test        # Run tests
make cov         # Run tests with coverage
make check-all   # Run the full local gate
```

## Commit Message Guidelines

We follow the [Conventional Commits](https://www.conventionalcommits.org/) specification.

### Examples

```bash
git commit -m "feat: add Excel dialect type mapping"
git commit -m "fix: handle NULL column types gracefully"
git commit -m "docs: improve ORM usage examples"
git commit -m "refactor: simplify compiler logic"
git commit -m "chore: update dev dependencies"
```

Use imperative present tense and keep the message concise.

## Code of Conduct

Be respectful and inclusive. See our [Code of Conduct](CODE_OF_CONDUCT.md) for details.
