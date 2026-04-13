# Release Notes Template

Use this template when publishing a new `sqlalchemy-excel` release and coordinating with `excel-dbapi`.

## Release Metadata

- Version: `X.Y.Z`
- Release date: `YYYY-MM-DD`
- Related issues/PRs:
  - `#...`
  - `#...`

## Summary

- <one-line release goal>
- <one-line impact for users>

## Added

- <new capability>
- <new capability>

## Changed

- <behavior or compatibility change>
- <internal/compiler/runtime change>

## Fixed

- <bug fix>
- <bug fix>

## Dependency Changes

- `excel-dbapi`: `<old range>` -> `<new range>`
- `sqlalchemy`: `<old range>` -> `<new range>`
- Optional extras impacted:
  - `graph`: `<notes>`

### Validation Checklist

- `pyproject.toml` dependency ranges updated
- Local install tested for both core and `graph` extras
- Cross-repo compatibility verified against targeted `excel-dbapi` version

## Cross-repo Features

List features that require coordinated releases between `sqlalchemy-excel` and `excel-dbapi`.

| Feature | sqlalchemy-excel change | excel-dbapi change | Minimum versions | Notes |
| --- | --- | --- | --- | --- |
| <feature name> | <PR/commit> | <PR/commit> | `sqlalchemy-excel>=X.Y.Z`, `excel-dbapi>=A.B.C` | <rollout notes> |

## Migration Steps

1. Upgrade dependencies:
   - `pip install -U sqlalchemy-excel[graph]`
2. Verify runtime versions:
   - `python -c "import sqlalchemy_excel, excel_dbapi; print(sqlalchemy_excel.__version__, excel_dbapi.__version__)"`
3. Re-run smoke checks:
   - dialect connection
   - basic CRUD
   - any feature-specific regression tests
4. Review breaking/behavior changes listed above and update application SQL usage if needed.

## Verification Evidence

- Tests: ``
- Type check: ``
- Additional checks: ``

## Rollback Plan

- Pin to prior known-good versions:
  - `sqlalchemy-excel==<previous>`
  - `excel-dbapi==<previous>`
- Re-run smoke checks and confirm service recovery.
