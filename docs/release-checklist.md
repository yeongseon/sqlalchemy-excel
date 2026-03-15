# Release Checklist

Use this checklist before cutting any `v*` release.

## 1) Pre-flight
- [ ] `main` branch is green on CI
- [ ] `CHANGELOG.md` updated with release date and highlights
- [ ] Version in `pyproject.toml` matches target tag
- [ ] Roadmap issues for the release milestone are triaged

## 2) Local verification
- [ ] `pip install -e ".[dev]"`
- [ ] `pytest -q`
- [ ] `ruff check src tests`
- [ ] `python -m build`

## 3) Security/operational checks
- [ ] `defusedxml` remains required dependency
- [ ] Upload endpoint limits reviewed (`max_upload_bytes`, content-types)
- [ ] Formula-injection guidance included in docs/examples

## 4) Publish
- [ ] Create and push signed tag (`vX.Y.Z`)
- [ ] Confirm GitHub Actions publish workflow succeeded
- [ ] Confirm package is installable from PyPI (`pip install sqlalchemy-excel==X.Y.Z`)

## 5) Post-release
- [ ] Create/verify GitHub release notes
- [ ] Announce release with migration notes (if needed)
- [ ] Open follow-up issues for deferred roadmap items
