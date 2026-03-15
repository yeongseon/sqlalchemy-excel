# sqlalchemy-excel — Maturity & Growth Roadmap (2026)

## 0) Goal
- Education-ready in 6 months
- Production-hardened in 12 months
- Clear boundary: **ORM-driven Excel workflow toolkit** (template/validate/import/export), not SQLAlchemy dialect

---

## 1) 30-Day Execution Plan (compressed)

### Week 1 — Packaging & release bootstrap
- [ ] Publish first stable package to PyPI (Trusted Publishing)
- [ ] Create first GitHub release/tag with release notes
- [ ] Add release checklist template (`docs/release-checklist.md`)
- [ ] Verify README install instructions for extras

**Definition of Done**
- `pip install sqlalchemy-excel` works
- Version/tag/release notes aligned

### Week 2 — CI and confidence
- [ ] Add CI smoke matrix for OS (ubuntu/windows/macos)
- [ ] Ensure coverage publishing is deterministic
- [ ] Add import strategy tests (insert/upsert/dry-run)
- [ ] Add FastAPI router integration tests

**Definition of Done**
- CI matrix green with stable coverage pipeline

### Week 3 — Education UX
- [ ] Add `10-minute backend upload pipeline` tutorial
- [ ] Add FastAPI+SQLite starter example (copy/paste)
- [ ] Add clear scope boundaries to README (non-goals)
- [ ] Add comparison table vs pandas/openpyxl-only workflows

**Definition of Done**
- Learner can run template→validate→import in <10 minutes

### Week 4 — Security/operations
- [ ] Validate file size/content-type defaults in docs and code path
- [ ] Ensure formula-injection sanitization examples are explicit
- [ ] Add “prod checklist” for upload endpoints
- [ ] Define interoperability contract with excel-dbapi version range

**Definition of Done**
- Safe defaults and operational notes are explicit and test-backed

---

## 2) Top-10 Priority Actions (issue-ready)

1. **PyPI publish + release tags**
   - Labels: `release`, `high-priority`
2. **Trusted Publishing verification**
   - Labels: `security`, `release`
3. **First stable changelog/release notes format**
   - Labels: `release`, `docs`
4. **OS CI smoke matrix**
   - Labels: `ci`
5. **Import mode test hardening (insert/upsert/dry-run)**
   - Labels: `test`, `importer`
6. **FastAPI router e2e tests**
   - Labels: `test`, `fastapi`
7. **Education quickstart pack**
   - Labels: `docs`, `education`
8. **Security defaults checklist (size/type/xml/formula)**
   - Labels: `security`
9. **excel-dbapi compatibility contract tests**
   - Labels: `interop`, `test`
10. **Community templates + public roadmap issues**
   - Labels: `community`, `docs`

---

## 3) v1.0 / v1.1 Release Checklists

### v1.0 (Education-ready)
- [ ] Package published and installable
- [ ] Template→validate→import path stable
- [ ] CLI + FastAPI starter examples verified
- [ ] Clear non-goal boundaries in README/PRD
- [ ] Release notes and versioning policy documented

### v1.1 (Production-hardened)
- [ ] Large-upload guidance (and optional streaming path) documented
- [ ] Upsert portability notes by dialect explicit
- [ ] Security checklist enforced with tests
- [ ] Interop compatibility with excel-dbapi locked and tested

---

## 4) PRD/TDD/ARCH Sync Tasks
- [ ] PRD: Explicit education + backend segments, non-goals highlighted
- [ ] TDD: Security/error-report/import strategy tests linked to claims
- [ ] ARCH: Dual-channel session and dependency contract tightened
- [ ] README: mirrors PRD one-liner and quickstart path

---

## 5) Suggested Milestone Names
- `M1: First Public Release`
- `M2: CI/Test Confidence`
- `M3: Education Quickstart`
- `M4: Security & Interop`

---

## 6) Success Metrics
- Time-to-first-success: <10 min
- CI pass rate: >95%
- Upload-pipeline bug re-open rate: down by 50%
- Monthly release cadence maintained
