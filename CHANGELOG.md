# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.0.0] - 2026-03-15
### Added
- OS CI smoke matrix and deterministic coverage upload.
- FastAPI router e2e tests for template/validate/import/health flows.
- Security guardrails for FastAPI uploads (`max_upload_bytes`, content type checks).
- Release checklist, security checklist, and 10-minute backend upload tutorial.
- Community contribution templates (bug report, feature request, roadmap task, PR template).
- excel-dbapi compatibility contract test and explicit dependency range.

### Changed
- Upgraded package metadata to stable `1.0.0` release posture.
- Improved importer mode tests for insert/upsert/dry-run failure semantics.
