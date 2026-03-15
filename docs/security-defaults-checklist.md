# Security Defaults Checklist

Use this checklist when deploying upload endpoints.

- [ ] Accept only XLSX-compatible content types
- [ ] Enforce max upload size at app/router and reverse proxy layers
- [ ] Keep `defusedxml` installed (required by package)
- [ ] Validate workbook structure before import
- [ ] Keep formula-injection sanitization for exported/template sample values
- [ ] Return structured validation errors (avoid leaking stack traces)
- [ ] Use transactional imports + rollback on failure
- [ ] Restrict import routes with authn/authz in application layer
