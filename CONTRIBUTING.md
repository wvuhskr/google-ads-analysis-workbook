# Contributing

Thank you for improving Google Ads Analysis Workbook. Before proposing a change:

1. Read the [Quickstart](docs/QUICKSTART.md) and [Testing guide](docs/TESTING.md).
2. Keep the exporter read-only with respect to Google Ads and preserve the
   native Sheet as the primary deliverable. Writing output and checkpoints to
   the configured Sheet is expected.
3. Add or update behavior-focused tests for executable changes.
4. Run `npm run check` and `npm test`.
5. Never add advertiser workbooks, logs, Sheet URLs, customer IDs, search-term
   data, performance values, secrets, or internal planning material.
6. Keep public claims tied to tested behavior. Do not describe the project as
   an exhaustive data export, editable account backup, API-independent tool,
   single-step workflow, MCC exporter, automated optimizer, or anonymization
   service.

Use the issue forms for discussion. Sensitive security and privacy reports use
private vulnerability reporting, not public issues.
