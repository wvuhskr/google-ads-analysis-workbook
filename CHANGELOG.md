# Changelog

All notable changes to this project are documented here.

## v1.0.2

- Scheduled runs no longer fail forever after an interrupted export. When the
  saved checkpoint is older than `MAX_RESUME_AGE_HOURS` and belongs to the same
  account and workbook, `main()` now discards it and starts a replacement
  export instead of stopping with "Saved export state has expired". Checkpoints
  for a different account or workbook still stop the run. The status write
  that follows a discard or reset is retried.
- A transient Google Sheets error while refreshing `_export_info` no longer
  ends the run. The refresh is retried, and a periodic mid-run refresh that
  still fails is logged and skipped. Refreshes that record a status change
  still fail loudly.
- Drop-in upgrade: replace the script and keep your existing workbook and
  schedule. The exporter still writes `v1.0.0` into the workbook as its
  compatibility marker, so no reset is needed. The sanitizer is unchanged.

## v1.0.1

- Corrected the MIT License copyright holder to Alex Murtha. No exporter or
  sanitizer code changes; the scripts are byte-identical to v1.0.0 and still
  self-report v1.0.0.

## v1.0.0

- Initial public package for one individual Google Ads advertiser account at a
  time.
- Read-only Google Ads workflow that writes a documented native Google Sheet
  for authorized human or LLM-assisted analysis.
- Exactly 90 complete performance days, with a disclosed 28-complete-day
  change-history exception and sequential configuration snapshots.
- Resumable workbook checkpoints, explicit `_export_info` readiness states,
  per-tab row and limitation reporting, and embedded data and field
  dictionaries.
- Selected performance, structure, targeting, negative keyword, audience,
  conversion, creative, landing-page, Quality Score, and change context.
- Sensitive change actor and old/new value details disabled by default.
- Optional local XLSX hyperlink correction and package verification using
  Node.js built-ins. The tool does not anonymize the workbook.
