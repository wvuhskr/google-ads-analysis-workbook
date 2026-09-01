# Changelog

All notable changes to this project are documented here.

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
