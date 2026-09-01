# Quickstart

Google Ads Analysis Workbook creates a documented, resumable Google Sheet structured
for authorized human or LLM-assisted analysis. The exporter is read-only with
respect to Google Ads and writes its output and checkpoints to the configured
Google Sheet. It does not change Google Ads account settings.

It runs without separately provisioning a Google Ads API developer token,
OAuth client, Google Cloud project, paid connector, or MCP server. Google
authorization is still required for the advertiser account and destination
Sheet.

Run this version only from the individual advertiser account whose data you
want to export. Never run it from a Google Ads manager account (MCC). If you
access the advertiser through a manager account, switch into the client
advertiser account first.

It is not an editable Google Ads account backup, real-time or atomic snapshot,
automated optimizer, exhaustive export of every Google Ads resource, or
anonymization tool. It does not connect directly to an LLM or automatically
create an XLSX file.

## Requirements

- Permission to run Google Ads Scripts in the intended advertiser account.
- Permission to create and edit a blank Google Sheet for that advertiser.
- Node.js 22 or 24 LTS only for the optional local XLSX sanitizer. Node.js 24
  LTS is recommended.

## Export to a native Google Sheet

1. Create a blank Google Sheet and copy its URL.
2. Create a Google Ads Script, open
   `google-ads-analysis-workbook-v1.0.0.js`, and paste all of its contents into
   the script editor.
3. Set `CONFIG.SPREADSHEET_URL` to the blank Sheet URL.
4. Save and use **Preview** first. Preview is read-only compatibility evidence;
   it does not complete an export.
5. Use **Run** to begin the export. If it is resumable, run it again within 24
   hours of the original start without editing, renaming, or deleting
   exporter-owned tabs. If the checkpoint expires, follow
   `_export_info.next_action`.
6. Apply the `_export_info` acceptance gate below.

Proceed to **Run** only when Preview reaches `Diagnostics complete`, reports
`native_sheet_target: SUPPORTED`, and you have confirmed that Google Ads is
showing the intended individual advertiser account. Stop if Preview ends with
an error before that point, if the native Sheet target is not supported, or if
the selected advertiser is wrong.

`SUPPORTED; rows_read=0` is normal: Google accepted that probe but no matching
row existed in the short diagnostic range. `UNSUPPORTED_OR_ERROR` identifies a
resource that needs review. Some such resources are optional and will become a
`LIMITED` row during the export; others may fail a required export job. If the
resource is necessary for your analysis, stop. Otherwise you may Run, but the
final `_export_info` gate below remains binding.

The completed native Google Sheet is the primary deliverable. Performance tabs
cover exactly 90 complete days ending yesterday. Change history covers the most
recent 28 complete days. Structure and configuration tabs are sequential
point-in-time reads, not one atomic account snapshot.

## `_export_info` acceptance gate

Use `_export_info.workbook_status` as the primary technical go/no-go field and
follow `_export_info.next_action`:

- `READY`: the workbook passed the exporter's technical completion checks and
  is available for authorized use, analysis, or download.
- `READY_WITH_LIMITATIONS`: use it only after you review every `LIMITED` row and
  decide that every limitation is acceptable for the intended analysis.
- `IN_PROGRESS`: do not analyze, download, or share it. Follow `next_action` to
  resume the export.
- `NEEDS_REVIEW`: do not analyze, use, download, or share it. Review
  `overall_status`, every error or preserved-data row, and `next_action`.

`COMPLETE_WITH_ERRORS` maps to `NEEDS_REVIEW`; it is not success. Do not use or
share an export whose `overall_status` is anything other than `COMPLETE` or
`COMPLETE_WITH_LIMITATIONS`. Follow the exact `next_action` value.

`READY` is a technical status. It does not grant permission to share or upload
the workbook. Confirm advertiser authorization, organizational policy, and the
receiving provider's terms separately.

## Optional downloaded-XLSX sanitizer

After the native workbook is technically ready, download it as XLSX and run:

```sh
node sanitize-downloaded-xlsx-v1.0.0.js downloaded.xlsx downloaded-sanitized.xlsx
```

The sanitizer uses Node.js built-ins and makes no network requests. It corrects
unintended converter-created hyperlinks and verifies package invariants. It
does not anonymize, redact, de-identify, or authorize sharing of the workbook.
Read [Data handling](DATA-HANDLING.md) before sending output to another person
or service.

Repository: https://github.com/wvuhskr/google-ads-analysis-workbook

Public bug reports: https://github.com/wvuhskr/google-ads-analysis-workbook/issues/new/choose

Private security or privacy reports: https://github.com/wvuhskr/google-ads-analysis-workbook/security/advisories/new

This independent project is not affiliated with or endorsed by Google.
