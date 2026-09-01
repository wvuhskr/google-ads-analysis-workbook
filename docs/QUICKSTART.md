# Quickstart

Google Ads Analysis Workbook creates a documented, resumable Google Sheet structured
for authorized human or LLM-assisted analysis. The exporter is read-only with
respect to Google Ads and writes its output and checkpoints to the configured
Google Sheet. It does not change campaigns, budgets, bids, targeting, ads, or
conversion settings.

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
- A blank Google Sheet dedicated to this export.
- Node.js 22 or 24 LTS only if you want to run the optional downloaded-XLSX
  sanitizer. Node.js 24 LTS is recommended.

## Export to a native Google Sheet

1. Create a blank Google Sheet and copy its URL.
2. Create a Google Ads Script and paste the contents of
   `google-ads-analysis-workbook.js`.
3. Set `CONFIG.SPREADSHEET_URL` to the blank Sheet URL.
4. Save the script and choose **Preview**. Review the compatibility diagnostics.
   Preview is read-only and does not complete an export.
5. Choose **Run**. The script writes its workbook tabs and checkpoints its
   progress when another manually started run is needed.
6. If the run is resumable, run it again within 24 hours of the original start.
   Do not edit, rename, or delete exporter-owned tabs while it is running or
   resumable. If the checkpoint expires, follow `_export_info.next_action` to
   reset it safely and start a replacement export.
7. Apply the `_export_info` acceptance gate below.

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

The primary deliverable is the completed native Google Sheet. Performance tabs
contain exactly 90 complete days ending yesterday, and weekly tabs use those
same boundaries. Change history covers the most recent 28 complete days.
Structure and configuration tabs are sequential point-in-time reads taken
during the export, not one atomic account snapshot.

## `_export_info` acceptance gate

Use `_export_info.workbook_status` as the primary technical go/no-go field and
always follow `_export_info.next_action`:

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

Only after the native workbook is technically ready, download it as XLSX and
run:

```sh
node tools/sanitize-downloaded-xlsx.js downloaded.xlsx downloaded-sanitized.xlsx
```

The sanitizer uses only Node.js built-ins and makes no network requests. It
corrects unintended converter-created hyperlinks, preserves expected workbook
navigation, and verifies package invariants. It does not anonymize, redact,
de-identify, or authorize sharing of the data. Read
[Data handling](../DATA-HANDLING.md) before sending the output anywhere.

For support, releases, and policies, use the
[canonical repository](https://github.com/wvuhskr/google-ads-analysis-workbook).

This independent project is not affiliated with or endorsed by Google.
