# Data handling

The exported native Google Sheet and any downloaded XLSX contain confidential
advertiser data. Store, share, upload, and retain them only with advertiser
authorization and under your organization's data policy.

## Data flow

The exporter reads through Google Ads Scripts and writes output and checkpoints
to the configured Google Sheet. It is read-only with respect to Google Ads and
includes no built-in integration that uploads the workbook to an AI assistant,
paid connector, MCP server, or other third-party endpoint. Google authorization
is still required. The user controls any later sharing, download, or upload.

The optional XLSX sanitizer runs locally with Node.js built-ins and makes no
network requests.

## Readiness is not sharing authorization

`_export_info.workbook_status` describes technical readiness. `READY` or
`READY_WITH_LIMITATIONS` does not grant permission to share or upload the
workbook and does not evaluate the receiving provider's terms or retention.

## Sanitized is not anonymized

Here, **sanitized** means the optional XLSX tool corrected converter-created
hyperlinks and checked package invariants. It does not mean anonymized,
redacted, de-identified, or appropriate for unrestricted sharing.

## Chatbots and storage providers

Before sending an export to another person or service, confirm advertiser
authorization, organizational policy, and the provider's terms, retention
settings, and data-handling practices.

## Sensitive reports

Do not post workbooks, Google Sheet URLs, customer IDs, campaign or search-term
data, or performance values publicly.

Repository and policies: https://github.com/wvuhskr/google-ads-analysis-workbook

Non-sensitive bug reports: https://github.com/wvuhskr/google-ads-analysis-workbook/issues/new/choose

Private security or privacy reports: https://github.com/wvuhskr/google-ads-analysis-workbook/security/advisories/new

Private vulnerability reporting is available only after the repository is
published and that GitHub feature is enabled. If the private route is
unavailable, do not disclose sensitive details publicly; wait for a private
maintainer route to be enabled.

This independent project is not affiliated with or endorsed by Google.
