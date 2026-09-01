# Data handling

The exported native Google Sheet and any downloaded XLSX contain confidential
advertiser data. That can include performance values, configuration, targeting,
search terms, landing-page URLs, audience context, and change information.
Store, share, upload, and retain an output only with advertiser authorization
and under your organization's data policy.

## Data flow

The exporter reads account data through Google Ads Scripts and writes output
and checkpoints to the Google Sheet identified by `CONFIG.SPREADSHEET_URL`. It
is read-only with respect to Google Ads. The exporter includes no built-in
integration that uploads the workbook to an AI assistant, paid connector, MCP
server, or other third-party endpoint.

Google authorization is still required for the advertiser account and
destination Sheet. After the export, the user controls whether the Sheet is
shared, downloaded, or uploaded elsewhere.

The optional XLSX sanitizer runs locally with Node.js built-ins and makes no
network requests. It reads the downloaded workbook and writes a corrected file
to the local paths supplied by the user.

## Readiness is not sharing authorization

`_export_info.workbook_status` describes technical export readiness. `READY`
or `READY_WITH_LIMITATIONS` does not grant permission to share or upload the
workbook. It also does not evaluate a recipient's terms, retention policy,
security controls, or authorization to receive advertiser data.

## Sanitized is not anonymized

In this project, **sanitized** means the optional XLSX tool corrected
converter-created hyperlinks and checked its package invariants. It does not
mean the workbook was anonymized, redacted, de-identified, or made appropriate
for unrestricted sharing. Review every workbook before sending it to another
person or service.

## Chatbots and storage providers

You are responsible for deciding whether an export may be uploaded to a
chatbot, AI assistant, storage provider, ticketing system, or any other third
party. Confirm the advertiser authorization, your organization's policy, and
the provider's terms, retention settings, and data-handling practices first.

## Public support boundary

Do not attach workbooks, screenshots of live data, Google Sheet URLs, customer
IDs, campaign or search-term data, or performance values to public issues.
For a security or privacy concern, use GitHub private vulnerability reporting
at
https://github.com/wvuhskr/google-ads-analysis-workbook/security/advisories/new as
described in [SECURITY.md](SECURITY.md). For non-sensitive bugs, use
https://github.com/wvuhskr/google-ads-analysis-workbook/issues/new/choose.

This independent project is not affiliated with or endorsed by Google.
