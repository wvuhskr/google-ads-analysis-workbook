# Maintainer release procedure

This procedure prepares release bytes without creating a remote repository,
pushing, publishing, or changing public Git state without approval.

## 1. Freeze and inspect the candidate

1. Reconcile every tracked and untracked change. Preserve intentional work and
   remove private evidence from the public candidate rather than copying the
   internal working directory wholesale.
2. Confirm the only required user configuration is
   `CONFIG.SPREADSHEET_URL` and that the exporter is scoped to one individual
   advertiser account at a time.
3. Review every public path, generated asset, and documentation claim. The
   candidate must contain only synthetic examples and the allowlisted files
   enforced by `scripts/audit-public-package.mjs`.

## 2. Run pre-commit gates

Run these commands on the candidate bytes:

```sh
npm run check
npm test
git diff --check
gitleaks dir . --redact --no-banner
```

Copy the candidate into a fresh temporary repository, create a disposable
synthetic commit and annotated `v1.0.0` tag there, and run the strict release
builder and audit against that temporary root. This proves the clean-root,
source-byte, archive, checksum, and tag gates without changing the final local
repository before approval.

Independently inspect the ZIP member list, recalculate every SHA-256 value,
compare released JavaScript bytes with their source files, and scan the fresh
temporary repository's complete reachable Git history with `gitleaks git`.

## 3. Complete live validation

1. Create a blank Google Sheet dedicated to the test.
2. Run Preview from the intended individual advertiser account and confirm the
   compatibility diagnostics contain no unsupported required resource.
3. Run or resume the export until `_export_info.overall_status` is terminal.
4. Independently inspect the native workbook, including
   `_export_info.workbook_status`, `next_action`, the complete 41-sheet
   topology, declared versus physical rows, date windows, limitations, errors,
   hidden tabs, temporary tabs, and synthetic-or-authorized evidence handling.
5. Download the completed Sheet as XLSX, run the optional sanitizer against
   that real download outside the repository, and verify its reported hash and
   package audit. Never add the workbook, logs, URLs, IDs, or metrics to Git.

## 4. Obtain independent reviews

Obtain both a marketer-adoption review and a technical and privacy review.
Resolve all material findings and rerun every affected gate. Review reports may
remain outside the public package, but the final handoff must summarize their
scope and disposition.

## 5. Approval-controlled local Git steps

After every pre-commit, privacy, live-workbook, sanitizer, and independent
review gate passes, obtain explicit approval before creating or replacing the
final local commit or annotated tag.

After approval, create a fresh reviewed public root using a synthetic release
identity, create the approved commit and annotated `v1.0.0` tag, run
`npm run build:release` and `npm run audit:release`, and generate the final
public-file and release-asset hash manifest from those exact bytes.

Do not create a remote repository, push, publish, upload release assets, or
rewrite any existing repository history unless that separate action is
explicitly approved.
