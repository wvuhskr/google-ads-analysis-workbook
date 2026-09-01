# Testing

Use Node.js 22 or 24 LTS. Node.js 24 LTS is recommended for local development.
Release maintainers also need Info-ZIP zip 3.x and Info-ZIP unzip 6.x. The
runtime exporter and optional sanitizer do not depend on those tools; they are
used only to build and independently inspect the release archive.

```sh
npm run check
npm test
npm run build:release
npm run audit:release
```

`npm run check` parses the exporter and sanitizer. `npm test` runs the
dependency-free regression and public-release contract tests. The release
builder creates local assets in `release-assets/v1.0.0/`; the audit checks
allowed contents, checksums, common privacy-boundary failures, and required
product-scope disclosures.

`npm run build:release` is intentionally strict. It requires a clean Git tree,
an annotated `v1.0.0` tag directly at `HEAD`, and release-input bytes identical
to `HEAD`. CI uses `npm run build:ci`, which skips only the tag requirement; it
still requires a clean tree and exact `HEAD` input bytes.

The builder will not overwrite a release directory. For a safe fresh local
build without deleting or replacing existing assets, use a newly created empty
temporary directory and audit that exact output:

```sh
release_dir="$(mktemp -d)"
node scripts/build-release.mjs "$release_dir"
node scripts/audit-public-package.mjs . "$release_dir"
```

These local checks do not replace a real Google Ads Scripts export to a blank
workbook. Before distributing a workbook, run a fresh export to a blank
workbook, verify its terminal `_export_info` status, and independently review
the completed native Sheet. Then, if an XLSX is needed, download that completed
Sheet and run the sanitizer end to end against the real download before sharing
either artifact.

Maintainers must also follow the complete [release procedure](RELEASING.md),
including privacy, independent-review, approval, and final-manifest gates.
