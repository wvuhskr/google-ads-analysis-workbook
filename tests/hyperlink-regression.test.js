'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  loadExporter,
  createPersistentRichTextHarness,
  simulatedDownloadedLinks,
} = require('./load-exporter');

function linkedCharacters(richText) {
  if (!richText) return [];
  const text = richText.getText();
  const linked = [];
  for (let index = 0; index < text.length; index += 1) {
    const url = richText.getLinkUrl(index, index + 1);
    if (url !== null) linked.push({ index, url });
  }
  return linked;
}

function writeQuotePrefixedRows(sheet, rows) {
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows.map((row) => (
    row.map((value) => (
      typeof value === 'string' && value !== '' ? `'${value}` : value
    ))
  )));
}

test('report formatting preserves quotePrefix for native plain-text safety', () => {
  // Production mutation caught: reintroducing a rich-text rewrite can strip
  // native Sheet quotePrefix before the separate XLSX sanitization boundary.
  const harness = createPersistentRichTextHarness({
    omitBuilderSetLinkUrl: true,
    richTextClearsQuotePrefix: true,
  });
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign');
  const headers = Array.from(api.headersForJob(job));
  const sheet = harness.createSheet('campaign', [headers.map(() => '')]);
  writeQuotePrefixedRows(sheet, [headers]);

  assert.doesNotThrow(() => api.formatReportSheet(sheet, job, 'USD'));

  const column = headers.indexOf('campaign.id') + 1;
  const cell = sheet.ensureCell(1, column);
  assert.equal(cell.value, 'campaign.id');
  assert.equal(cell.quotePrefix, true);
  assert.deepEqual(linkedCharacters(cell.richText), []);
});

test('both metadata dictionaries retain native quotePrefix on dotted field values', () => {
  // Production mutation caught: protecting only report headers leaves field and
  // key metadata vulnerable to native formula or automatic-link interpretation.
  const harness = createPersistentRichTextHarness({
    omitBuilderSetLinkUrl: true,
    richTextClearsQuotePrefix: true,
  });
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const manifest = api.getManifestDefinition();
  const fixtures = [
    {
      tab: '_data_dictionary',
      rows: api.buildDataDictionaryRows(manifest).map((row) => Array.from(row)),
    },
    {
      tab: '_field_dictionary',
      rows: api.buildFieldDictionaryRows(manifest).map((row) => Array.from(row)),
    },
  ];

  for (const fixture of fixtures) {
    const job = manifest.find((candidate) => candidate.tab === fixture.tab);
    const sheet = harness.createSheet(
      fixture.tab,
      fixture.rows.map((row) => row.map(() => '')),
    );
    writeQuotePrefixedRows(sheet, fixture.rows);

    assert.doesNotThrow(() => api.formatReportSheet(sheet, job, 'USD'));

    const fieldColumn = fixture.rows[0].indexOf(fixture.tab === '_field_dictionary' ? 'field' : 'keys');
    const dataRow = fixture.rows.findIndex((row, index) => (
      index > 0 && String(row[fieldColumn]).includes('campaign.id')
    )) + 1;
    assert.ok(dataRow > 1, `${fixture.tab} fixture needs dotted field metadata`);
    const cell = sheet.ensureCell(dataRow, fieldColumn + 1);
    assert.equal(cell.quotePrefix, true);
    assert.deepEqual(linkedCharacters(cell.richText), []);
  }
});

test('URL body strings remain exact native Sheet text before XLSX sanitization', () => {
  // Production mutation caught: exempting explicit URLs from quotePrefix changes
  // the native Sheet safety contract before the downloaded package is sanitized.
  const harness = createPersistentRichTextHarness({ richTextClearsQuotePrefix: true });
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'ads');
  const headers = Array.from(api.headersForJob(job));
  const urlColumn = headers.indexOf('ad_group_ad.ad.final_urls');
  const url = 'https://example.test/landing-page?a=1,b=2';
  const body = headers.map(() => '');
  body[urlColumn] = url;
  const sheet = harness.createSheet('ads', [headers, body].map((row) => row.map(() => '')));
  writeQuotePrefixedRows(sheet, [headers, body]);

  api.formatReportSheet(sheet, job, 'USD');

  const cell = sheet.ensureCell(2, urlColumn + 1);
  assert.equal(cell.value, url);
  assert.equal(cell.quotePrefix, true);
  assert.deepEqual(linkedCharacters(cell.richText), []);
});

test('download simulation does not mistake quotePrefix for an XLSX unlink guarantee', () => {
  // Characterization of the live Google converter: native Sheet state can be
  // unlinked while the downloaded package reinfers campaign.id and URL links.
  const harness = createPersistentRichTextHarness({ richTextClearsQuotePrefix: true });
  const sheet = harness.createSheet('metadata', [['', '']]);

  writeQuotePrefixedRows(sheet, [[
    'campaign.id',
    'https://example.test/landing-page',
  ]]);

  assert.deepEqual(simulatedDownloadedLinks(sheet.ensureCell(1, 1)), [{
    start: 0, end: 11, url: 'http://campaign.id',
  }]);
  assert.deepEqual(simulatedDownloadedLinks(sheet.ensureCell(1, 2)), [{
    start: 0, end: 33, url: 'https://example.test/landing-page',
  }]);
});

test('download simulation preserves typed numbers while exposing inferred body links', () => {
  const harness = createPersistentRichTextHarness({ richTextClearsQuotePrefix: true });
  const sheet = harness.createSheet('ads', []);
  writeQuotePrefixedRows(sheet, [['campaign.id', 'https://example.test/path', 7]]);

  assert.equal(simulatedDownloadedLinks(sheet.ensureCell(1, 1)).length, 1);
  assert.equal(simulatedDownloadedLinks(sheet.ensureCell(1, 2)).length, 1);
  assert.equal(sheet.ensureCell(1, 3).value, 7);
  assert.equal(typeof sheet.ensureCell(1, 3).value, 'number');
});
