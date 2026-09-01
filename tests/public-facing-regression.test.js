'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadExporter, createPersistentRichTextHarness } = require('./load-exporter');

function completedState() {
  return {
    status: 'COMPLETE',
    accountId: '123-456-7890',
    accountName: 'Example Advertiser',
    accountCurrencyCode: 'USD',
    accountTimeZone: 'America/New_York',
    ranges: {
      aggregate: { start: '2026-05-30', end: '2026-08-27' },
      weekly: { start: '2026-05-30', end: '2026-08-27' },
      change: { start: '2026-07-31', end: '2026-08-27' },
    },
    tabs: {},
  };
}

function zeroConversionCampaign() {
  return {
    'campaign.id': '123',
    'campaign.name': 'Zero Conversion Campaign',
    'campaign.advertising_channel_type': 'SEARCH',
    'campaign_budget.has_recommended_budget': false,
    'metrics.impressions': 100,
    'metrics.clicks': 10,
    'metrics.cost_micros': 1000000,
    'metrics.conversions': 0,
    'metrics.conversions_value': 0,
  };
}

test('public summary uses Google Ads Conversions terminology everywhere', () => {
  const { api } = loadExporter();
  const rows = api.buildStartHereRows(
    api.buildStartHereModel(
      completedState(),
      [zeroConversionCampaign()],
      api.getManifestDefinition(),
    ),
  ).map((row) => Array.from(row));
  const values = rows.flat().map(String);

  for (const label of ['Conversions', 'Conversion rate', 'Cost / conversion', 'Conversion value']) {
    assert.equal(values.includes(label), true, `missing Google Ads label: ${label}`);
  }
  assert.equal(values.includes('Spend with zero conversions'), true);
  assert.equal(
    values.includes('Campaign has nonzero cost and zero conversions in the aggregate range.'),
    true,
  );
  assert.equal(values.some((value) => /primary conversions?/i.test(value)), false);
});

test('public summary distinguishes native use from a sanitized XLSX download', () => {
  const { api } = loadExporter();
  const rows = api.buildStartHereRows(
    api.buildStartHereModel(completedState(), [], api.getManifestDefinition()),
  ).map((row) => Array.from(row));
  const values = rows.flat().map(String);

  assert.equal(values.includes('Deliverable: This Google Sheet'), true);
  assert.equal(
    values.some((value) => /XLSX.*sanitiz|sanitiz.*XLSX/i.test(value)),
    true,
  );
});

test('START_HERE normalizes stale styles and keeps every light-background cell readable', () => {
  const harness = createPersistentRichTextHarness();
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const rows = api.buildStartHereRows(
    api.buildStartHereModel(
      completedState(),
      [zeroConversionCampaign()],
      api.getManifestDefinition(),
    ),
  ).map((row) => Array.from(row));
  const sheet = harness.createSheet('START_HERE', rows);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  // Reproduce a resumable-sheet formatting remnant: white text survives on
  // cells that later become light-blue or white body cells.
  sheet.getRange(1, 1, lastRow, lastColumn)
    .setBackground('#D9EAF7')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold');

  api.formatStartHereSheet(sheet, 'USD');

  const reviewSectionRow = rows.findIndex((row) => row[0] === 'REVIEW FIRST') + 1;
  const reviewHeaderRow = reviewSectionRow + 1;
  for (let column = 1; column <= 8; column += 1) {
    const cell = sheet.ensureCell(reviewHeaderRow, column);
    assert.equal(cell.background, '#D9EAF7');
    assert.equal(cell.fontColor, '#1F1F1F');
  }
  for (let row = 1; row <= lastRow; row += 1) {
    for (let column = 1; column <= lastColumn; column += 1) {
      const cell = sheet.ensureCell(row, column);
      if (cell.background === '#FFFFFF' || cell.background === '#D9EAF7') {
        assert.notEqual(cell.fontColor, '#FFFFFF', `unreadable light cell at R${row}C${column}`);
      }
    }
  }
});

test('default report layout freezes only one header row and no raw-data columns', () => {
  const harness = createPersistentRichTextHarness();
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const campaignJob = api.getManifestDefinition().find((job) => job.tab === 'campaign');
  const report = harness.createSheet('campaign', [
    Array.from(api.headersForJob(campaignJob)),
    Array.from(api.headersForJob(campaignJob), () => ''),
  ]);
  report.ensureCell(2, 1).value = '1234567890';

  api.formatReportSheet(report, campaignJob, 'USD');
  assert.equal(report.frozenRows, 1);
  assert.equal(report.frozenColumns, 0);
});

test('dictionary layouts freeze one header row and no columns', () => {
  const harness = createPersistentRichTextHarness();
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  for (const tab of ['_data_dictionary', '_field_dictionary']) {
    const dictionaryJob = api.getManifestDefinition().find((job) => job.tab === tab);
    const dictionary = harness.createSheet(tab, [
      Array.from(api.headersForJob(dictionaryJob)),
      Array.from(api.headersForJob(dictionaryJob), () => 'value'),
    ]);
    api.formatReportSheet(dictionary, dictionaryJob, 'USD');
    assert.equal(dictionary.frozenRows, 1, `${tab} frozen rows`);
    assert.equal(dictionary.frozenColumns, 0, `${tab} frozen columns`);
  }
});

test('_export_info freezes one row and no columns', () => {
  const harness = createPersistentRichTextHarness();
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const info = harness.createSheet('_export_info', Array.from({ length: 9 }, (_, row) => (
    Array.from({ length: 8 }, (_, column) => (row === 0 || row === 7 ? `header-${column}` : 'value'))
  )));
  api.formatExportInfoSheet(info);
  assert.equal(info.frozenRows, 1);
  assert.equal(info.frozenColumns, 0);
});

test('audience contract keeps reliable performance and signal tabs but removes user-list inventory', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const tabs = manifest.map((job) => job.tab);
  const probes = api.diagnosticProbes({ start: '2026-08-21', end: '2026-08-27' })
    .map((probe) => probe.name);

  assert.equal(api.OUTPUT_SCHEMA_VERSION, 9);
  assert.equal(api.RUNTIME_CONTRACT_VERSION, 10);
  assert.equal(manifest.length, 39);
  assert.equal(api.preferredTabOrder(manifest).length, 41);
  assert.equal(api.buildFieldDictionaryRows(manifest).length - 1, 914);
  assert.equal(api.buildDataDictionaryRows(manifest).length - 1, 41);
  assert.equal(tabs.includes('user_lists'), false);
  assert.equal(tabs.includes('user_list_performance'), true);
  assert.equal(tabs.includes('pmax_audience_signals'), true);
  assert.equal(api.headersForJob(manifest.find((job) => job.tab === 'user_list_performance')).length, 22);
  assert.equal(api.headersForJob(manifest.find((job) => job.tab === 'pmax_audience_signals')).length, 17);
  assert.equal(api.supportedJobKinds().includes('user_lists'), false);
  assert.equal(probes.includes('user_list_inventory'), false);
  assert.equal(probes.includes('campaign_audience'), true);
  assert.equal(probes.includes('ad_group_audience'), true);

  const directoryTabs = api.buildStartHereModel(
    completedState(),
    [],
    manifest,
  ).directory.map((item) => item.tab);
  assert.equal(directoryTabs.includes('user_lists'), false);
  assert.equal(directoryTabs.includes('user_list_performance'), true);
  assert.equal(directoryTabs.includes('pmax_audience_signals'), true);

});

test('retained audience outputs disclose that they are not complete inventory', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const dictionaryRows = api.buildDataDictionaryRows(manifest).map((row) => Array.from(row));
  const dictionaryHeaders = dictionaryRows[0];
  const performanceRow = dictionaryRows.find((row) => row[0] === 'user_list_performance');
  const signalsRow = dictionaryRows.find((row) => row[0] === 'pmax_audience_signals');
  const limitation = String(
    performanceRow[dictionaryHeaders.indexOf('google_side_limitations')],
  );
  const sensitivityIndex = dictionaryHeaders.indexOf('sensitive_data');
  assert.match(limitation, /returned performance coverage/i);
  assert.match(limitation, /not current or complete audience inventory/i);
  assert.match(limitation, /do not infer/i);
  assert.match(String(performanceRow[sensitivityIndex]), /resource identifiers and performance/i);
  assert.doesNotMatch(String(performanceRow[sensitivityIndex]), /signals/i);
  assert.match(String(signalsRow[sensitivityIndex]), /audience and search-theme signals/i);
  assert.doesNotMatch(String(signalsRow[sensitivityIndex]), /performance (?:data|metrics?)/i);
});
