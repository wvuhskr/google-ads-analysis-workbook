'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadExporter, createPersistentRichTextHarness } = require('./load-exporter');

function state(status = 'COMPLETE') {
  return {
    status,
    accountId: '000-000-0000',
    accountName: 'Synthetic Demo Account',
    accountCurrencyCode: 'USD',
    ranges: { aggregate: { start: '2026-05-29', end: '2026-08-26' } },
    tabs: {},
  };
}

test('START_HERE totals campaign metrics and groups values by channel', () => {
  const { api } = loadExporter();
  const campaigns = [
    {
      'campaign.id': '1', 'campaign.name': 'Search One',
      'campaign.advertising_channel_type': 'SEARCH',
      'campaign_budget.amount_micros': 50000000,
      'campaign_budget.has_recommended_budget': false,
      'campaign_budget.recommended_budget_amount_micros': '',
      'metrics.impressions': 1000, 'metrics.clicks': 100, 'metrics.interactions': 80,
      'metrics.cost_micros': 250000000, 'metrics.conversions': 10,
      'metrics.conversions_value': 1000,
    },
    {
      'campaign.id': '2', 'campaign.name': 'PMax One',
      'campaign.advertising_channel_type': 'PERFORMANCE_MAX',
      'campaign_budget.amount_micros': 70000000,
      'campaign_budget.has_recommended_budget': false,
      'campaign_budget.recommended_budget_amount_micros': '',
      'metrics.impressions': 3000, 'metrics.clicks': 300, 'metrics.interactions': 600,
      'metrics.cost_micros': 750000000, 'metrics.conversions': 30,
      'metrics.conversions_value': 3000,
    },
  ];
  const model = api.buildStartHereModel(state(), campaigns, api.getManifestDefinition());
  assert.deepEqual(JSON.parse(JSON.stringify(model.kpis)), {
    cost: 1000, impressions: 4000, clicks: 400, ctr: 0.1,
    conversions: 40, conversionRate: 40 / 680, costPerConversion: 25,
    conversionValue: 4000,
  });
  assert.deepEqual(model.channels.map((row) => row.channel), ['PERFORMANCE_MAX', 'SEARCH']);
  assert.equal(model.channels.reduce((sum, row) => sum + row.cost, 0), 1000);
  assert.equal(model.channels.find((row) => row.channel === 'PERFORMANCE_MAX').conversionRate, 0.05);
  assert.equal(model.channels.find((row) => row.channel === 'SEARCH').conversionRate, 0.125);
});

test('START_HERE reads typed campaign values instead of formatted display strings', () => {
  const { api } = loadExporter();
  const typedRows = [[
    'campaign.id', 'campaign.name', 'campaign.advertising_channel_type',
    'campaign_budget.amount_micros', 'campaign_budget.has_recommended_budget',
    'campaign_budget.recommended_budget_amount_micros',
    'metrics.impressions', 'metrics.clicks', 'metrics.interactions', 'metrics.cost_micros',
    'metrics.conversions', 'metrics.conversions_value',
  ], [
    '0000000000000001', 'Northstar Search', 'SEARCH',
    75000000, true, 150000000,
    1200, 120, 150, 34567890, 4.5, 123.45,
  ]];
  const displayRows = [typedRows[0], [
    '0000000000000001', 'Northstar Search', 'SEARCH',
    '75,000,000', 'TRUE', '150,000,000',
    '1,200', '120', '150', '34,567,890', '4.50', '123.45',
  ]];
  const campaignSheet = {
    getLastRow() { return 2; },
    getLastColumn() { return typedRows[0].length; },
    getRange() {
      return {
        getValues() { return typedRows.map((row) => row.slice()); },
        getDisplayValues() { return displayRows.map((row) => row.slice()); },
      };
    },
  };
  const workbook = {
    getSheetByName(name) { return name === 'campaign' ? campaignSheet : null; },
  };

  const campaigns = api.readSheetObjectsRuntime(workbook, 'campaign');
  const model = api.buildStartHereModel(state(), campaigns, api.getManifestDefinition());

  assert.equal(typeof campaigns[0]['metrics.cost_micros'], 'number');
  assert.deepEqual(JSON.parse(JSON.stringify(model.kpis)), {
    cost: 34.56789,
    impressions: 1200,
    clicks: 120,
    ctr: 120 / 1200,
    conversions: 4.5,
    conversionRate: 4.5 / 150,
    costPerConversion: 34.56789 / 4.5,
    conversionValue: 123.45,
  });
  const recommendation = model.reviews.find((row) => row.kind === 'BUDGET_RECOMMENDATION');
  assert.equal(recommendation.campaignId, '0000000000000001');
  assert.equal(recommendation.currentBudget, 75);
  assert.equal(recommendation.recommendedBudget, 150);
});

test('START_HERE leaves ratios blank when their denominator is zero', () => {
  const { api } = loadExporter();
  const model = api.buildStartHereModel(state(), [{
    'campaign.id': '1', 'campaign.name': 'Empty',
    'campaign.advertising_channel_type': 'SEARCH',
    'metrics.impressions': 0, 'metrics.clicks': 0, 'metrics.interactions': 0,
    'metrics.cost_micros': 0, 'metrics.conversions': 0,
    'metrics.conversions_value': 0,
  }], api.getManifestDefinition());
  assert.equal(model.kpis.ctr, '');
  assert.equal(model.kpis.conversionRate, '');
  assert.equal(model.kpis.costPerConversion, '');
});

test('START_HERE orders limitations before campaign review facts', () => {
  const { api } = loadExporter();
  const current = state('COMPLETE_WITH_LIMITATIONS');
  current.tabs.user_list_performance = {
    status: 'LIMITED',
    limitation: 'AD_GROUP audience scope is unavailable',
  };
  const campaigns = [
    {
      'campaign.id': '20', 'campaign.name': 'Budget Candidate',
      'campaign.advertising_channel_type': 'SEARCH',
      'campaign_budget.amount_micros': 100000000,
      'campaign_budget.has_recommended_budget': true,
      'campaign_budget.recommended_budget_amount_micros': 200000000,
      'metrics.cost_micros': 50000000, 'metrics.conversions': 2,
    },
    {
      'campaign.id': '10', 'campaign.name': 'Zero Conversion',
      'campaign.advertising_channel_type': 'SEARCH',
      'campaign_budget.has_recommended_budget': false,
      'metrics.cost_micros': 90000000, 'metrics.conversions': 0,
    },
  ];
  const model = api.buildStartHereModel(current, campaigns, api.getManifestDefinition());
  assert.deepEqual(Array.from(model.reviews, (row) => row.kind), [
    'LIMITED_COVERAGE', 'ZERO_CONVERSION_SPEND', 'BUDGET_RECOMMENDATION',
  ]);
  assert.match(model.reviews[0].detail, /AD_GROUP audience scope/i);
});

test('START_HERE caps campaign review categories at 25 and reports omissions', () => {
  const { api } = loadExporter();
  const campaigns = Array.from({ length: 30 }, (_, index) => ({
    'campaign.id': String(index + 1),
    'campaign.name': `Campaign ${index + 1}`,
    'campaign.advertising_channel_type': 'SEARCH',
    'campaign_budget.has_recommended_budget': false,
    'metrics.cost_micros': (index + 1) * 1000000,
    'metrics.conversions': 0,
  }));
  const model = api.buildStartHereModel(state(), campaigns, api.getManifestDefinition());
  const zeroRows = model.reviews.filter((row) => row.kind === 'ZERO_CONVERSION_SPEND');
  assert.equal(zeroRows.length, 25);
  assert.equal(model.reviews.some((row) => row.kind === 'OMITTED_COUNT' && row.count === 5), true);
  assert.equal(
    model.reviews.some((row) => (
      row.kind === 'OMITTED_COUNT' &&
      row.detail === '5 additional campaigns with spend and zero conversions omitted; review campaign for the complete list.'
    )),
    true,
  );
});

test('START_HERE rows are rectangular values with no formulas', () => {
  const { api } = loadExporter();
  const model = api.buildStartHereModel(state(), [], api.getManifestDefinition());
  const rows = api.buildStartHereRows(model).map((row) => Array.from(row));
  assert.equal(rows.every((row) => row.length === 8), true);
  assert.equal(rows.flat().some((value) => typeof value === 'string' && value.startsWith('=')), false);
  assert.equal(rows.flat().includes('No deterministic review flags were produced'), true);
});

test('START_HERE uses Google Ads conversion labels throughout', () => {
  const { api } = loadExporter();
  const current = state('COMPLETE_WITH_LIMITATIONS');
  current.tabs.user_list_performance = {
    status: 'LIMITED',
    rows: 2,
    limitation: 'AD_GROUP audience scope is unavailable',
  };
  const campaigns = [{
    'campaign.id': '123',
    'campaign.name': 'Zero Conversions',
    'campaign.advertising_channel_type': 'SEARCH',
    'metrics.cost_micros': 1000000,
    'metrics.conversions': 0,
  }];
  const rows = api.buildStartHereRows(
    api.buildStartHereModel(current, campaigns, api.getManifestDefinition()),
  ).map((row) => Array.from(row));
  const values = rows.flat().map(String);

  for (const label of [
    'Conversions',
    'Conversion rate',
    'Cost / conversion',
    'Conversion value',
  ]) {
    assert.equal(values.includes(label), true, `missing ${label}`);
  }
  assert.equal(values.some((value) => /primary conversions?/i.test(value)), false);
  assert.equal(values.includes('Data status: Complete with limitations'), true);
  assert.equal(values.includes('Limited data coverage'), true);
  assert.equal(values.includes('Spend with zero conversions'), true);
  assert.equal(
    values.includes('Campaign has nonzero cost and zero conversions in the aggregate range.'),
    true,
  );
  assert.equal(
    values.some((value) => /AD_GROUP audience scope is unavailable/i.test(value)),
    true,
  );
});

test('START_HERE identifies the native Google Sheet and the optional sanitized XLSX path', () => {
  const { api } = loadExporter();
  const current = state();
  const model = api.buildStartHereModel(current, [], api.getManifestDefinition());
  const rows = api.buildStartHereRows(model).map((row) => Array.from(row));
  const values = rows.flat().map(String);

  assert.equal(values.includes('Deliverable: This Google Sheet'), true);
  assert.equal(values.includes('Data status: Complete'), true);
  assert.equal(
    values.includes('XLSX downloads: sanitize with the bundled local tool before upload or sharing.'),
    true,
  );
  assert.equal(values.some((value) => /distribution|automatic.*XLSX|XLSX.*automatic/i.test(value)), false);
});

test('START_HERE directory groups every workbook tab with status and row counts', () => {
  const { api } = loadExporter();
  const current = state();
  current.tabs.campaign = { status: 'OK', rows: 8 };
  current.tabs.geo_proximity_targets = { status: 'OK', rows: 0 };
  const model = api.buildStartHereModel(current, [], api.getManifestDefinition());
  assert.equal(model.directory.length, 40);
  assert.equal(model.directory[0].group, 'performance');
  assert.equal(model.directory.some((row) => row.tab === '_export_info' && row.group === 'governance'), true);
  assert.equal(model.directory.some((row) => /^https?:\/\//i.test(row.tab)), false);
  assert.deepEqual(
    JSON.parse(JSON.stringify(model.directory.find((row) => row.tab === 'campaign'))),
    {
      group: 'performance',
      tab: 'campaign',
      status: 'OK',
      rows: 8,
      purpose: 'Campaign performance',
      grain: 'campaign',
      dateRange: 'Frozen 90 complete days ending yesterday',
      recommendedUse: 'Use to analyze account results and performance drivers.',
    },
  );

  const rows = api.buildStartHereRows(model).map((row) => Array.from(row));
  const directoryHeader = rows.find((row) => row[0] === 'Group' && row[1] === 'Tab');
  assert.deepEqual(directoryHeader, [
    'Group', 'Tab', 'Status', 'Rows', 'Purpose', 'Row grain', 'Date range', 'Recommended use',
  ]);
  const empty = rows.find((row) => row[1] === 'geo_proximity_targets');
  assert.equal(empty[2], 'OK — no matching records');
  assert.equal(empty[3], 0);
});

test('START_HERE gives metadata, KPI labels, and review currency enough space', () => {
  const harness = createPersistentRichTextHarness();
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const rows = api.buildStartHereRows(
    api.buildStartHereModel(state('COMPLETE_WITH_LIMITATIONS'), [], api.getManifestDefinition()),
  );
  const sheet = harness.createSheet('START_HERE', rows);

  api.formatStartHereSheet(sheet, 'USD');

  assert.ok(sheet.rowHeights[3] >= 72);
  assert.ok(sheet.rowHeights[7] >= 54);
  assert.ok(sheet.columnWidths[6] >= 210);
  assert.ok(sheet.columnWidths[8] >= 360);
});

test('START_HERE links directory and review evidence without formulas', () => {
  const harness = createPersistentRichTextHarness({ richTextClearsQuotePrefix: true });
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const current = state();
  current.tabs.campaign = { status: 'OK', rows: 1 };
  const campaigns = [{
    'campaign.id': '0000000000000042',
    'campaign.name': 'Northstar Search',
    'campaign.advertising_channel_type': 'SEARCH',
    'campaign_budget.amount_micros': 100000000,
    'campaign_budget.has_recommended_budget': true,
    'campaign_budget.recommended_budget_amount_micros': 200000000,
    'metrics.cost_micros': 50000000,
    'metrics.conversions': 2,
  }];
  const rows = api.buildStartHereRows(
    api.buildStartHereModel(current, campaigns, api.getManifestDefinition()),
  ).map((row) => Array.from(row));
  const start = harness.createSheet('START_HERE', rows);
  const campaign = harness.createSheet('campaign', [
    ['campaign.id', 'campaign.name'],
    ['0000000000000042', 'Northstar Search'],
  ]);
  start.getSheetId = () => 1;
  campaign.getSheetId = () => 77;
  const spreadsheet = {
    getId() { return 'sheet-123'; },
    getSheetByName(name) {
      return name === 'START_HERE' ? start : (name === 'campaign' ? campaign : null);
    },
  };

  api.formatStartHereSheet(start, 'USD');
  assert.equal(api.applyStartHereNavigationLinks(spreadsheet, start), 3);

  const directoryRow = rows.findIndex((row) => row[1] === 'campaign') + 1;
  const reviewRow = rows.findIndex((row) => row[1] === 'Google budget recommendation') + 1;
  assert.match(
    start.ensureCell(directoryRow, 2).richText.getLinkUrl(),
    /\/edit#gid=77&range=A1$/,
  );
  assert.match(
    start.ensureCell(reviewRow, 3).richText.getLinkUrl(),
    /\/edit#gid=77&range=A1$/,
  );
  assert.equal(start.ensureCell(reviewRow, 4).richText.getLinkUrl(), null);
  assert.equal(start.ensureCell(reviewRow, 4).numberFormat, '@');
  assert.equal(start.ensureCell(reviewRow, 4).value, '0000000000000042');
  assert.match(start.ensureCell(reviewRow, 5).richText.getLinkUrl(), /\/edit#gid=77&range=A2$/);
  assert.equal(start.ensureCell(reviewRow, 5).numberFormat, '@');
  assert.equal(start.ensureCell(reviewRow, 5).value, 'Northstar Search');
  assert.equal(rows.flat().some((value) => String(value).startsWith('=')), false);
});
