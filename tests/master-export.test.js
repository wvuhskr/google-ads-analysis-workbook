'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadExporter, createPersistentRichTextHarness } = require('./load-exporter');

function displayedSheetValue(value) {
  return typeof value === 'string' && value.startsWith("'") ? value.slice(1) : value;
}

class FakeRange {
  constructor(sheet, row, column, rowCount, columnCount) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.rowCount = rowCount;
    this.columnCount = columnCount;
  }

  setValues(values) {
    assert.equal(values.length, this.rowCount);
    for (const valuesRow of values) assert.equal(valuesRow.length, this.columnCount);
    for (let r = 0; r < values.length; r += 1) {
      const targetRow = this.row - 1 + r;
      if (!this.sheet.data[targetRow]) this.sheet.data[targetRow] = [];
      for (let c = 0; c < values[r].length; c += 1) {
        this.sheet.data[targetRow][this.column - 1 + c] = displayedSheetValue(values[r][c]);
      }
    }
    return this;
  }

  getDisplayValues() {
    const values = [];
    for (let r = 0; r < this.rowCount; r += 1) {
      const source = this.sheet.data[this.row - 1 + r] || [];
      const row = [];
      for (let c = 0; c < this.columnCount; c += 1) {
        const value = source[this.column - 1 + c];
        row.push(value === null || value === undefined ? '' : String(value));
      }
      values.push(row);
    }
    return values;
  }

  getValues() {
    const values = [];
    for (let r = 0; r < this.rowCount; r += 1) {
      const source = this.sheet.data[this.row - 1 + r] || [];
      const row = [];
      for (let c = 0; c < this.columnCount; c += 1) {
        const value = source[this.column - 1 + c];
        row.push(value === null || value === undefined ? '' : value);
      }
      values.push(row);
    }
    return values;
  }

  setRichTextValues(values) {
    this.record('richTextValues', values);
    return this.setValues(values.map((row) => row.map((value) => value.text)));
  }

  clearContent() {
    for (let r = 0; r < this.rowCount; r += 1) {
      for (let c = 0; c < this.columnCount; c += 1) {
        const target = this.sheet.data[this.row - 1 + r];
        if (target) target[this.column - 1 + c] = '';
      }
    }
    return this;
  }

  record(type, value) {
    this.sheet.operations.push({
      type,
      value,
      row: this.row,
      column: this.column,
      rowCount: this.rowCount,
      columnCount: this.columnCount,
    });
    return this;
  }

  setBackground(value) { return this.record('background', value); }
  setFontColor(value) { return this.record('fontColor', value); }
  setFontWeight(value) { return this.record('fontWeight', value); }
  setHorizontalAlignment(value) { return this.record('horizontalAlignment', value); }
  setNumberFormat(value) { return this.record('numberFormat', value); }
  setShowHyperlink(value) { return this.record('showHyperlink', value); }
  setVerticalAlignment(value) { return this.record('verticalAlignment', value); }
  setWrap(value) { return this.record('wrap', value); }

  createFilter() {
    this.sheet.filter = {
      row: this.row,
      column: this.column,
      rowCount: this.rowCount,
      columnCount: this.columnCount,
      remove: () => { this.sheet.filter = null; },
    };
    return this.sheet.filter;
  }
}

class FakeSheet {
  constructor(parent, name, data = [], maxRows = 10, maxColumns = 5) {
    this.parent = parent;
    this.name = name;
    this.data = data.map((row) => row.slice());
    this.maxRows = maxRows;
    this.maxColumns = maxColumns;
    this.columnWidths = {};
    this.filter = null;
    this.frozenColumns = 0;
    this.frozenRows = 0;
    this.hiddenColumns = new Set();
    this.operations = [];
    this.rowHeights = {};
    this.tabColor = '';
  }

  getName() { return this.name; }
  getParent() { return this.parent; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxColumns; }
  getLastRow() {
    let last = 0;
    this.data.forEach((row, index) => {
      if (row.some((value) => value !== '' && value !== null && value !== undefined)) last = index + 1;
    });
    return last;
  }
  getLastColumn() {
    let last = 0;
    this.data.forEach((row) => row.forEach((value, index) => {
      if (value !== '' && value !== null && value !== undefined) last = Math.max(last, index + 1);
    }));
    return last;
  }
  getRange(row, column, rowCount, columnCount) {
    assert.ok(row >= 1 && column >= 1 && rowCount >= 1 && columnCount >= 1);
    assert.ok(row + rowCount - 1 <= this.maxRows, `range exceeds ${this.name} row bounds`);
    assert.ok(column + columnCount - 1 <= this.maxColumns, `range exceeds ${this.name} column bounds`);
    return new FakeRange(this, row, column, rowCount, columnCount);
  }
  insertRowsAfter(_after, count) { this.maxRows += count; }
  insertColumnsAfter(_after, count) { this.maxColumns += count; }
  deleteRows(start, count) {
    const remainingRows = this.maxRows - count;
    if (this.frozenRows > 0 && remainingRows <= this.frozenRows) {
      throw new Error('Sorry, it is not possible to delete all non-frozen rows.');
    }
    this.data.splice(start - 1, count);
    this.maxRows -= count;
    return this;
  }
  deleteColumns(start, count) {
    const remainingColumns = this.maxColumns - count;
    if (this.frozenColumns > 0 && remainingColumns <= this.frozenColumns) {
      throw new Error('Sorry, it is not possible to delete all non-frozen columns.');
    }
    for (const row of this.data) row.splice(start - 1, count);
    this.maxColumns -= count;
    return this;
  }
  clearContents() { this.data = []; return this; }
  getFilter() { return this.filter; }
  hideColumns(column) { this.hiddenColumns.add(column); return this; }
  setColumnWidth(column, width) { this.columnWidths[column] = width; return this; }
  setColumnWidths(startColumn, count, width) {
    for (let column = startColumn; column < startColumn + count; column += 1) {
      this.columnWidths[column] = width;
    }
    return this;
  }
  getFrozenColumns() { return this.frozenColumns; }
  getFrozenRows() { return this.frozenRows; }
  setFrozenColumns(count) { this.frozenColumns = count; return this; }
  setFrozenRows(count) { this.frozenRows = count; return this; }
  setName(name) { this.parent.rename(this, name); return this; }
  setRowHeight(row, height) { this.rowHeights[row] = height; return this; }
  setTabColor(value) { this.tabColor = value; return this; }
}

class FakeSpreadsheet {
  constructor(specs = []) {
    this.sheets = [];
    this.activeSheet = null;
    this.failRename = null;
    for (const spec of specs) {
      this.sheets.push(new FakeSheet(this, spec.name, spec.data, spec.maxRows, spec.maxColumns));
    }
    this.activeSheet = this.sheets[0] || null;
  }
  getSheets() { return this.sheets.slice(); }
  getSheetByName(name) { return this.sheets.find((sheet) => sheet.name === name) || null; }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`duplicate sheet ${name}`);
    const sheet = new FakeSheet(this, name);
    this.sheets.push(sheet);
    return sheet;
  }
  deleteSheet(sheet) { this.sheets = this.sheets.filter((candidate) => candidate !== sheet); }
  getActiveSheet() { return this.activeSheet; }
  moveActiveSheet(position) {
    const index = this.sheets.indexOf(this.activeSheet);
    assert.notEqual(index, -1);
    this.sheets.splice(index, 1);
    this.sheets.splice(position - 1, 0, this.activeSheet);
  }
  rename(sheet, targetName) {
    if (this.failRename && this.failRename.from === sheet.name && this.failRename.to === targetName) {
      throw new Error('injected rename failure');
    }
    const collision = this.getSheetByName(targetName);
    if (collision && collision !== sheet) throw new Error(`duplicate sheet ${targetName}`);
    sheet.name = targetName;
  }
  setActiveSheet(sheet) { this.activeSheet = sheet; return sheet; }
}

test('loads the revised exporter globals for local verification', () => {
  const { api } = loadExporter();

  assert.equal(api.VERSION, 'v1.0.0');
});

test('detects manager accounts from customer.manager instead of global availability', () => {
  function loadWithManagerFlag(value) {
    return loadExporter({
      AdsApp: {
        report(query) {
          assert.match(query, /SELECT\s+customer\.manager\s+FROM customer/);
          return {
            rows() {
              let read = false;
              return {
                hasNext() { return !read; },
                next() { read = true; return { 'customer.manager': value }; },
              };
            },
          };
        },
      },
    });
  }

  assert.equal(loadWithManagerFlag(false).api.assertAdvertiserAccount(), true);
  assert.throws(
    () => loadWithManagerFlag('true').api.assertAdvertiserAccount(),
    /individual advertiser account/,
  );
  assert.doesNotMatch(loadWithManagerFlag(false).source, /typeof AdsManagerApp/);
});

test('uses one exact last-90-complete-day range for aggregate and weekly reporting', () => {
  const { api } = loadExporter();

  assert.deepEqual(
    JSON.parse(JSON.stringify(api.buildFrozenRanges('2026-01-01'))),
    {
      aggregate: { start: '2025-10-03', end: '2025-12-31' },
      weekly: { start: '2025-10-03', end: '2025-12-31' },
      change: { start: '2025-12-04', end: '2025-12-31' },
    },
  );
});

test('keeps partial first and last weeks inside the exact 90-day reporting window', () => {
  const { api } = loadExporter();
  const ranges = api.buildFrozenRanges('2026-08-28');

  assert.deepEqual(JSON.parse(JSON.stringify(ranges.aggregate)), {
    start: '2026-05-30',
    end: '2026-08-27',
  });
  assert.deepEqual(
    JSON.parse(JSON.stringify(ranges.weekly)),
    JSON.parse(JSON.stringify(ranges.aggregate)),
  );
});

test('survey eligibility removes inactive zero-data entities but keeps enabled or historically active entities', () => {
  const { api } = loadExporter();
  const zeroMetrics = {
    'metrics.impressions': '0',
    'metrics.clicks': '0',
    'metrics.cost_micros': '0',
    'metrics.conversions': '0',
    'metrics.conversions_value': '0',
    'metrics.all_conversions': '0',
    'metrics.all_conversions_value': '0',
  };

  assert.equal(api.shouldIncludeSurveyEntity(
    { ...zeroMetrics, 'campaign.status': 'ENABLED' },
    ['campaign.status'],
  ), true);
  assert.equal(api.shouldIncludeSurveyEntity(
    { ...zeroMetrics, 'campaign.status': 'PAUSED' },
    ['campaign.status'],
  ), false);
  assert.equal(api.shouldIncludeSurveyEntity(
    { ...zeroMetrics, 'campaign.status': 'REMOVED' },
    ['campaign.status'],
  ), false);
  assert.equal(api.shouldIncludeSurveyEntity(
    { ...zeroMetrics, 'campaign.status': 'PAUSED', 'metrics.impressions': '1' },
    ['campaign.status'],
  ), true);
  assert.equal(api.shouldIncludeSurveyEntity(
    { ...zeroMetrics, 'campaign.status': 'REMOVED', 'metrics.conversions_value': '-2.5' },
    ['campaign.status'],
  ), true);
  assert.equal(api.shouldIncludeSurveyEntity(
    { ...zeroMetrics, 'ad_group_ad_asset_view.enabled': false },
    ['ad_group_ad_asset_view.enabled'],
  ), false);
});

test('campaign eligibility freezes current non-inactive campaigns plus 90-day active historical campaigns', () => {
  const { api } = loadExporter();
  const range = { start: '2026-05-27', end: '2026-08-24' };
  const queries = api.buildCampaignEligibilityQueries(range);

  assert.match(queries.current, /campaign\.status NOT IN \('PAUSED', 'REMOVED'\)/);
  assert.match(queries.activity, /segments\.date BETWEEN '2026-05-27' AND '2026-08-24'/);
  assert.match(queries.activity, /metrics\.impressions/);
  assert.deepEqual(
    Array.from(api.collectEligibleCampaignIds(
      [
        { 'campaign.id': '2' },
        { 'campaign.id': '101' },
      ],
      [
        { 'campaign.id': '300', 'metrics.impressions': '0', 'metrics.conversions': '0' },
        { 'campaign.id': '400', 'metrics.impressions': '7' },
        { 'campaign.id': '9007199254740993', 'metrics.conversions_value': '12.5' },
      ],
    )),
    ['2', '101', '400', '9007199254740993'],
  );
});

test('validates unique manifest job IDs and final tab names', () => {
  const { api } = loadExporter();

  assert.equal(api.validateManifest([
    { id: 'campaign', tab: 'campaign' },
    { id: 'ads', tab: 'ads' },
  ]), true);
  assert.throws(
    () => api.validateManifest([
      { id: 'duplicate', tab: 'one' },
      { id: 'duplicate', tab: 'two' },
    ]),
    /duplicate job id/i,
  );
  assert.throws(
    () => api.validateManifest([
      { id: 'one', tab: 'duplicate' },
      { id: 'two', tab: 'duplicate' },
    ]),
    /duplicate final tab/i,
  );
});

test('every production tab has unique output headers', () => {
  const { api } = loadExporter();

  for (const job of api.getManifestDefinition()) {
    const headers = Array.from(api.headersForJob(job));
    assert.equal(
      new Set(headers).size,
      headers.length,
      `${job.tab} contains duplicate output headers`,
    );
  }
});

test('accepts only blank or exporter-owned workbook summaries', () => {
  const { api } = loadExporter();

  assert.equal(api.assertWorkbookOwnership([]), 'blank');
  assert.equal(api.assertWorkbookOwnership([
    { name: 'Sheet1', blank: true, marker: '' },
  ]), 'blank');
  assert.equal(api.assertWorkbookOwnership([
    { name: '_export_info', blank: false, marker: api.OWNER_KEY },
    { name: 'campaign', blank: false, marker: '' },
  ]), 'owned');
  assert.equal(api.assertWorkbookOwnership([
    { name: '_export_state', blank: false, marker: api.OWNER_KEY },
    { name: '__gads_export_stage__campaign', blank: false, marker: '' },
  ]), 'owned');
  assert.throws(
    () => api.assertWorkbookOwnership([
      { name: 'Quarterly Board Report', blank: false, marker: '' },
    ]),
    /populated.*not owned/i,
  );
});

test('blank-sheet cleanup rejects cell-empty tabs with embedded objects', () => {
  const { api } = loadExporter();
  const base = {
    getLastRow() { return 0; },
    getLastColumn() { return 0; },
  };
  assert.equal(api.sheetIsSafelyRemovableBlank(base), true);
  for (const method of ['getCharts', 'getImages', 'getDrawings']) {
    assert.equal(
      api.sheetIsSafelyRemovableBlank({ ...base, [method]() { return [{}]; } }),
      false,
      `${method} objects must prevent automatic tab deletion`,
    );
  }
});

test('finalizes the workbook into a deliberate order and removes only captured blank tabs', () => {
  const { api } = loadExporter();
  const manifest = [
    { id: 'campaign', tab: 'campaign' },
    { id: 'search_terms', tab: 'search_terms' },
    { id: 'data_dictionary', tab: '_data_dictionary' },
    { id: 'field_dictionary', tab: '_field_dictionary' },
  ];
  const ss = new FakeSpreadsheet([
    { name: 'Sheet1', data: [] },
    { name: 'search_terms', data: [['query']] },
    { name: '_data_dictionary', data: [['tab']] },
    { name: '_field_dictionary', data: [['tab', 'field']] },
    { name: 'campaign', data: [['campaign.id']] },
    { name: '_export_info', data: [['google-ads-analysis-workbook']] },
    { name: 'START_HERE', data: [['Google Ads Analysis Workbook']] },
  ]);

  api.finalizeWorkbookLayout(ss, manifest, ['Sheet1']);

  assert.deepEqual(ss.getSheets().map((sheet) => sheet.getName()), [
    'START_HERE', '_export_info', '_data_dictionary', '_field_dictionary', 'campaign', 'search_terms',
  ]);
  assert.equal(ss.getActiveSheet().getName(), 'START_HERE');
  assert.equal(ss.getSheetByName('START_HERE').tabColor, '#70AD47');

  const populatedDefault = new FakeSpreadsheet([
    { name: 'Sheet1', data: [['keep me']] },
    { name: '_export_info', data: [['google-ads-analysis-workbook']] },
    { name: '_data_dictionary', data: [['tab']] },
    { name: '_field_dictionary', data: [['tab', 'field']] },
    { name: 'campaign', data: [['campaign.id']] },
    { name: 'search_terms', data: [['query']] },
    { name: 'START_HERE', data: [['Google Ads Analysis Workbook']] },
  ]);
  api.finalizeWorkbookLayout(populatedDefault, manifest, ['Sheet1']);
  assert.equal(populatedDefault.getSheetByName('Sheet1').data[0][0], 'keep me');

  const chartOnlyDefault = new FakeSpreadsheet([
    { name: 'Sheet1', data: [] },
    { name: '_export_info', data: [['google-ads-analysis-workbook']] },
    { name: '_data_dictionary', data: [['tab']] },
    { name: '_field_dictionary', data: [['tab', 'field']] },
    { name: 'campaign', data: [['campaign.id']] },
    { name: 'search_terms', data: [['query']] },
    { name: 'START_HERE', data: [['Google Ads Analysis Workbook']] },
  ]);
  chartOnlyDefault.getSheetByName('Sheet1').getCharts = () => [{ id: 'chart-1' }];
  api.finalizeWorkbookLayout(chartOnlyDefault, manifest, ['Sheet1']);
  assert.notEqual(chartOnlyDefault.getSheetByName('Sheet1'), null);
});

test('writes and formats START_HERE as the first values-only sheet', () => {
  const { api } = loadExporter({ SpreadsheetApp: { flush() {} } });
  const manifest = api.getManifestDefinition();
  const campaignHeaders = [
    'campaign.id', 'campaign.name', 'campaign.advertising_channel_type',
    'metrics.impressions', 'metrics.clicks', 'metrics.cost_micros',
    'metrics.conversions', 'metrics.conversions_value',
  ];
  const ss = new FakeSpreadsheet([{
    name: 'campaign',
    data: [campaignHeaders, ['1', 'Search', 'SEARCH', 100, 10, 10000000, 2, 50]],
    maxRows: 2,
    maxColumns: campaignHeaders.length,
  }]);
  const current = {
    status: 'COMPLETE', accountId: '123-456-7890', accountName: 'Example',
    accountCurrencyCode: 'USD',
    ranges: { aggregate: { start: '2026-05-29', end: '2026-08-26' } },
    tabs: {},
  };

  api.writeStartHereRuntime(ss, current, manifest);

  const sheet = ss.getSheetByName('START_HERE');
  assert.ok(sheet);
  assert.equal(sheet.data.flat().some((value) => String(value).includes('Example')), true);
  assert.equal(sheet.frozenRows, 1);
  assert.equal(sheet.columnWidths[1] >= 180, true);
  assert.equal(sheet.operations.some((entry) => entry.type === 'background'), true);
  assert.equal(sheet.operations.some((entry) => entry.type === 'showHyperlink' && entry.value === false), true);
  assert.equal(sheet.tabColor, '#70AD47');
  assert.equal(
    sheet.data.flat().some((value) => typeof value === 'string' && value.startsWith('=')),
    false,
  );
});

test('formats report sheets for filtering and human-readable analysis', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: 'campaign',
    data: [headers, headers.map(() => 0)],
    maxRows: 2,
    maxColumns: headers.length,
  }]);
  const sheet = ss.getSheetByName('campaign');

  api.formatReportSheet(sheet, job, 'USD');

  assert.deepEqual(
    { row: sheet.filter.row, column: sheet.filter.column, rowCount: sheet.filter.rowCount, columnCount: sheet.filter.columnCount },
    { row: 1, column: 1, rowCount: 2, columnCount: headers.length },
  );
  assert.equal(sheet.frozenRows, 1);
  assert.equal(sheet.frozenColumns, 0);
  assert.equal(Object.keys(sheet.columnWidths).length, headers.length);
  assert.equal(Math.max(...Object.values(sheet.columnWidths)) <= 320, true);
  assert.ok(sheet.rowHeights[1] >= 56, 'wrapped technical headers need readable height');

  function operation(type, header) {
    const column = headers.indexOf(header) + 1;
    return sheet.operations.find((candidate) => candidate.type === type && candidate.column === column);
  }

  assert.equal(operation('background', headers[0]).value, '#1F4E78');
  assert.equal(operation('fontColor', headers[0]).value, '#FFFFFF');
  assert.equal(operation('fontWeight', headers[0]).value, 'bold');
  assert.equal(operation('showHyperlink', headers[0]).value, false);
  assert.equal(operation('wrap', headers[0]).value, true);
  assert.equal(operation('numberFormat', 'metrics.ctr').value, '0.00%');
  assert.equal(operation('numberFormat', 'metrics.interactions').value, '#,##0');
  assert.equal(operation('numberFormat', 'conversion_rate').value, '0.00%');
  assert.equal(operation('numberFormat', 'cost').value, '"USD" #,##0.00');
  assert.equal(operation('numberFormat', 'metrics.cost_micros').value, '#,##0');
  assert.equal(operation('numberFormat', 'campaign.id').value, '@');
  assert.equal(sheet.hiddenColumns.has(headers.indexOf('metrics.cost_micros') + 1), true);
  assert.ok(sheet.columnWidths[headers.indexOf('metrics.clicks') + 1] >= 145);
  assert.ok(sheet.columnWidths[headers.indexOf('campaign.name') + 1] >= 250);
  assert.equal(
    sheet.columnWidths[headers.indexOf('campaign.name') + 1] >
      sheet.columnWidths[headers.indexOf('metrics.clicks') + 1],
    true,
  );
  assert.equal(
    sheet.operations.some((entry) => (
      entry.type === 'wrap' && entry.value === true && entry.row === 2 &&
      entry.column === headers.indexOf('campaign.name') + 1
    )),
    true,
    'campaign names must wrap in the compact campaign report body',
  );
});

test('wraps only the human-review fields in change history', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'change_history');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: 'change_history',
    data: [headers, headers.map(() => 'long value')],
    maxRows: 2,
    maxColumns: headers.length,
  }]);
  const sheet = ss.getSheetByName('change_history');

  api.formatReportSheet(sheet, job, 'USD');

  for (const header of [
    'campaign.name', 'ad_group.name', 'change_resource_name',
    'changed_fields', 'change_event_resource_name',
  ]) {
    const column = headers.indexOf(header) + 1;
    assert.ok(column > 0, `missing ${header}`);
    assert.equal(
      sheet.operations.some((entry) => (
        entry.type === 'wrap' && entry.value === true && entry.row === 2 &&
        entry.column === column
      )),
      true,
      `${header} body values must wrap`,
    );
  }
});

test('marks a successful header-only report without adding a fake data row', () => {
  const harness = createPersistentRichTextHarness();
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const job = api.getManifestDefinition().find(
    (candidate) => candidate.tab === 'geo_proximity_targets',
  );
  const headers = Array.from(api.headersForJob(job));
  const sheet = harness.createSheet(job.tab, [headers]);

  api.formatReportSheet(sheet, job, 'USD');

  assert.equal(sheet.getLastRow(), 1);
  assert.match(sheet.ensureCell(1, 1).note, /completed successfully.*no matching records/i);
});

test('formats field-dictionary ordinals as integers without freezing data columns', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === '_field_dictionary');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: '_field_dictionary',
    data: [headers, headers.map(() => 'value')],
    maxRows: 2,
    maxColumns: headers.length,
  }]);
  const sheet = ss.getSheetByName('_field_dictionary');

  api.formatReportSheet(sheet, job, 'USD');

  const ordinalColumn = headers.indexOf('column_ordinal') + 1;
  const ordinalFormat = sheet.operations.find((operation) => (
    operation.type === 'numberFormat' && operation.column === ordinalColumn
  ));
  assert.equal(sheet.frozenColumns, 0);
  assert.equal(ordinalFormat.value, '#,##0');
});

test('trims unused grid rows and columns without deleting report data', () => {
  const { api } = loadExporter();
  const ss = new FakeSpreadsheet([{
    name: 'report',
    data: [['one', 'two', 'three'], ['a', 'b', 'c']],
    maxRows: 1000,
    maxColumns: 26,
  }]);
  const sheet = ss.getSheetByName('report');

  api.trimSheetGrid(sheet);

  assert.equal(sheet.getMaxRows(), 2);
  assert.equal(sheet.getMaxColumns(), 3);
  assert.deepEqual(sheet.data, [['one', 'two', 'three'], ['a', 'b', 'c']]);
});

test('trim preserves one non-frozen row and column beyond frozen headers', () => {
  const { api } = loadExporter();
  const ss = new FakeSpreadsheet([{
    name: 'stage', data: [['header']], maxRows: 1000, maxColumns: 26,
  }]);
  const sheet = ss.getSheetByName('stage');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  assert.doesNotThrow(() => api.trimSheetGrid(sheet));
  assert.equal(sheet.getLastRow(), 1);
  assert.equal(sheet.getMaxRows(), 2);
  assert.equal(sheet.getMaxColumns(), 2);
  assert.deepEqual(sheet.data[0], ['header']);
});

test('formats a header-only trimmed report without requesting a nonexistent body range', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'geo_proximity_targets');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: job.tab, data: [headers], maxRows: 1, maxColumns: headers.length,
  }]);
  const sheet = ss.getSheetByName(job.tab);

  assert.doesNotThrow(() => api.formatReportSheet(sheet, job, 'USD'));
  assert.equal(sheet.operations.some((operation) => operation.row === 2), false);
});

test('formats export metadata with an independent status-table filter and ISO timestamps', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  state.status = 'COMPLETE_WITH_LIMITATIONS';
  state.tabs.one = { status: 'LIMITED', rows: 2, durationMs: 1500, limitation: 'scope unavailable' };
  const rows = api.buildExportInfoRows(state, [{ id: 'job-one', tab: 'one', required: false }]);
  const ss = new FakeSpreadsheet([{
    name: '_export_info', data: rows, maxRows: rows.length, maxColumns: 8,
  }]);
  const sheet = ss.getSheetByName('_export_info');

  api.formatExportInfoSheet(sheet);

  assert.equal(rows[1][4], 'started_at');
  assert.match(rows[1][5], /^\d{4}-\d{2}-\d{2}T/);
  assert.equal(rows[2][4], 'updated_at');
  assert.match(rows[2][5], /^\d{4}-\d{2}-\d{2}T/);
  assert.deepEqual(
    { row: sheet.filter.row, rowCount: sheet.filter.rowCount, columnCount: sheet.filter.columnCount },
    { row: 8, rowCount: rows.length - 7, columnCount: 8 },
  );
  assert.equal(sheet.frozenRows, 1);
  assert.equal(sheet.frozenColumns, 0);
  assert.equal(sheet.operations.some((entry) => entry.type === 'background' && entry.row === 1), true);
  assert.equal(sheet.operations.some((entry) => entry.type === 'background' && entry.row === 8), true);
});

test('formats overall, limited, and error statuses with explicit palettes', () => {
  const { api } = loadExporter();
  const rows = [
    ['google-ads-analysis-workbook', 'v1.0.0', 'output_schema_version', 6, 'Confidential', '', '', ''],
    ['run_id', 'run-1', 'overall_status', 'COMPLETE_WITH_LIMITATIONS', 'started_at', '2026-08-27T00:00:00.000Z', 'workbook_status', 'READY_WITH_LIMITATIONS'],
    ['account_id', '123-456-7890', 'account_name', 'Example', 'updated_at', '2026-08-27T00:01:00.000Z', 'deliverable_type', 'NATIVE_GOOGLE_SHEET'],
    ['', '', '', '', '', '', 'last_complete_day', '2026-08-26'],
    ['', '', '', '', '', '', 'workbook_url', ['https://docs.google.com', 'spreadsheets', 'd', 'sheet-123', 'edit'].join('/')],
    ['', '', '', '', '', '', 'reporting_window', 'LAST_90_COMPLETE_DAYS'],
    ['', '', '', '', '', '', 'refresh_behavior', 'Run main() again after completion for a fresh export.'],
    ['tab', 'status', 'rows', 'duration_seconds', 'source_read_started_at', 'source_read_completed_at', 'prior_data_preserved', 'limitation_or_error'],
    ['campaign', 'OK', 1, 1, '', '', 'NO', ''],
    ['optional_tab', 'LIMITED', 1, 1, '', '', 'NO', 'optional source unavailable'],
    ['required_tab', 'ERROR', 0, 1, '', '', 'NO', 'failed'],
  ];
  const ss = new FakeSpreadsheet([{
    name: '_export_info', data: rows, maxRows: rows.length, maxColumns: 8,
  }]);
  const sheet = ss.getSheetByName('_export_info');

  api.formatExportInfoSheet(sheet);

  const yellow = sheet.operations.filter((entry) => (
    entry.type === 'background' && entry.value === '#FFF2CC'
  ));
  const red = sheet.operations.filter((entry) => (
    entry.type === 'background' && entry.value === '#F4CCCC'
  ));
  assert.equal(yellow.some((entry) => entry.row === 2), true);
  assert.equal(
    yellow.some((entry) => entry.row === 10 && entry.column === 1 && entry.columnCount === 8),
    true,
  );
  assert.equal(
    red.some((entry) => entry.row === 11 && entry.column === 1 && entry.columnCount === 8),
    true,
  );
  assert.equal(sheet.columnWidths[4] >= 240, true);
});

test('creates deterministic run identity and accepts only a compatible fresh resume', () => {
  const { api } = loadExporter();
  const now = Date.UTC(2026, 7, 25, 12, 0, 0);
  const identity = {
    version: 'v1.0.0',
    accountId: '123-456-7890',
    spreadsheetId: 'sheet-123',
    configSignature: 'config-abc',
  };
  const state = api.createRunState(identity, now, { aggregate: {}, weekly: {}, change: {} }, ['one', 'two']);

  assert.equal(state.status, 'RUNNING');
  assert.equal(state.jobIndex, 0);
  assert.deepEqual(Array.from(state.manifest), ['one', 'two']);
  assert.equal(api.assertStateCompatible(state, identity, now + 60_000, 24), true);
  assert.throws(
    () => api.assertStateCompatible(state, { ...identity, accountId: '999-999-9999' }, now, 24),
    /account/i,
  );
  assert.throws(
    () => api.assertStateCompatible(state, { ...identity, spreadsheetId: 'other-sheet' }, now, 24),
    /workbook/i,
  );
  assert.throws(
    () => api.assertStateCompatible(state, { ...identity, configSignature: 'changed' }, now, 24),
    /configuration/i,
  );
  assert.throws(
    () => api.assertStateCompatible(state, identity, now + (25 * 60 * 60 * 1000), 24),
    /expired/i,
  );
});

test('stable stringification ignores object key insertion order', () => {
  const { api } = loadExporter();

  assert.equal(
    api.stableStringify({ z: 3, nested: { b: 2, a: 1 } }),
    api.stableStringify({ nested: { a: 1, b: 2 }, z: 3 }),
  );
});

test('reset planning is guarded and never targets final report tabs', () => {
  const { api } = loadExporter();
  const sheetNames = [
    '_export_state',
    '__gads_export_stage__campaign',
    '__gads_export_backup__ads',
    '_export_info',
    'campaign',
    'Personal Notes',
  ];

  assert.throws(() => api.planReset(false, sheetNames), /ALLOW_RESET/i);
  assert.deepEqual(Array.from(api.planReset(true, sheetNames)), [
    '_export_state',
    '__gads_export_stage__campaign',
    '__gads_export_backup__ads',
  ]);
});

test('safe row buffer batches sanitized rows and enforces workbook cell limit', () => {
  const { api } = loadExporter();
  const ss = new FakeSpreadsheet([{ name: 'stage', data: [['name', 'value']], maxRows: 2, maxColumns: 2 }]);
  const sheet = ss.getSheetByName('stage');
  const writer = api.createRowBuffer(sheet, ['name', 'value'], {
    batchRows: 2,
    cellLimit: 20,
    retries: 1,
    sleep() {},
  });

  writer.push(['safe', -2]);
  writer.push(['=unsafe', '0000000000000001']);
  writer.flush();
  assert.deepEqual(sheet.data.slice(0, 3), [
    ['name', 'value'],
    ['safe', -2],
    ['=unsafe', '0000000000000001'],
  ]);
  assert.equal(writer.count(), 2);

  const constrained = new FakeSpreadsheet([{ name: 'tiny', data: [['a', 'b']], maxRows: 2, maxColumns: 2 }]);
  const constrainedWriter = api.createRowBuffer(
    constrained.getSheetByName('tiny'),
    ['a', 'b'],
    { batchRows: 3, cellLimit: 4, retries: 1, sleep() {} },
  );
  constrainedWriter.push(['one', 'two']);
  constrainedWriter.push(['three', 'four']);
  assert.throws(() => constrainedWriter.flush(), /cell safety limit/i);
});

test('rolls back rows from an interrupted chunk start without touching prior rows', () => {
  const { api } = loadExporter();
  const ss = new FakeSpreadsheet([{
    name: 'stage',
    data: [['header'], ['complete'], ['partial-a'], ['partial-b']],
  }]);
  const sheet = ss.getSheetByName('stage');

  assert.equal(api.rollbackPartialChunk(sheet, 3), 2);
  assert.deepEqual(sheet.data.slice(0, 4), [['header'], ['complete'], [''], ['']]);
});

test('commits a staged sheet transactionally and removes the backup', () => {
  const { api } = loadExporter();
  const stageName = api.stageSheetName('campaign');
  const ss = new FakeSpreadsheet([
    { name: 'campaign', data: [['old']] },
    { name: stageName, data: [['new']] },
  ]);

  api.commitStagedSheet(ss, stageName, 'campaign');
  assert.deepEqual(ss.getSheetByName('campaign').data, [['new']]);
  assert.equal(ss.getSheetByName(api.backupSheetName('campaign')), null);
});

test('restores the prior final sheet when staged replacement fails', () => {
  const { api } = loadExporter();
  const stageName = api.stageSheetName('campaign');
  const ss = new FakeSpreadsheet([
    { name: 'campaign', data: [['old']] },
    { name: stageName, data: [['new']] },
  ]);
  ss.failRename = { from: stageName, to: 'campaign' };

  assert.throws(() => api.commitStagedSheet(ss, stageName, 'campaign'), /injected rename failure/);
  assert.deepEqual(ss.getSheetByName('campaign').data, [['old']]);
  assert.deepEqual(ss.getSheetByName(stageName).data, [['new']]);
});

function makeEngineState(api) {
  return api.createRunState(
    {
      version: 'v1.0.0',
      accountId: '123-456-7890',
      spreadsheetId: 'sheet-123',
      configSignature: 'config-abc',
    },
    Date.UTC(2026, 7, 25, 12, 0, 0),
    { aggregate: {}, weekly: {}, change: {} },
    ['job-one'],
  );
}

function makeEngineAdapter(overrides = {}) {
  const events = [];
  const adapter = {
    events,
    remainingSeconds() { return 1_000; },
    saveState(state) { events.push(['save', state.status, state.jobIndex, state.chunkIndex, state.chunkInProgress]); },
    writeInfo(state) { events.push(['info', state.status]); },
    writeProgressSummary(state) { events.push(['progress', state.status]); },
    startJob(job) { events.push(['start', job.id]); return `stage-${job.id}`; },
    getChunkCount() { return 1; },
    getChunkStartRow() { return 2; },
    rollbackChunk(job, state) { events.push(['rollback', job.id, state.chunkStartRow]); },
    runChunk(job, _state, chunkIndex) { events.push(['chunk', job.id, chunkIndex]); return 3; },
    commitJob(job) { events.push(['commit', job.id]); },
    commitEmptyLimitedJob(job) { events.push(['commit-empty-limited', job.id]); },
    isExpectedOptionalSourceError(_job, error) {
      return /not supported/i.test(String(error && error.message ? error.message : error));
    },
    abortJob(job) { events.push(['abort', job.id]); },
    finalizeWorkbook(state) { events.push(['finalize', state.status]); },
    publishWorkbook(state) { events.push(['publish', state.status]); },
    writeFailureSummary(state) { events.push(['failure-summary', state.status]); },
    hasPriorFinal() { return false; },
    clearState() { events.push(['clear']); },
    nowMs() { return Date.UTC(2026, 7, 25, 12, 5, 0); },
    ...overrides,
  };
  return adapter;
}

test('pauses cleanly before starting a job when execution time is low', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({ remainingSeconds() { return 100; } });

  const result = api.runManifestEngine(state, [{ id: 'job-one', tab: 'one', required: true }], adapter, 180);

  assert.equal(result.status, 'PAUSED');
  assert.equal(result.jobIndex, 0);
  assert.equal(adapter.events.some(([name]) => name === 'start'), false);
  assert.equal(adapter.events.some(([name]) => name === 'clear'), false);
});

test('invalidates a prior completed START_HERE while a replacement run is resumable', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  state.status = 'COMPLETE';
  const adapter = makeEngineAdapter({ remainingSeconds() { return 100; } });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'PAUSED');
  assert.deepEqual(
    adapter.events.filter(([name]) => name === 'progress'),
    [['progress', 'RUNNING'], ['progress', 'PAUSED']],
  );
  assert.equal(adapter.events.some(([name]) => name === 'start'), false);
});

test('checkpoints before each chunk and pauses between chunks', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const remaining = [1_000, 1_000, 100];
  const adapter = makeEngineAdapter({
    remainingSeconds() { return remaining.shift(); },
    getChunkCount() { return 2; },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true, chunked: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'PAUSED');
  assert.equal(result.chunkIndex, 1);
  assert.equal(result.chunkInProgress, false);
  assert.equal(result.tabs.one.rows, 3);
  assert.deepEqual(adapter.events.filter(([name]) => name === 'chunk'), [['chunk', 'job-one', 0]]);
});

test('resumed jobs preserve duration accumulated by earlier invocations', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  state.stageSheetName = 'stage-job-one';
  state.chunkIndex = 1;
  state.tabs.one = {
    status: 'RUNNING', rows: 3, durationMs: 5_000,
    error: '', limitation: '', partialLimited: false, priorPreserved: false,
  };
  const adapter = makeEngineAdapter({
    getChunkCount() { return 2; },
    nowMs() { return 10_000; },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true, chunked: true }],
    adapter,
    180,
  );

  assert.equal(result.tabs.one.rows, 6);
  assert.equal(result.tabs.one.durationMs, 5_000);
});

test('pauses before the formatting commit when a large final chunk uses the time reserve', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const remaining = [1_000, 1_000, 100];
  const adapter = makeEngineAdapter({
    remainingSeconds() { return remaining.shift(); },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true, chunked: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'PAUSED');
  assert.equal(result.chunkIndex, 1);
  assert.equal(result.chunkInProgress, false);
  assert.equal(adapter.events.some(([name]) => name === 'commit'), false);
  assert.equal(adapter.events.some(([name]) => name === 'clear'), false);
});

test('pauses after the last commit when the finalization reserve is exhausted', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const remaining = [1_000, 1_000, 1_000, 100];
  const first = makeEngineAdapter({
    remainingSeconds() { return remaining.shift(); },
  });

  const paused = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true }],
    first,
    180,
  );

  assert.equal(paused.status, 'PAUSED');
  assert.equal(paused.jobIndex, 1);
  assert.equal(first.events.filter(([name]) => name === 'commit').length, 1);
  assert.equal(first.events.some(([name]) => name === 'finalize'), false);
  assert.equal(first.events.some(([name]) => name === 'clear'), false);
  assert.equal(first.events.some(([name]) => name === 'publish'), false);

  const resumed = makeEngineAdapter();
  const completed = api.runManifestEngine(
    paused,
    [{ id: 'job-one', tab: 'one', required: true }],
    resumed,
    180,
  );

  assert.equal(completed.status, 'COMPLETE');
  assert.equal(resumed.events.some(([name]) => name === 'start'), false);
  assert.deepEqual(
    resumed.events.slice(-3).map(([name]) => name),
    ['finalize', 'clear', 'publish'],
  );
});

test('rechecks the reserve after validation before deleting the checkpoint', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const remaining = [1_000, 1_000, 1_000, 1_000, 100];
  const first = makeEngineAdapter({
    remainingSeconds() { return remaining.shift(); },
  });

  const paused = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true }],
    first,
    180,
  );

  assert.equal(paused.status, 'PAUSED');
  assert.equal(first.events.filter(([name]) => name === 'finalize').length, 1);
  assert.equal(first.events.some(([name]) => name === 'clear'), false);
  assert.equal(first.events.some(([name]) => name === 'publish'), false);

  const resumed = makeEngineAdapter();
  const completed = api.runManifestEngine(
    paused,
    [{ id: 'job-one', tab: 'one', required: true }],
    resumed,
    180,
  );

  assert.equal(completed.status, 'COMPLETE');
  assert.equal(resumed.events.some(([name]) => name === 'start'), false);
  assert.deepEqual(
    resumed.events.slice(-3).map(([name]) => name),
    ['finalize', 'clear', 'publish'],
  );
});

test('rolls back a previously in-progress chunk before rerunning it', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  state.stageSheetName = 'stage-job-one';
  state.chunkInProgress = true;
  state.chunkStartRow = 7;
  const adapter = makeEngineAdapter();

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true, chunked: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE');
  assert.deepEqual(adapter.events.filter(([name]) => name === 'rollback'), [['rollback', 'job-one', 7]]);
  assert.deepEqual(adapter.events.filter(([name]) => name === 'chunk'), [['chunk', 'job-one', 0]]);
  assert.equal(adapter.events.at(-1)[0], 'publish');
});

test('finalizes workbook presentation before deleting the resumable checkpoint', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter();

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE');
  assert.deepEqual(
    adapter.events.slice(-3).map(([name]) => name),
    ['finalize', 'clear', 'publish'],
  );
  assert.equal(
    adapter.events.some(([name, status]) => name === 'info' && status === 'COMPLETE'),
    false,
  );
  assert.equal(
    adapter.events.some(([name, status]) => name === 'info' && status === 'FINALIZING'),
    true,
  );
  assert.deepEqual(
    adapter.events.filter(([name]) => name === 'finalize'),
    [['finalize', 'FINALIZING']],
  );
  assert.deepEqual(
    adapter.events.filter(([name]) => name === 'publish'),
    [['publish', 'COMPLETE']],
  );
});

test('keeps the checkpoint and reports an error when workbook finalization fails', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({
    finalizeWorkbook() { throw new Error('formatting failed'); },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE_WITH_ERRORS');
  assert.match(result.workbookError, /formatting failed/);
  assert.equal(adapter.events.some(([name]) => name === 'clear'), false);
  assert.equal(
    adapter.events.some(([name, status]) => (
      name === 'failure-summary' && status === 'COMPLETE_WITH_ERRORS'
    )),
    true,
  );
});

test('downgrades the workbook and reports an error when checkpoint cleanup fails', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({
    clearState() { throw new Error('checkpoint deletion failed'); },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE_WITH_ERRORS');
  assert.match(result.workbookError, /checkpoint deletion failed/);
  assert.equal(
    adapter.events.some(([name, status]) => (
      name === 'failure-summary' && status === 'COMPLETE_WITH_ERRORS'
    )),
    true,
  );
});

test('reports an error when final native publication fails after checkpoint cleanup', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  let checkpointAvailable = true;
  const adapter = makeEngineAdapter({
    publishWorkbook() { throw new Error('terminal metadata write failed'); },
  });
  adapter.clearState = function clearState() {
    adapter.events.push(['clear']);
    checkpointAvailable = false;
  };
  adapter.saveState = function saveState(current) {
    if (!checkpointAvailable) throw new Error('state sheet is missing');
    adapter.events.push([
      'save', current.status, current.jobIndex, current.chunkIndex, current.chunkInProgress,
    ]);
  };

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: true }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE_WITH_ERRORS');
  assert.match(result.workbookError, /terminal metadata write failed/);
  assert.match(result.workbookError, /could not retain the checkpoint/i);
  assert.equal(adapter.events.some(([name]) => name === 'clear'), true);
  assert.equal(
    adapter.events.some(([name, status]) => (
      name === 'failure-summary' && status === 'COMPLETE_WITH_ERRORS'
    )),
    true,
  );
  const rows = api.buildExportInfoRows(result, [{ id: 'job-one', tab: 'one', required: true }]);
  assert.match(
    String(rows.find((row) => row[0] === 'next_action')[1]),
    /ALLOW_RESET|resetExportState/i,
  );
});

test('preserves a prior tab and completes with errors after a required job failure', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({
    runChunk() { throw new Error('required query failed'); },
    hasPriorFinal() { return true; },
  });

  const result = api.runManifestEngine(state, [{ id: 'job-one', tab: 'one', required: true }], adapter, 180);

  assert.equal(result.status, 'COMPLETE_WITH_ERRORS');
  assert.equal(result.tabs.one.status, 'ERROR_PREVIOUS_PRESERVED');
  assert.match(result.tabs.one.error, /required query failed/);
  assert.equal(adapter.events.some(([name]) => name === 'abort'), true);
  assert.equal(adapter.events.some(([name]) => name === 'clear'), false);
  assert.equal(adapter.events.some(([name]) => name === 'publish'), false);
  const rows = api.buildExportInfoRows(result, [{ id: 'job-one', tab: 'one', required: true }]);
  const nextAction = String(rows.find((row) => row[0] === 'next_action')[1]);
  assert.match(nextAction, /ALLOW_RESET|resetExportState/i);
  assert.doesNotMatch(nextAction, /Run main\(\) again to retry/i);
});

test('labels an unsupported optional job LIMITED without treating it as a required error', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({ runChunk() { throw new Error('resource not supported'); } });

  const result = api.runManifestEngine(state, [{ id: 'job-one', tab: 'one', required: false }], adapter, 180);

  assert.equal(result.status, 'COMPLETE_WITH_LIMITATIONS');
  assert.equal(result.tabs.one.status, 'LIMITED');
  assert.equal(result.tabs.one.rows, 0);
  assert.equal(result.tabs.one.priorPreserved, false);
  assert.match(result.tabs.one.error, /resource not supported/);
  assert.deepEqual(
    adapter.events.filter(([name]) => name === 'commit-empty-limited'),
    [['commit-empty-limited', 'job-one']],
  );
  assert.equal(adapter.events.some(([name]) => name === 'abort'), false);
});

test('optional-source classifier accepts only enumerated query incompatibility codes', () => {
  const { api } = loadExporter();
  for (const code of [
    'UNRECOGNIZED_FIELD',
    'PROHIBITED_FIELD_IN_SELECT_CLAUSE',
    'PROHIBITED_RESOURCE_TYPE_IN_SELECT_CLAUSE',
    'UNSUPPORTED_FIELD',
    'UNSUPPORTED_RESOURCE',
  ]) {
    assert.equal(
      api.isExpectedOptionalSourceUnavailableError(new Error(`QueryError.${code}`)),
      true,
      code,
    );
  }
  for (const message of [
    'resource temporarily unavailable',
    'service unavailable',
    'quota exceeded',
    'permission denied',
    'request timed out',
    'field is not available',
  ]) {
    assert.equal(
      api.isExpectedOptionalSourceUnavailableError(new Error(message)),
      false,
      message,
    );
  }
});

test('every partial-source runner routes caught query errors through the strict classifier', () => {
  const { source } = loadExporter();
  for (const functionName of [
    'resolveGeoTargetMetadata_',
    'runAudienceRuntime_',
    'runAssetExtensionsRuntime_',
  ]) {
    const start = source.indexOf(`function ${functionName}(`);
    const end = source.indexOf('\nfunction ', start + 1);
    assert.notEqual(start, -1, functionName);
    assert.match(
      source.slice(start, end),
      /expectedPartialSourceLimitation_\(/,
      functionName,
    );
  }
});

test('does not hide an optional staging or empty-tab commit failure as LIMITED', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({
    runChunk() { throw new Error('resource not supported'); },
    commitEmptyLimitedJob() { throw new Error('could not commit empty tab'); },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: false }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE_WITH_ERRORS');
  assert.equal(result.tabs.one.status, 'ERROR');
  assert.match(result.tabs.one.error, /resource not supported/);
  assert.match(result.tabs.one.error, /could not commit empty tab/);
  assert.equal(adapter.events.some(([name]) => name === 'abort'), true);
});

test('treats an unexpected optional-job exception as a terminal export error', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({
    runChunk() { throw new TypeError('unexpected row mapping bug'); },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: false }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE_WITH_ERRORS');
  assert.equal(result.tabs.one.status, 'ERROR');
  assert.match(result.tabs.one.error, /unexpected row mapping bug/);
  assert.equal(adapter.events.some(([name]) => name === 'commit-empty-limited'), false);
  assert.equal(adapter.events.some(([name]) => name === 'abort'), true);
});

test('core report catalog includes performance and zero-impression inventory views', () => {
  const { api } = loadExporter();
  const jobs = api.getManifestDefinition();
  const tabs = new Set(jobs.map((job) => job.tab));
  const requiredTabs = [
    'campaign', 'campaign_weekly', 'imp_share', 'keywords', 'search_terms', 'ads',
    'rsa_assets', 'demandgen_assets', 'landing_pages', 'campaign_device_network',
    'ad_schedule', 'campaign_geo', 'ad_group_weekly', 'ad_group',
    'campaign_inventory', 'ad_group_inventory', 'keyword_inventory', 'ad_inventory',
  ];

  for (const tab of requiredTabs) assert.equal(tabs.has(tab), true, `missing ${tab}`);
  assert.equal(api.validateManifest(jobs), true);
});

test('high-volume reports are campaign chunked and queries receive the chunk filter', () => {
  const { api } = loadExporter();
  const jobs = api.getManifestDefinition();
  const ranges = {
    aggregate: { start: '2026-05-27', end: '2026-08-24' },
    weekly: { start: '2026-05-27', end: '2026-08-24' },
  };

  for (const tab of ['keywords', 'search_terms', 'ads', 'landing_pages']) {
    const job = jobs.find((candidate) => candidate.tab === tab);
    assert.equal(job.chunked, true, `${tab} must be chunked`);
    assert.match(api.buildGaqlQuery(job, ranges, ['101', '202']), /campaign\.id IN \(101, 202\)/);
  }
});

test('production GAQL excludes ORDER BY fields rejected by the live account', () => {
  const { api } = loadExporter();
  const jobs = api.getManifestDefinition();
  const ranges = { aggregate: { start: '2026-05-27', end: '2026-08-24' } };

  const accountLinks = jobs.find((job) => job.tab === 'neg_keyword_account_links');
  const accountQuery = api.buildGaqlQuery(accountLinks, ranges, null);
  const accountOrder = accountQuery.match(/ORDER BY (.*)$/m);
  assert.equal(accountOrder, null, 'account negative-list links need no unsupported server-side sort');

  const pmaxSignals = jobs.find((job) => job.tab === 'pmax_audience_signals');
  const pmaxQuery = api.buildGaqlQuery(pmaxSignals, ranges, ['101']);
  assert.match(pmaxQuery, /ORDER BY campaign\.id, asset_group\.id$/m);
  assert.doesNotMatch(
    pmaxQuery.match(/ORDER BY (.*)$/m)[1],
    /asset_group_signal\.resource_name/,
  );
});

test('RSA assets are restricted to Search responsive search ads', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'rsa_assets');
  const query = api.buildGaqlQuery(job, {
    aggregate: { start: '2026-05-27', end: '2026-08-24' },
  }, ['101']);

  assert.match(query, /campaign\.advertising_channel_type = 'SEARCH'/);
  assert.match(query, /ad_group_ad\.ad\.type = 'RESPONSIVE_SEARCH_AD'/);
});

test('source has no general silent row cap or exportToSheet write bypass', () => {
  const { source } = loadExporter();

  assert.doesNotMatch(source, /HEAVY_LIMIT|LIMIT\s+120000/i);
  assert.doesNotMatch(source, /\.exportToSheet\s*\(/);
});

test('streams declarative GAQL rows through the common safe writer', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign_inventory');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: 'stage', data: [headers], maxRows: 10, maxColumns: headers.length,
  }]);
  let capturedQuery = '';
  const reportRows = [{
    'customer.id': '1234567890',
    'customer.descriptive_name': '=unsafe account name',
    'campaign.id': '101',
    'campaign.name': 'Brand Search',
    'campaign.status': 'ENABLED',
  }];
  const runtime = {
    report(query) {
      capturedQuery = query;
      let index = 0;
      return {
        rows() {
          return {
            hasNext() { return index < reportRows.length; },
            next() { return reportRows[index++]; },
          };
        },
      };
    },
    sleep() {},
  };

  const count = api.runGaqlChunk(
    job,
    { aggregate: { start: '2026-05-27', end: '2026-08-24' } },
    ['101'],
    ss.getSheetByName('stage'),
    runtime,
    { batchRows: 10, cellLimit: 100_000, retries: 1, sleep() {} },
  );

  assert.equal(count, 1);
  assert.match(capturedQuery, /campaign\.id IN \(101\)/);
  assert.equal(ss.getSheetByName('stage').data[1][1], '=unsafe account name');
});

test('preserves unavailable Average CPC as blank in direct GAQL output', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: 'stage', data: [headers], maxRows: 10, maxColumns: headers.length,
  }]);
  const reportRows = [{
    'campaign.id': '101',
    'campaign.name': 'No clicks',
    'metrics.impressions': '5',
    'metrics.clicks': '0',
    'metrics.interactions': '0',
    'metrics.average_cpc': '',
    'metrics.cost_micros': '0',
    'metrics.conversions': '0',
  }];
  let index = 0;

  api.runGaqlChunk(
    job,
    { aggregate: { start: '2026-05-27', end: '2026-08-24' } },
    ['101'],
    ss.getSheetByName('stage'),
    {
      report: () => ({ rows: () => ({
        hasNext: () => index < reportRows.length,
        next: () => reportRows[index++],
      }) }),
      sleep() {},
    },
    { batchRows: 10, cellLimit: 100_000, retries: 1, sleep() {} },
  );

  const output = ss.getSheetByName('stage').data[1];
  assert.equal(output[headers.indexOf('metrics.average_cpc')], '');
  assert.equal(output[headers.indexOf('average_cpc')], '');
});

test('GAQL performance tabs omit paused or removed zero-data rows', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'ad_group');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: 'stage', data: [headers], maxRows: 10, maxColumns: headers.length,
  }]);
  const reportRows = [
    { 'campaign.id': '1', 'ad_group.id': '10', 'ad_group.status': 'ENABLED' },
    { 'campaign.id': '1', 'ad_group.id': '20', 'ad_group.status': 'PAUSED' },
    {
      'campaign.id': '1', 'ad_group.id': '30', 'ad_group.status': 'PAUSED',
      'metrics.impressions': '5',
    },
    {
      'campaign.id': '1', 'ad_group.id': '40', 'ad_group.status': 'REMOVED',
      'metrics.conversions': '1',
    },
  ];
  const runtime = {
    report() {
      let index = 0;
      return { rows: () => ({
        hasNext: () => index < reportRows.length,
        next: () => reportRows[index++],
      }) };
    },
    sleep() {},
  };

  const count = api.runGaqlChunk(
    job,
    { aggregate: { start: '2026-05-27', end: '2026-08-24' } },
    ['1'],
    ss.getSheetByName('stage'),
    runtime,
    { batchRows: 10, cellLimit: 100_000, retries: 1, sleep() {} },
  );

  assert.equal(count, 3);
  assert.deepEqual(
    ss.getSheetByName('stage').data.slice(1).map((row) => row[headers.indexOf('ad_group.id')]),
    ['10', '30', '40'],
  );
});

test('inventory tabs join 90-day activity before retaining inactive entities', () => {
  const { api } = loadExporter();
  const specs = [
    {
      tab: 'ad_group_inventory', idField: 'ad_group.id', statusField: 'ad_group.status',
      base: { 'campaign.id': '1' },
    },
    {
      tab: 'keyword_inventory', idField: 'ad_group_criterion.criterion_id',
      statusField: 'ad_group_criterion.status', base: { 'campaign.id': '1', 'ad_group.id': '10' },
    },
    {
      tab: 'ad_inventory', idField: 'ad_group_ad.ad.id', statusField: 'ad_group_ad.status',
      base: { 'campaign.id': '1', 'ad_group.id': '10' },
    },
  ];

  for (const spec of specs) {
    const job = api.getManifestDefinition().find((candidate) => candidate.tab === spec.tab);
    const headers = Array.from(api.headersForJob(job));
    const ss = new FakeSpreadsheet([{
      name: 'stage', data: [headers], maxRows: 10, maxColumns: headers.length,
    }]);
    const configRows = [
      { ...spec.base, [spec.idField]: '1', [spec.statusField]: 'ENABLED' },
      { ...spec.base, [spec.idField]: '2', [spec.statusField]: 'PAUSED' },
      { ...spec.base, [spec.idField]: '3', [spec.statusField]: 'REMOVED' },
    ];
    const activityRows = [
      { ...spec.base, [spec.idField]: '3', 'metrics.impressions': '4' },
    ];
    const queries = [];
    const runtime = {
      report(query) {
        queries.push(query);
        const rows = /metrics\.impressions/.test(query) ? activityRows : configRows;
        let index = 0;
        return { rows: () => ({
          hasNext: () => index < rows.length,
          next: () => rows[index++],
        }) };
      },
      sleep() {},
    };

    const count = api.runGaqlChunk(
      job,
      { aggregate: { start: '2026-05-27', end: '2026-08-24' } },
      ['1'],
      ss.getSheetByName('stage'),
      runtime,
      { batchRows: 10, cellLimit: 100_000, retries: 1, sleep() {} },
    );

    assert.equal(queries.length, 2, `${spec.tab} must query activity and configuration`);
    assert.equal(count, 2, `${spec.tab} should omit only the inactive zero-data row`);
    assert.deepEqual(
      ss.getSheetByName('stage').data.slice(1).map((row) => row[headers.indexOf(spec.idField)]),
      ['1', '3'],
    );
  }
});

test('direct ad-group negatives use the canonical 90-day ad-group universe and filtered rows cannot reappear in the union', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'neg_keywords_ad_group');
  const headers = Array.from(api.headersForJob(job));
  const ss = new FakeSpreadsheet([{
    name: 'stage', data: [headers], maxRows: 10, maxColumns: headers.length,
  }]);
  const configRows = [
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED',
      'ad_group.id': '10', 'ad_group.status': 'ENABLED',
      'ad_group_criterion.criterion_id': '110',
      'ad_group_criterion.status': 'ENABLED',
      'ad_group_criterion.keyword.text': 'enabled zero',
    },
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED',
      'ad_group.id': '20', 'ad_group.status': 'PAUSED',
      'ad_group_criterion.criterion_id': '120',
      'ad_group_criterion.status': 'ENABLED',
      'ad_group_criterion.keyword.text': 'paused zero',
    },
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED',
      'ad_group.id': '30', 'ad_group.status': 'PAUSED',
      'ad_group_criterion.criterion_id': '130',
      'ad_group_criterion.status': 'ENABLED',
      'ad_group_criterion.keyword.text': 'paused active',
    },
    {
      'campaign.id': '2', 'campaign.status': 'PAUSED',
      'ad_group.id': '40', 'ad_group.status': 'ENABLED',
      'ad_group_criterion.criterion_id': '140',
      'ad_group_criterion.status': 'ENABLED',
      'ad_group_criterion.keyword.text': 'paused campaign zero',
    },
    {
      'campaign.id': '2', 'campaign.status': 'PAUSED',
      'ad_group.id': '50', 'ad_group.status': 'ENABLED',
      'ad_group_criterion.criterion_id': '150',
      'ad_group_criterion.status': 'ENABLED',
      'ad_group_criterion.keyword.text': 'paused campaign active',
    },
  ];
  const activityRows = [
    { 'campaign.id': '1', 'ad_group.id': '30', 'metrics.impressions': '5' },
    { 'campaign.id': '2', 'ad_group.id': '50', 'metrics.clicks': '1' },
  ];
  const queries = [];
  const runtime = {
    report(query) {
      queries.push(query);
      const rows = /FROM ad_group\n/.test(query) && /metrics\.impressions/.test(query)
        ? activityRows : configRows;
      let index = 0;
      return { rows: () => ({
        hasNext: () => index < rows.length,
        next: () => rows[index++],
      }) };
    },
    sleep() {},
  };

  const count = api.runGaqlChunk(
    job,
    { aggregate: { start: '2026-05-27', end: '2026-08-24' } },
    ['1', '2'],
    ss.getSheetByName('stage'),
    runtime,
    { batchRows: 10, cellLimit: 100_000, retries: 1, sleep() {} },
  );
  const directRows = ss.getSheetByName('stage').data.slice(1).map((values) => (
    Object.fromEntries(headers.map((header, index) => [header, values[index]]))
  ));
  const unionRows = api.buildNegativeUnionRows({ neg_keywords_ad_group: directRows });

  assert.equal(queries.length, 2, 'negative criteria require a separate parent activity lookup');
  assert.equal(count, 3);
  assert.deepEqual(directRows.map((row) => row['ad_group.id']), ['10', '30', '50']);
  assert.deepEqual(Array.from(unionRows, (row) => row[5]), ['10', '30', '50']);
});

test('PMax asset-group inventory includes enabled zero-data groups and signals use that exact universe', () => {
  const { api } = loadExporter();
  const jobs = new Map(api.getManifestDefinition().map((job) => [job.tab, job]));
  const pmax = jobs.get('pmax_asset_groups');
  const staticRows = [
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED',
      'asset_group.id': '100', 'asset_group.name': 'Enabled zero', 'asset_group.status': 'ENABLED',
    },
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED',
      'asset_group.id': '200', 'asset_group.name': 'Paused zero', 'asset_group.status': 'PAUSED',
    },
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED',
      'asset_group.id': '300', 'asset_group.name': 'Paused active', 'asset_group.status': 'PAUSED',
    },
  ];
  const metricRows = [
    {
      'campaign.id': '1', 'asset_group.id': '300', 'asset_group.status': 'PAUSED',
      'metrics.impressions': '5', 'metrics.interactions': '4', 'metrics.conversions': '1',
    },
    {
      'campaign.id': '1', 'asset_group.id': '400', 'asset_group.status': 'REMOVED',
      'metrics.clicks': '2',
    },
    {
      'campaign.id': '1', 'asset_group.id': '500', 'asset_group.status': 'REMOVED',
    },
  ];

  const queries = api.buildEntityPerformanceQueries(
    pmax,
    { start: '2026-05-27', end: '2026-08-24' },
    ['1'],
  );
  assert.match(queries.current, /FROM asset_group/);
  assert.doesNotMatch(queries.current, /metrics\.|segments\.date/);
  assert.match(queries.activity, /metrics\.impressions/);
  assert.match(queries.activity, /metrics\.interactions/);
  assert.match(queries.activity, /segments\.date BETWEEN '2026-05-27' AND '2026-08-24'/);

  const pmaxRows = api.buildEntityPerformanceRows(pmax, staticRows, metricRows);
  const pmaxHeaders = Array.from(api.headersForJob(pmax));
  const pmaxObjects = pmaxRows.map((values) => (
    Object.fromEntries(pmaxHeaders.map((header, index) => [header, values[index]]))
  ));
  assert.deepEqual(Array.from(pmaxObjects, (row) => row['asset_group.id']), ['100', '300', '400']);
  assert.equal(pmaxObjects[0]['metrics.impressions'], 0);
  assert.equal(pmaxObjects[0]['metrics.average_cpc'], '');
  assert.equal(pmaxObjects[0].average_cpc, '');
  assert.equal(pmaxObjects[0].conversion_rate, '');
  assert.equal(pmaxObjects[1]['metrics.impressions'], 5);
  assert.equal(pmaxObjects[1]['metrics.interactions'], 4);
  assert.equal(pmaxObjects[1].conversion_rate, 0.25);

  const signals = jobs.get('pmax_audience_signals');
  const signalHeaders = Array.from(api.headersForJob(signals));
  const signalSheet = new FakeSpreadsheet([{
    name: 'signals', data: [signalHeaders], maxRows: 10, maxColumns: signalHeaders.length,
  }]).getSheetByName('signals');
  const signalRows = [
    { 'campaign.id': '1', 'asset_group.id': '100', 'asset_group_signal.resource_name': 'signals/100' },
    { 'campaign.id': '1', 'asset_group.id': '200', 'asset_group_signal.resource_name': 'signals/200' },
    { 'campaign.id': '1', 'asset_group.id': '300', 'asset_group_signal.resource_name': 'signals/300' },
    { 'campaign.id': '1', 'asset_group.id': '400', 'asset_group_signal.resource_name': 'signals/400' },
  ];
  let signalIndex = 0;
  const signalQuery = api.buildGaqlQuery(signals, {}, ['1']);
  assert.doesNotMatch(signalQuery, /asset_group\.status/, 'status is enriched from the canonical parent tab');
  const signalCount = api.runGaqlChunk(
    signals,
    {},
    ['1'],
    signalSheet,
    {
      report: () => ({ rows: () => ({
        hasNext: () => signalIndex < signalRows.length,
        next: () => signalRows[signalIndex++],
      }) }),
      sleep() {},
    },
    {
      batchRows: 10, cellLimit: 100_000, retries: 1, sleep() {},
      eligibleSourceRows: pmaxObjects,
    },
  );

  assert.equal(signalCount, 3);
  const signalObjects = signalSheet.data.slice(1).map((values) => (
    Object.fromEntries(signalHeaders.map((header, index) => [header, values[index]]))
  ));
  assert.deepEqual(Array.from(signalObjects, (row) => row['asset_group.id']), ['100', '300', '400']);
  assert.deepEqual(
    Array.from(signalObjects, (row) => row['asset_group.status']),
    ['ENABLED', 'PAUSED', 'REMOVED'],
  );
});

test('asset performance tabs apply status plus 90-day activity eligibility', () => {
  const { api } = loadExporter();
  const jobs = new Map(api.getManifestDefinition().map((job) => [job.tab, job]));
  const expectedStatusFields = {
    rsa_assets: ['campaign.status', 'ad_group.status', 'ad_group_ad.status', 'ad_group_ad_asset_view.enabled'],
    demandgen_assets: ['campaign.status', 'ad_group.status', 'ad_group_ad.status', 'ad_group_ad_asset_view.enabled'],
    pmax_asset_group_weekly: ['campaign.status', 'asset_group.status'],
    pmax_assets: ['campaign.status', 'asset_group.status', 'asset_group_asset.status'],
  };

  for (const [tab, fields] of Object.entries(expectedStatusFields)) {
    assert.deepEqual(Array.from(jobs.get(tab).surveyStatusFields), fields, `${tab} status contract`);
  }
  assert.deepEqual(
    Array.from(jobs.get('pmax_asset_groups').eligibility.statusFields),
    ['campaign.status', 'asset_group.status'],
    'pmax_asset_groups hybrid inventory status contract',
  );

  const rsa = jobs.get('rsa_assets');
  const headers = Array.from(api.headersForJob(rsa));
  const ss = new FakeSpreadsheet([{
    name: 'stage', data: [headers], maxRows: 10, maxColumns: headers.length,
  }]);
  const rows = [
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED', 'ad_group.id': '2',
      'ad_group.status': 'ENABLED', 'ad_group_ad.ad.id': '3', 'ad_group_ad.status': 'ENABLED',
      'ad_group_ad_asset_view.enabled': true, 'asset.id': '10',
    },
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED', 'ad_group.id': '2',
      'ad_group.status': 'ENABLED', 'ad_group_ad.ad.id': '3', 'ad_group_ad.status': 'ENABLED',
      'ad_group_ad_asset_view.enabled': false, 'asset.id': '20',
    },
    {
      'campaign.id': '1', 'campaign.status': 'ENABLED', 'ad_group.id': '2',
      'ad_group.status': 'ENABLED', 'ad_group_ad.ad.id': '3', 'ad_group_ad.status': 'PAUSED',
      'ad_group_ad_asset_view.enabled': false, 'asset.id': '30', 'metrics.impressions': '2',
    },
  ];
  let index = 0;
  const runtime = {
    report: () => ({ rows: () => ({
      hasNext: () => index < rows.length,
      next: () => rows[index++],
    }) }),
    sleep() {},
  };

  assert.equal(api.runGaqlChunk(
    rsa,
    { aggregate: { start: '2026-05-27', end: '2026-08-24' } },
    ['1'],
    ss.getSheetByName('stage'),
    runtime,
    { batchRows: 10, cellLimit: 100_000, retries: 1, sleep() {} },
  ), 2);
  assert.deepEqual(
    ss.getSheetByName('stage').data.slice(1).map((row) => row[headers.indexOf('asset.id')]),
    ['10', '30'],
  );
});

test('keeps raw micros and adds readable budget, bid, and target columns', () => {
  const { api } = loadExporter();
  const jobs = new Map(api.getManifestDefinition().map((job) => [job.tab, job]));
  const expected = {
    campaign: [
      'campaign_budget.amount',
      'campaign_budget.recommended_budget_amount',
      'campaign_budget.total_amount',
      'campaign.maximize_conversions.target_cpa',
      'campaign.target_cpa.target_cpa',
    ],
    campaign_weekly: [
      'campaign_budget.amount',
      'campaign.maximize_conversions.target_cpa',
      'campaign.target_cpa.target_cpa',
    ],
    ad_group: ['ad_group.target_cpa'],
    ad_group_weekly: ['ad_group.target_cpa'],
    campaign_inventory: ['campaign_budget.amount'],
    ad_group_inventory: [
      'ad_group.cpc_bid',
      'ad_group.cpm_bid',
      'ad_group.target_cpa',
    ],
  };

  for (const [tab, readableHeaders] of Object.entries(expected)) {
    const job = jobs.get(tab);
    const headers = Array.from(api.headersForJob(job));
    for (const header of readableHeaders) {
      assert.equal(headers.includes(header), true, `${tab} missing readable ${header}`);
      assert.equal(headers.includes(`${header}_micros`), true, `${tab} missing raw ${header}_micros`);
    }
  }

  const campaign = jobs.get('campaign');
  const amount = campaign.derived.find((column) => column.header === 'campaign_budget.amount');
  assert.equal(amount.compute({ 'campaign_budget.amount_micros': '12,500,000' }), 12.5);
  assert.equal(amount.compute({ 'campaign_budget.amount_micros': '' }), '');
});

test('catalog covers direct, shared-list, and account-level negative keywords', () => {
  const { api } = loadExporter();
  const jobs = api.getManifestDefinition();
  const tabs = new Set(jobs.map((job) => job.tab));
  for (const tab of [
    'neg_keywords_campaign', 'neg_keywords_ad_group', 'neg_keywords_shared',
    'neg_keyword_shared_links', 'neg_keyword_account_links', 'negative_keywords_all',
  ]) assert.equal(tabs.has(tab), true, `missing ${tab}`);

  const shared = jobs.find((job) => job.tab === 'neg_keywords_shared');
  const sharedQuery = api.buildGaqlQuery(shared, {}, null);
  assert.match(sharedQuery, /shared_criterion\.keyword\.text/);
  assert.match(sharedQuery, /ACCOUNT_LEVEL_NEGATIVE_KEYWORDS/);
  assert.match(sharedQuery, /shared_set\.status != 'REMOVED'/);

  const directCampaign = jobs.find((job) => job.tab === 'neg_keywords_campaign');
  assert.match(
    api.buildGaqlQuery(directCampaign, {}, ['101']),
    /campaign_criterion\.status NOT IN \('PAUSED', 'REMOVED'\)/,
  );
  const directAdGroup = jobs.find((job) => job.tab === 'neg_keywords_ad_group');
  assert.match(
    api.buildGaqlQuery(directAdGroup, {}, ['101']),
    /ad_group_criterion\.status NOT IN \('PAUSED', 'REMOVED'\)/,
  );

  const accountLinks = jobs.find((job) => job.tab === 'neg_keyword_account_links');
  assert.match(
    api.buildGaqlQuery(accountLinks, {}, null),
    /customer_negative_criterion\.negative_keyword_list\.shared_set/,
  );
});

test('catalog includes location and proximity targeting with current proximity fields', () => {
  const { api } = loadExporter();
  const jobs = api.getManifestDefinition();
  const proximity = jobs.find((job) => job.tab === 'geo_proximity_targets');
  const query = api.buildGaqlQuery(proximity, {}, ['101']);

  assert.ok(jobs.find((job) => job.tab === 'geo_targets'));
  assert.match(query, /campaign_criterion\.proximity\.radius/);
  assert.match(query, /campaign_criterion\.proximity\.geo_point\.latitude_in_micro_degrees/);
  assert.match(query, /campaign_criterion\.proximity\.address\.postal_code/);
});

test('audience coverage declares both campaign and ad-group scopes plus PMax signals', () => {
  const { api } = loadExporter();
  const jobs = api.getManifestDefinition();
  const audience = jobs.find((job) => job.tab === 'user_list_performance');
  const signals = jobs.find((job) => job.tab === 'pmax_audience_signals');

  assert.deepEqual(Array.from(audience.scopes), ['CAMPAIGN', 'AD_GROUP']);
  assert.equal(jobs.some((job) => job.tab === 'user_lists'), false);
  assert.ok(signals);
  const query = api.buildGaqlQuery(signals, {}, ['101']);
  assert.match(query, /asset_group_signal\.audience\.audience/);
  assert.match(query, /asset_group_signal\.search_theme\.text/);
});

test('catalog retains conversion, quality-score, asset, PMax, and ad-to-landing-page views', () => {
  const { api } = loadExporter();
  const tabs = new Set(api.getManifestDefinition().map((job) => job.tab));
  for (const tab of [
    'conversion_actions', 'conversion_action_config', 'quality_score_keywords',
    'asset_extensions', 'pmax_asset_groups', 'pmax_assets', 'ad_to_lp_map',
  ]) assert.equal(tabs.has(tab), true, `missing ${tab}`);
});

test('uses current API fields for campaign dates and PMax asset serving status', () => {
  const { api } = loadExporter();
  const jobs = api.getManifestDefinition();
  const campaignInventory = jobs.find((job) => job.tab === 'campaign_inventory');
  const campaignQuery = api.buildGaqlQuery(campaignInventory, {}, ['101']);
  assert.match(campaignQuery, /campaign\.start_date_time/);
  assert.match(campaignQuery, /campaign\.end_date_time/);
  assert.doesNotMatch(campaignQuery, /campaign\.start_date(?:,|\n)/);
  assert.doesNotMatch(campaignQuery, /campaign\.end_date(?:,|\n)/);

  const pmaxAssets = jobs.find((job) => job.tab === 'pmax_assets');
  const headers = Array.from(api.headersForJob(pmaxAssets));
  for (const field of [
    'asset_group_asset.primary_status',
    'asset_group_asset.primary_status_reasons',
    'asset_group_asset.source',
  ]) assert.equal(headers.includes(field), true, `missing ${field}`);
  assert.equal(headers.includes('asset_group_asset.performance_label'), false);
});

test('uses Google interactions as the conversion-rate denominator in every performance output', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const performanceJobs = manifest.filter((job) => api.headersForJob(job).includes('conversion_rate'));

  assert.equal(performanceJobs.length, 19);
  for (const job of performanceJobs) {
    assert.equal(
      api.headersForJob(job).includes('metrics.interactions'),
      true,
      `${job.tab} must expose the Google interactions denominator`,
    );
    const rate = job.derived.find((column) => column.header === 'conversion_rate');
    assert.ok(rate, `${job.tab} must derive conversion_rate`);
    assert.deepEqual(Array.from(rate.sourceFields), ['metrics.conversions', 'metrics.interactions']);
    assert.equal(rate.compute({
      'metrics.clicks': 5,
      'metrics.interactions': 12,
      'metrics.conversions': 3,
    }), 0.25, `${job.tab} conversion rate`);
    assert.equal(rate.compute({
      'metrics.clicks': 5,
      'metrics.interactions': 0,
      'metrics.conversions': 0,
    }), '', `${job.tab} zero-interaction conversion rate`);
  }

  for (const definition of api.buildAudienceQueries(
    { start: '2026-05-27', end: '2026-08-24' },
    ['1'],
  )) {
    assert.match(definition.query, /metrics\.interactions/);
  }
});

test('pins every live report call to the reviewed Google Ads API version', () => {
  const calls = [];
  const { api } = loadExporter({
    AdsApp: {
      report(query, options) {
        calls.push({ query, options });
        return { rows: () => ({ hasNext: () => false }) };
      },
    },
  });

  api.reportRuntime('SELECT customer.id FROM customer');
  api.reportRuntime(
    'SELECT geographic_view.country_criterion_id FROM geographic_view',
    { resolveGeoNames: false },
  );

  assert.deepEqual(JSON.parse(JSON.stringify(calls)), [
    {
      query: 'SELECT customer.id FROM customer',
      options: { apiVersion: 'v25' },
    },
    {
      query: 'SELECT geographic_view.country_criterion_id FROM geographic_view',
      options: { apiVersion: 'v25', resolveGeoNames: false },
    },
  ]);
});

test('unifies direct, shared campaign-list, and account-list negatives with provenance', () => {
  const { api } = loadExporter();
  const rows = api.buildNegativeUnionRows({
    neg_keywords_campaign: [{
      'campaign.id': '1', 'campaign.name': 'Brand',
      'campaign_criterion.criterion_id': '11', 'campaign_criterion.keyword.text': 'free',
      'campaign_criterion.keyword.match_type': 'BROAD', 'campaign_criterion.status': 'ENABLED',
    }],
    neg_keywords_ad_group: [{
      'campaign.id': '1', 'campaign.name': 'Brand', 'ad_group.id': '2', 'ad_group.name': 'Exact',
      'ad_group_criterion.criterion_id': '22', 'ad_group_criterion.keyword.text': 'jobs',
      'ad_group_criterion.keyword.match_type': 'PHRASE', 'ad_group_criterion.status': 'ENABLED',
    }],
    neg_keywords_shared: [
      {
        'shared_set.id': '100', 'shared_set.name': 'Campaign Exclusions',
        'shared_set.type': 'NEGATIVE_KEYWORDS', 'shared_set.status': 'ENABLED',
        'shared_criterion.criterion_id': '33', 'shared_criterion.keyword.text': 'training',
        'shared_criterion.keyword.match_type': 'BROAD',
      },
      {
        'shared_set.id': '200', 'shared_set.name': 'Account Exclusions',
        'shared_set.type': 'ACCOUNT_LEVEL_NEGATIVE_KEYWORDS', 'shared_set.status': 'ENABLED',
        'shared_criterion.criterion_id': '44', 'shared_criterion.keyword.text': 'manual',
        'shared_criterion.keyword.match_type': 'EXACT',
      },
    ],
    neg_keyword_shared_links: [{
      'campaign.id': '1', 'campaign.name': 'Brand', 'shared_set.id': '100',
    }],
    neg_keyword_account_links: [{
      'customer_negative_criterion.negative_keyword_list.shared_set': 'customers/9/sharedSets/200',
    }],
  });

  assert.deepEqual(JSON.parse(JSON.stringify(rows.map((row) => [row[0], row[1], row[8]]))), [
    ['DIRECT_CAMPAIGN', 'CAMPAIGN', 'free'],
    ['DIRECT_AD_GROUP', 'AD_GROUP', 'jobs'],
    ['SHARED_LIST', 'CAMPAIGN_LIST', 'training'],
    ['ACCOUNT_LIST', 'ACCOUNT', 'manual'],
  ]);
});

test('derived outputs refuse to mix current rows with a preserved stale source tab', () => {
  const { api } = loadExporter();

  assert.deepEqual(
    Array.from(api.derivedSourceLimitations(
      ['direct', 'shared'],
      {
        direct: { status: 'OK', priorPreserved: false },
        shared: { status: 'LIMITED', priorPreserved: false },
      },
    )),
    ['shared source status is LIMITED'],
  );
  assert.throws(
    () => api.derivedSourceLimitations(
      ['direct', 'shared'],
      {
        direct: { status: 'OK', priorPreserved: false },
        shared: { status: 'LIMITED', priorPreserved: true },
      },
    ),
    /refusing to mix.*preserved source tab shared/i,
  );
});

test('builds separate campaign and ad-group audience performance queries', () => {
  const { api } = loadExporter();
  const queries = api.buildAudienceQueries(
    { start: '2026-05-27', end: '2026-08-24' },
    ['101', '202'],
  );

  assert.equal(queries.length, 2);
  assert.match(queries[0].query, /FROM campaign_audience_view/);
  assert.match(queries[0].query, /campaign_criterion\.user_list\.user_list/);
  assert.match(queries[1].query, /FROM ad_group_audience_view/);
  assert.match(queries[1].query, /ad_group_criterion\.user_list\.user_list/);
  for (const query of queries) assert.match(query.query, /campaign\.id IN \(101, 202\)/);
});

test('commits partial optional output and labels it LIMITED', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({
    runChunk() { return { rows: 2, status: 'LIMITED', limitation: 'AD_GROUP scope unsupported' }; },
  });

  const result = api.runManifestEngine(
    state,
    [{ id: 'job-one', tab: 'one', required: false }],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE_WITH_LIMITATIONS');
  assert.equal(result.tabs.one.status, 'LIMITED');
  assert.equal(result.tabs.one.rows, 2);
  assert.match(result.tabs.one.limitation, /AD_GROUP/);
  assert.equal(adapter.events.some(([name]) => name === 'commit'), true);
});

test('required errors take precedence over optional limitations in the overall status', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  const adapter = makeEngineAdapter({
    runChunk(job) {
      if (job.id === 'optional') {
        return { rows: 1, status: 'LIMITED', limitation: 'one scope unavailable' };
      }
      throw new Error('required source failed');
    },
  });

  const result = api.runManifestEngine(
    state,
    [
      { id: 'optional', tab: 'optional', required: false },
      { id: 'required', tab: 'required', required: true },
    ],
    adapter,
    180,
  );

  assert.equal(result.status, 'COMPLETE_WITH_ERRORS');
  assert.equal(result.tabs.optional.status, 'LIMITED');
  assert.equal(result.tabs.required.status, 'ERROR');
});

test('change-history headers and query omit sensitive fields by default', () => {
  const { api } = loadExporter();
  const safeHeaders = Array.from(api.changeHistoryColumns(false), (column) => column.header);
  const sensitiveHeaders = Array.from(api.changeHistoryColumns(true), (column) => column.header);
  const safeQuery = api.buildChangeHistoryQuery(
    { start: '2026-07-28', end: '2026-08-24' },
    '2026-08-24 23:59:59',
    false,
    10_000,
  );

  for (const field of ['user_email', 'old_resource', 'new_resource']) {
    assert.equal(safeHeaders.includes(field), false);
    assert.equal(sensitiveHeaders.includes(field), true);
    assert.doesNotMatch(safeQuery, new RegExp(`change_event\\.${field}`));
  }
  assert.match(safeQuery, /LIMIT 10000/);
  assert.match(safeQuery, /change_event\.client_type/);
  assert.match(
    safeQuery,
    /change_event\.client_type IN \('GOOGLE_ADS_WEB_CLIENT'\)/,
  );
  assert.doesNotMatch(safeQuery, /GOOGLE_ADS_EDITOR/);
});

test('stops change-history paging transparently before the execution deadline', () => {
  const { api } = loadExporter();
  const page = [
    { 'change_event.resource_name': 'a', 'change_event.change_date_time': '2026-08-24 12:00:00' },
    { 'change_event.resource_name': 'b', 'change_event.change_date_time': '2026-08-23 12:00:00' },
  ];
  let calls = 0;
  const runtime = {
    report() {
      calls += 1;
      let index = 0;
      return { rows: () => ({
        hasNext: () => index < page.length,
        next: () => page[index++],
      }) };
    },
    remainingSeconds: () => 100,
    minRemainingSeconds: 180,
  };

  const result = api.paginateChangeHistory(
    { start: '2026-07-28', end: '2026-08-24' }, false, runtime, () => {}, 2,
  );

  assert.equal(calls, 1);
  assert.equal(result.rows, 2);
  assert.equal(result.limited, true);
  assert.match(result.limitation, /execution time/i);
});

test('paginates inclusive change-event boundaries and deduplicates repeated rows', () => {
  const { api } = loadExporter();
  const pages = [
    [
      { 'change_event.resource_name': 'a', 'change_event.change_date_time': '2026-08-24 12:00:00' },
      { 'change_event.resource_name': 'b', 'change_event.change_date_time': '2026-08-23 10:00:00' },
      { 'change_event.resource_name': 'c', 'change_event.change_date_time': '2026-08-23 10:00:00' },
    ],
    [
      { 'change_event.resource_name': 'b', 'change_event.change_date_time': '2026-08-23 10:00:00' },
      { 'change_event.resource_name': 'c', 'change_event.change_date_time': '2026-08-23 10:00:00' },
      { 'change_event.resource_name': 'd', 'change_event.change_date_time': '2026-08-22 09:00:00' },
    ],
    [
      { 'change_event.resource_name': 'd', 'change_event.change_date_time': '2026-08-22 09:00:00' },
    ],
  ];
  const emitted = [];
  const runtime = {
    report() {
      const rows = pages.shift();
      let index = 0;
      return { rows: () => ({
        hasNext: () => index < rows.length,
        next: () => rows[index++],
      }) };
    },
  };

  const result = api.paginateChangeHistory(
    { start: '2026-07-28', end: '2026-08-24' },
    false,
    runtime,
    (row) => emitted.push(row['change_event.resource_name']),
    3,
  );

  assert.deepEqual(emitted, ['a', 'b', 'c', 'd']);
  assert.equal(result.rows, 4);
  assert.equal(result.pages, 3);
  assert.equal(result.limited, false);
});

test('marks change-history pagination LIMITED when an inclusive cursor cannot advance', () => {
  const { api } = loadExporter();
  const stuckPage = [
    { 'change_event.resource_name': 'a', 'change_event.change_date_time': '2026-08-24 12:00:00' },
    { 'change_event.resource_name': 'b', 'change_event.change_date_time': '2026-08-24 12:00:00' },
  ];
  const runtime = {
    report() {
      let index = 0;
      return { rows: () => ({
        hasNext: () => index < stuckPage.length,
        next: () => stuckPage[index++],
      }) };
    },
  };

  const result = api.paginateChangeHistory(
    { start: '2026-07-28', end: '2026-08-24' }, false, runtime, () => {}, 2,
  );

  assert.equal(result.rows, 2);
  assert.equal(result.limited, true);
  assert.match(result.limitation, /cursor.*advance/i);
});

test('builds rectangular export-info rows with owner, run, and per-tab status', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  state.status = 'PAUSED';
  state.currentJobId = 'keywords';
  state.ranges = {
    aggregate: { start: '2026-05-27', end: '2026-08-24' },
    weekly: { start: '2026-05-27', end: '2026-08-24' },
    change: { start: '2026-07-28', end: '2026-08-24' },
  };
  state.tabs.campaign = {
    status: 'LIMITED', rows: 12, durationMs: 1_500, priorPreserved: false,
    limitation: 'one scope unavailable', error: 'later chunk failed',
  };

  const rows = api.buildExportInfoRows(state, [{ id: 'campaign', tab: 'campaign' }]);

  assert.equal(rows[0][0], api.OWNER_KEY);
  assert.equal(rows.every((row) => row.length === 8), true);
  assert.equal(rows.some((row) => row[0] === 'campaign' && row[1] === 'LIMITED' && row[2] === 12), true);
  const campaignRow = rows.find((row) => row[0] === 'campaign');
  const statusHeader = rows.find((row) => row[0] === 'tab');
  const limitationColumn = statusHeader.indexOf('limitation_or_error');
  assert.match(campaignRow[limitationColumn], /one scope unavailable/);
  assert.match(campaignRow[limitationColumn], /later chunk failed/);
  assert.equal(rows.some((row) => String(row.join(' ')).includes('Run main() again')), true);
});

test('export-info next action makes limited coverage explicit', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  state.status = 'COMPLETE_WITH_LIMITATIONS';
  state.tabs.one = {
    status: 'LIMITED', rows: 2, durationMs: 1000, priorPreserved: false,
    limitation: 'audience scope unavailable', error: '',
  };

  const rows = api.buildExportInfoRows(state, [{ id: 'job-one', tab: 'one', required: false }]);
  const statusRow = rows.find((row) => row[0] === 'run_id');
  const nextAction = rows.find((row) => row[0] === 'next_action');

  assert.equal(statusRow[3], 'COMPLETE_WITH_LIMITATIONS');
  assert.match(nextAction[1], /limited coverage/i);
});

test('export-info exposes workbook grid usage and remaining safety headroom', () => {
  const { api } = loadExporter();
  const state = makeEngineState(api);
  state.workbookGridCells = 3_000_000;
  state.workbookCellSafetyLimit = 9_000_000;

  const rows = api.buildExportInfoRows(state, [{ id: 'job-one', tab: 'one' }]);
  const nextAction = rows.find((row) => row[0] === 'next_action');

  assert.equal(nextAction[2], 'workbook_grid_cells');
  assert.equal(nextAction[3], 3_000_000);
  assert.equal(nextAction[4], 'cell_safety_headroom');
  assert.equal(nextAction[5], 6_000_000);
});

test('data dictionary has one unique entry for every final output tab', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const rows = api.buildDataDictionaryRows(manifest);
  const headers = Array.from(rows[0]);
  const objects = rows.slice(1).map((row) => (
    Object.fromEntries(headers.map((header, index) => [header, row[index]]))
  ));
  const tabNames = objects.map((row) => row.tab);

  assert.equal(tabNames.includes('START_HERE'), true);
  assert.equal(tabNames.includes('_export_info'), true);
  assert.equal(tabNames.includes('_data_dictionary'), true);
  assert.equal(tabNames.includes('change_history'), true);
  assert.equal(new Set(tabNames).size, tabNames.length);
  assert.equal(
    tabNames.length,
    manifest.length + 2,
    'dictionary job replaces its own metadata entry while START_HERE and export info are added',
  );
  const byTab = Object.fromEntries(objects.map((row) => [row.tab, row]));
  assert.match(byTab.campaign_inventory.material_filters, /currently non-inactive campaigns|90-day activity/i);
  assert.match(byTab.ad_inventory.material_filters, /inactive.*zero.*omitted/i);
  assert.match(byTab.rsa_assets.material_filters, /inactive.*zero.*omitted/i);
  assert.equal(headers.includes('column_schema'), false);
  assert.ok(headers.includes('field_dictionary_reference'));
  assert.equal(byTab.campaign.field_dictionary_reference, '_field_dictionary');
});

test('dictionary keys use exact output headers and include every row-defining dimension', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const rows = api.buildDataDictionaryRows(manifest);
  const headers = Array.from(rows[0]);
  const objects = rows.slice(1).map((row) => (
    Object.fromEntries(headers.map((header, index) => [header, row[index]]))
  ));
  const byTab = Object.fromEntries(objects.map((row) => [row.tab, row]));

  for (const job of manifest) {
    const outputHeaders = new Set(api.headersForJob(job));
    const keys = String(byTab[job.tab].keys || '').split(',').map((key) => key.trim()).filter(Boolean);
    assert.ok(keys.length, `${job.tab} must declare at least one key field`);
    for (const key of keys) {
      assert.equal(outputHeaders.has(key), true, `${job.tab} dictionary key is not an exact header: ${key}`);
    }
  }

  const expected = {
    keywords: [
      'campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id',
      'segments.device', 'segments.ad_network_type',
    ],
    search_terms: [
      'campaign.id', 'ad_group.id', 'search_term_view.search_term',
      'segments.keyword.info.text', 'segments.keyword.info.match_type',
    ],
    ads: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
    ad_inventory: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
    ad_to_lp_map: ['campaign.id', 'ad_group.id', 'ad_id', 'url_source', 'final_url_raw'],
    pmax_assets: [
      'campaign.id', 'asset_group.id', 'asset.id', 'asset_group_asset.field_type',
    ],
    user_list_performance: [
      'scope', 'user_list_resource', 'criterion_id', 'campaign.id', 'ad_group.id',
    ],
  };
  for (const [tab, keyFields] of Object.entries(expected)) {
    assert.deepEqual(byTab[tab].keys.split(', '), keyFields, `${tab} grain contract`);
  }
});

test('field dictionary provides one structured row per output field', () => {
  const { api } = loadExporter();
  const rows = api.buildFieldDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  const objects = rows.slice(1).map((row) => (
    Object.fromEntries(headers.map((header, index) => [header, row[index]]))
  ));
  const byPair = Object.fromEntries(objects.map((row) => [`${row.tab}\u0000${row.field}`, row]));

  assert.equal(byPair['campaign\u0000campaign.id'].data_type, 'identifier');
  assert.equal(byPair['campaign\u0000campaign.id'].is_key, true);
  assert.match(byPair['campaign\u0000metrics.cost_micros'].unit, /micro-units/i);
  assert.match(byPair['campaign\u0000cost'].unit, /account currency/i);
  assert.match(byPair['campaign\u0000campaign.id'].blank_when, /not applicable|unavailable/i);
  assert.equal(byPair['keywords\u0000segments.device'].is_key, true);
});

test('asset-association dictionary entries explicitly warn that metrics are non-additive', () => {
  const { api } = loadExporter();
  const rows = api.buildDataDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  const objects = rows.slice(1).map((row) => (
    Object.fromEntries(headers.map((header, index) => [header, row[index]]))
  ));
  const byTab = Object.fromEntries(objects.map((row) => [row.tab, row]));

  for (const tab of ['rsa_assets', 'demandgen_assets', 'pmax_assets', 'asset_extensions']) {
    assert.match(byTab[tab].google_side_limitations, /non-additive|do not sum/i, `${tab} needs a do-not-sum warning`);
  }
  assert.match(byTab.pmax_assets.google_side_limitations, /pmax_asset_groups|campaign/i);
});

test('dictionary formatting preserves quote-prefixed plain text', () => {
  const harness = createPersistentRichTextHarness({ richTextClearsQuotePrefix: true });
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === '_data_dictionary');
  const rows = api.buildDataDictionaryRows(api.getManifestDefinition());
  const sheet = harness.createSheet('_data_dictionary', rows.map((row) => row.map(() => '')));
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows.map((row) => (
    row.map((value) => typeof value === 'string' && value !== '' ? `'${value}` : value)
  )));

  api.formatReportSheet(sheet, job, 'USD');

  const richValues = sheet.getRange(1, 1, rows.length, rows[0].length).getRichTextValues();
  richValues.forEach((row, rowIndex) => row.forEach((rich, columnIndex) => {
    const expected = String(rows[rowIndex][columnIndex] ?? '');
    assert.equal(rich ? rich.getText() : '', expected);
    if (expected !== '') assert.equal(sheet.ensureCell(rowIndex + 1, columnIndex + 1).quotePrefix, true);
    for (let offset = 0; offset < expected.length; offset += 1) {
      assert.equal(rich.getLinkUrl(offset, offset + 1), null);
    }
  }));
});

test('diagnostic sample redaction removes identifiers, names, emails, queries, and URLs', () => {
  const { api } = loadExporter();
  const redacted = api.redactDiagnosticSample({
    'campaign.id': '101',
    'campaign.name': 'Private Campaign',
    'geographic_view.country_criterion_id': '2840',
    'segments.geo_target_most_specific_location': '1023191',
    'segments.geo_target_state': '21167',
    'change_event.user_email': 'person@example.test',
    'search_term_view.search_term': 'private query',
    'ad_group_ad.ad.final_urls': ['https://private.example.test'],
    'metrics.impressions': '12',
  });

  assert.equal(redacted['campaign.id'], '[REDACTED]');
  assert.equal(redacted['campaign.name'], '[REDACTED]');
  assert.equal(redacted['geographic_view.country_criterion_id'], '[REDACTED]');
  assert.equal(redacted['segments.geo_target_most_specific_location'], '[REDACTED]');
  assert.equal(redacted['segments.geo_target_state'], '[REDACTED]');
  assert.equal(redacted['change_event.user_email'], '[REDACTED]');
  assert.equal(redacted['search_term_view.search_term'], '[REDACTED]');
  assert.equal(redacted['ad_group_ad.ad.final_urls'], '[REDACTED]');
  assert.equal(redacted['metrics.impressions'], '12');
});

test('diagnostics probe both halves of the hybrid PMax asset-group contract', () => {
  const { api } = loadExporter();
  const probes = api.diagnosticProbes({ start: '2026-05-27', end: '2026-08-24' });
  const byName = Object.fromEntries(Array.from(probes, (probe) => [probe.name, probe]));

  assert.ok(byName.pmax_asset_group_inventory);
  assert.ok(byName.pmax_asset_group_metrics);
  assert.doesNotMatch(byName.pmax_asset_group_inventory.query, /metrics\.|segments\.date/);
  assert.match(byName.pmax_asset_group_metrics.query, /metrics\.impressions/);
  assert.match(byName.pmax_asset_group_metrics.query, /segments\.date BETWEEN '2026-05-27' AND '2026-08-24'/);
  assert.doesNotMatch(byName.pmax_signals.query, /asset_group\.status/);
});

test('Preview probes the interactions metric across every custom performance path', () => {
  const { api } = loadExporter();
  const probes = api.diagnosticProbes({ start: '2026-05-27', end: '2026-08-24' });
  const byName = Object.fromEntries(Array.from(probes, (probe) => [probe.name, probe]));

  for (const name of [
    'campaign', 'campaign_geo_raw_ids', 'rsa_assets', 'campaign_audience',
    'ad_group_audience', 'pmax_asset_group_metrics', 'pmax_asset_status',
    'quality_score_metrics', 'asset_extension_metrics',
  ]) {
    assert.ok(byName[name], `missing ${name} probe`);
    assert.match(byName[name].query, /metrics\.interactions/, `${name} must test interactions`);
  }
});

test('source contains public entry points and no client-derived examples', () => {
  const { source } = loadExporter();

  assert.match(source, /function main\s*\(/);
  assert.match(source, /function runDiagnostics\s*\(/);
  assert.match(source, /function resetExportState\s*\(/);
  assert.doesNotMatch(
    source,
    new RegExp(['Northstar Fixture', 'Harbor Sentinel', 'Orchard Placeholder', 'Synthetic Person'].join('|'), 'i'),
  );
});

test('campaign chunks preserve sorted string IDs without numeric precision loss', () => {
  const { api } = loadExporter();
  const chunks = api.chunkCampaignIds(
    ['9007199254740993', '2', '101', '2'],
    2,
  );

  assert.deepEqual(JSON.parse(JSON.stringify(chunks)), [
    ['2', '101'],
    ['9007199254740993'],
  ]);
});

test('runtime declares a handler for every manifest job kind', () => {
  const { api } = loadExporter();
  const supported = new Set(api.supportedJobKinds());
  for (const job of api.getManifestDefinition()) {
    assert.equal(supported.has(job.kind), true, `no runtime handler for ${job.kind} (${job.tab})`);
  }
});

test('identifies text and ID columns for explicit Sheets text formatting', () => {
  const { api } = loadExporter();
  const proximity = api.getManifestDefinition().find((job) => job.tab === 'geo_proximity_targets');
  const indexes = Array.from(api.textColumnIndexes(proximity));
  const headers = Array.from(api.headersForJob(proximity));

  assert.equal(indexes.includes(headers.indexOf('customer.id') + 1), true);
  assert.equal(indexes.includes(headers.indexOf('campaign_criterion.proximity.address.postal_code') + 1), true);
  assert.equal(indexes.includes(headers.indexOf('campaign_criterion.proximity.radius') + 1), false);
});

test('retains the original analytical coverage without overloading mixed grains', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const byTab = Object.fromEntries(manifest.map((job) => [job.tab, job]));

  assert.equal(byTab.ad_to_lp_map.kind, 'ad_to_lp_map');
  assert.equal(byTab.quality_score_keywords.kind, 'quality_score');
  assert.equal(byTab.user_list_performance.kind, 'audience_performance');
  assert.equal(byTab.user_lists, undefined);
  assert.ok(byTab.pmax_asset_group_weekly, 'PMax weekly rows need their own correctly named grain');

  const landingHeaders = Array.from(api.headersForJob(byTab.ad_to_lp_map));
  for (const header of ['final_url_raw', 'final_url_norm', 'domain', 'url_source']) {
    assert.equal(landingHeaders.includes(header), true, `ad_to_lp_map missing ${header}`);
  }

  const assetHeaders = Array.from(api.headersForJob(byTab.asset_extensions));
  for (const header of [
    'customer.id', 'asset.sitelink_asset.description1',
    'asset.structured_snippet_asset.values', 'metrics.impressions', 'cost',
  ]) {
    assert.equal(assetHeaders.includes(header), true, `asset_extensions missing ${header}`);
  }

  const audienceHeaders = Array.from(api.headersForJob(byTab.user_list_performance));
  for (const header of ['scope', 'user_list_resource', 'criterion_id', 'metrics.conversions']) {
    assert.equal(
      audienceHeaders.includes(header),
      true,
      `user_list_performance missing ${header}`,
    );
  }
});

test('normalizes one landing-page row per URL without corrupting commas inside a URL', () => {
  const { api } = loadExporter();
  const rows = api.buildAdToLandingPageRows({
    'customer.id': '123',
    'customer.descriptive_name': 'Example',
    'customer.currency_code': 'USD',
    'customer.time_zone': 'America/New_York',
    'campaign.id': '10',
    'campaign.name': 'Search',
    'ad_group.id': '20',
    'ad_group.name': 'Core',
    'ad_group_ad.ad.id': '30',
    'ad_group_ad.ad.final_urls': '["https://Example.com/path/?utm_source=x", "https://example.com/a,b?x=1"]',
    'ad_group_ad.ad.final_mobile_urls': 'https://m.example.com/mobile/',
    'ad_group_ad.ad.tracking_url_template': '{lpurl}?gclid={gclid}',
  });

  assert.equal(rows.length, 3);
  assert.deepEqual(JSON.parse(JSON.stringify(rows.map((row) => row.slice(9, 12)))), [
    ['https://Example.com/path/?utm_source=x', 'https://example.com/path', 'example.com'],
    ['https://example.com/a,b?x=1', 'https://example.com/a,b?x=1', 'example.com'],
    ['https://m.example.com/mobile/', 'https://m.example.com/mobile', 'm.example.com'],
  ]);
  assert.deepEqual(Array.from(rows, (row) => row[14]), ['FINAL', 'FINAL', 'MOBILE']);
});

test('ad-to-landing-page rows exclude inactive zero-data ads but retain historical activity', () => {
  const { api } = loadExporter();
  const base = {
    'customer.id': '123',
    'campaign.id': '10',
    'campaign.status': 'ENABLED',
    'ad_group.id': '20',
    'ad_group.status': 'ENABLED',
    'ad_group_ad.ad.id': '30',
    'ad_group_ad.ad.final_urls': 'https://example.test/',
  };
  const query = api.buildAdToLandingPageQuery(['10']);

  assert.match(query, /campaign\.status/);
  assert.match(query, /ad_group\.status/);
  assert.match(query, /ad_group_ad\.status/);
  assert.equal(api.buildAdToLandingPageRows({ ...base, 'ad_group_ad.status': 'ENABLED' }).length, 1);
  assert.equal(api.buildAdToLandingPageRows({ ...base, 'ad_group_ad.status': 'PAUSED' }).length, 0);
  assert.equal(api.buildAdToLandingPageRows(
    { ...base, 'ad_group_ad.status': 'REMOVED' },
    { '10|20|30': true },
  ).length, 1);
});

test('asset extensions retain historical active associations and omit all zero-data rows', () => {
  const { api } = loadExporter();
  const range = { start: '2026-05-27', end: '2026-08-24' };
  const definitions = api.buildAssetExtensionQueries(range);

  assert.equal(definitions.length, 3);
  for (const definition of definitions) {
    assert.match(definition.query, /segments\.date BETWEEN '2026-05-27' AND '2026-08-24'/);
    assert.doesNotMatch(definition.query, /\.status = 'ENABLED'/);
    assert.doesNotMatch(definition.query, /metrics\.impressions > 0/);
    assert.equal(api.buildAssetExtensionRow(
      api.getManifestDefinition().find((job) => job.tab === 'asset_extensions'),
      definition,
      { [`${definition.prefix}.status`]: 'PAUSED', 'asset.id': '1' },
    ), null);
    assert.ok(api.buildAssetExtensionRow(
      api.getManifestDefinition().find((job) => job.tab === 'asset_extensions'),
      definition,
      {
        [`${definition.prefix}.status`]: 'PAUSED',
        'asset.id': '1',
        'metrics.conversions_value': '2.5',
      },
    ));
  }
});

test('joins Quality Score to performance with the full campaign, ad-group, criterion key', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'quality_score_keywords');
  const staticRows = [
    {
      'customer.id': '1', 'campaign.id': '10', 'campaign.name': 'A',
      'ad_group.id': '100', 'ad_group.name': 'One',
      'ad_group_criterion.criterion_id': '777',
      'ad_group_criterion.keyword.text': 'plumber',
      'ad_group_criterion.quality_info.quality_score': '8',
    },
    {
      'customer.id': '1', 'campaign.id': '11', 'campaign.name': 'B',
      'ad_group.id': '101', 'ad_group.name': 'Two',
      'ad_group_criterion.criterion_id': '777',
      'ad_group_criterion.keyword.text': 'plumber',
      'ad_group_criterion.quality_info.quality_score': '3',
    },
  ];
  const metricRows = [
    {
      'campaign.id': '10', 'ad_group.id': '100',
      'ad_group_criterion.criterion_id': '777',
      'metrics.impressions': '10', 'metrics.clicks': '2', 'metrics.interactions': '4',
      'metrics.cost_micros': '4000000', 'metrics.conversions': '1',
    },
    {
      'campaign.id': '11', 'ad_group.id': '101',
      'ad_group_criterion.criterion_id': '777',
      'metrics.impressions': '20', 'metrics.clicks': '0', 'metrics.interactions': '0',
      'metrics.cost_micros': '0', 'metrics.conversions': '0',
    },
  ];

  const rows = api.buildQualityScoreRows(job, staticRows, metricRows);
  const headers = Array.from(api.headersForJob(job));
  const byAdGroup = Object.fromEntries(rows.map((row) => [row[headers.indexOf('ad_group.id')], row]));

  assert.equal(rows.length, 2);
  assert.equal(byAdGroup['100'][headers.indexOf('metrics.impressions')], 10);
  assert.equal(byAdGroup['100'][headers.indexOf('ad_group_criterion.quality_info.quality_score')], 8);
  assert.equal(byAdGroup['100'][headers.indexOf('metrics.interactions')], 4);
  assert.equal(byAdGroup['100'][headers.indexOf('conversion_rate')], 0.25);
  assert.equal(byAdGroup['101'][headers.indexOf('metrics.impressions')], 20);
  assert.equal(byAdGroup['101'][headers.indexOf('ad_group_criterion.quality_info.quality_score')], 3);
  assert.equal(byAdGroup['101'][headers.indexOf('metrics.average_cpc')], '');
  assert.equal(byAdGroup['101'][headers.indexOf('average_cpc')], '');
  assert.equal(byAdGroup['101'][headers.indexOf('conversion_rate')], '');
});

test('builds every declarative query without undefined values and keeps workbook names valid', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const ranges = api.buildFrozenRanges('2026-08-25');

  for (const job of manifest) {
    const headers = Array.from(api.headersForJob(job));
    assert.equal(new Set(headers).size, headers.length, `duplicate headers in ${job.tab}`);
    assert.ok(job.tab.length <= 100, `final tab name too long: ${job.tab}`);
    assert.ok(api.stageSheetName(job.tab).length <= 100, `stage tab name too long: ${job.tab}`);
    assert.ok(api.backupSheetName(job.tab).length <= 100, `backup tab name too long: ${job.tab}`);
    if (job.kind !== 'gaql') continue;
    const query = api.buildGaqlQuery(job, ranges, job.chunked ? ['1', '9007199254740993'] : null);
    assert.doesNotMatch(query, /undefined|null/);
    assert.match(query, new RegExp(`FROM ${job.resource}`));
  }
});

test('keeps derived-tab dependencies ordered before both dictionaries', () => {
  const { api } = loadExporter();
  const ids = api.getManifestDefinition().map((job) => job.id);

  assert.ok(ids.indexOf('ad_inventory') < ids.indexOf('ad_to_lp_map'));
  for (const source of [
    'neg_keywords_campaign', 'neg_keywords_ad_group', 'neg_keywords_shared',
    'neg_keyword_shared_links', 'neg_keyword_account_links',
  ]) {
    assert.ok(ids.indexOf(source) < ids.indexOf('negative_keywords_all'));
  }
  assert.ok(ids.indexOf('pmax_asset_groups') < ids.indexOf('pmax_audience_signals'));
  assert.equal(ids.includes('user_lists'), false);
  assert.deepEqual(Array.from(ids.slice(-2)), ['data_dictionary', 'field_dictionary']);
});

test('compacts worst-case resume metadata below the Google Sheets cell limit', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const state = api.createRunState(
    {
      version: api.VERSION,
      accountId: '123-456-7890',
      spreadsheetId: 'x'.repeat(64),
      configSignature: 'y'.repeat(180),
    },
    Date.now(),
    api.buildFrozenRanges('2026-08-25'),
    manifest.map((job) => job.id),
  );
  state.accountName = 'A'.repeat(120);
  for (const job of manifest) {
    state.tabs[job.tab] = {
      status: 'ERROR_PREVIOUS_PRESERVED', rows: 123456789, durationMs: 999999999,
      error: 'E'.repeat(500), limitation: 'L'.repeat(500),
      partialLimited: true, priorPreserved: true,
    };
  }

  const compact = api.compactStateForStorage(state, 45_000);
  assert.ok(JSON.stringify(compact).length <= 45_000);
  assert.equal(state.tabs.campaign.error.length, 500, 'compaction must not mutate in-memory metadata');
  assert.deepEqual(Array.from(compact.manifest), Array.from(state.manifest));
});

test('stores resumable state in the owned workbook and never calls unsupported PropertiesService', () => {
  const { source } = loadExporter();

  assert.doesNotMatch(source, /PropertiesService|STATE_PROPERTY_KEY/);
  assert.match(source, /function saveStateSheetRuntime_/);
  assert.match(source, /function loadStateSheetRuntime_/);
  assert.match(source, /state_json/);
});
