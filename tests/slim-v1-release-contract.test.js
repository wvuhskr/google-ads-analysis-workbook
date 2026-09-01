'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  loadExporter,
  createPersistentRichTextHarness,
} = require('./load-exporter');

const NOW_MS = Date.UTC(2026, 7, 26, 12, 0, 0);
const RELEASE_CANDIDATE_LANGUAGE = new RegExp(
  ['R' + 'C15', 'release' + ' candidate'].join('|'),
  'i',
);

function dictionaryObjects(api) {
  const rows = api.buildDataDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  return Object.fromEntries(rows.slice(1).map((row) => [
    row[0],
    Object.fromEntries(headers.map((header, index) => [header, row[index]])),
  ]));
}

function fieldDictionaryObject(api, tab, field) {
  const rows = api.buildFieldDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  const tabIndex = headers.indexOf('tab');
  const fieldIndex = headers.indexOf('field');
  const row = rows.slice(1).find((candidate) => (
    candidate[tabIndex] === tab && candidate[fieldIndex] === field
  ));
  assert.ok(row, `missing field dictionary row for ${tab}.${field}`);
  return Object.fromEntries(headers.map((header, index) => [header, row[index]]));
}

function completedState(api, status = 'COMPLETE') {
  const state = api.createRunState({
    version: api.VERSION,
    accountId: '123-456-7890',
    spreadsheetId: 'sheet-123',
    configSignature: 'config-abc',
  }, NOW_MS, {}, []);
  state.status = status;
  return state;
}

test('terminal workbooks explain that the native Google Sheet is the final deliverable', () => {
  const { api } = loadExporter();

  for (const status of ['COMPLETE', 'COMPLETE_WITH_LIMITATIONS']) {
    const state = completedState(api, status);
    const rows = api.buildExportInfoRows(state, []);
    const nextAction = rows.find((row) => row[0] === 'next_action');
    assert.ok(nextAction, `${status} is missing next_action metadata`);
    const instruction = String(nextAction[1] || '');
    assert.match(instruction, /workbook|Google Sheets/i);
    assert.match(instruction, /XLSX.*sanitiz|sanitiz.*XLSX/i);
    assert.doesNotMatch(instruction, /distribution|second function|change.*mode/i);
    if (status === 'COMPLETE') {
      assert.match(instruction, /ready for analysis/i);
    } else {
      assert.match(instruction, /review every LIMITED row/i);
    }
  }

  const failed = completedState(api);
  failed.status = 'COMPLETE_WITH_ERRORS';
  const failedInstruction = api.buildExportInfoRows(failed, [])
    .find((row) => row[0] === 'next_action')[1];
  assert.match(String(failedInstruction), /finished with errors|not.*ready/i);
  assert.match(String(failedInstruction), /review.*ERROR/i);
  assert.doesNotMatch(String(failedInstruction), /XLSX|distribution/i);
});

test('live main has no mode switch or second artifact-creation entry point', () => {
  const { api, context } = loadExporter();

  assert.equal(Object.prototype.hasOwnProperty.call(context.CONFIG, 'RUN_MODE'), false);
  assert.equal(typeof context.exportSanitizedXlsx, 'undefined');
  assert.equal(Object.prototype.hasOwnProperty.call(context.CONFIG, 'AUTOMATIC_XLSX_MAX_BYTES'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(context.CONFIG, 'MIN_DISTRIBUTION_REMAINING_SECONDS'), false);
  const state = completedState(api);
  assert.equal(Object.keys(state).some((key) => /^distribution/i.test(key)), false);
});

test('stable v1.0.0 keeps schema 9, runtime contract 10, and 41-sheet topology', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();

  assert.equal(api.VERSION, 'v1.0.0');
  assert.equal(api.OUTPUT_SCHEMA_VERSION, 9);
  assert.equal(api.RUNTIME_CONTRACT_VERSION, 10);
  assert.equal(manifest.length, 39);
  const order = Array.from(api.preferredTabOrder(manifest));
  assert.equal(order.length, 41);
  assert.deepEqual(order.slice(0, 4), [
    'START_HERE', '_export_info', '_data_dictionary', '_field_dictionary',
  ]);

  const identity = {
    version: api.VERSION,
    accountId: '123-456-7890',
    spreadsheetId: 'sheet-123',
    configSignature: 'config-abc',
  };
  const state = api.createRunState(identity, NOW_MS, {}, []);
  assert.equal(state.outputSchemaVersion, 9);
  assert.equal(state.runtimeContractVersion, 10);

  state.runtimeContractVersion = 5;
  assert.throws(
    () => api.assertStateCompatible(state, identity, NOW_MS + 1_000, 24),
    /runtime contract/i,
  );
});

test('START_HERE is cataloged without entering the 914-row field dictionary', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const catalog = api.buildDataDictionaryRows(manifest).map((row) => Array.from(row));
  const fieldRows = api.buildFieldDictionaryRows(manifest).map((row) => Array.from(row));
  const start = catalog.find((row) => row[0] === 'START_HERE');

  assert.ok(start);
  assert.match(String(start[1]), /summary|start|navigation/i);
  assert.equal(start[8], '');
  assert.equal(fieldRows.length - 1, 914);
  assert.equal(fieldRows.slice(1).some((row) => row[0] === 'START_HERE'), false);
});

test('stable final workbook contract has exactly 41 ordered visible sheets', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  class Sheet {
    constructor(name) { this.name = name; this.tabColor = ''; }
    getName() { return this.name; }
    getLastRow() { return 1; }
    getLastColumn() { return 1; }
    setTabColor(value) { this.tabColor = value; return this; }
  }
  class Workbook {
    constructor(names) {
      this.sheets = names.map((name) => new Sheet(name));
      this.active = this.sheets[0];
    }
    getSheets() { return this.sheets.slice(); }
    getSheetByName(name) { return this.sheets.find((sheet) => sheet.name === name) || null; }
    deleteSheet(sheet) { this.sheets = this.sheets.filter((candidate) => candidate !== sheet); }
    setActiveSheet(sheet) { this.active = sheet; return sheet; }
    moveActiveSheet(position) {
      this.sheets = this.sheets.filter((sheet) => sheet !== this.active);
      this.sheets.splice(position - 1, 0, this.active);
    }
  }
  const workbook = new Workbook(Array.from(api.preferredTabOrder(manifest)).reverse());

  const names = Array.from(api.finalizeWorkbookLayout(workbook, manifest, []));

  assert.equal(names.length, 41);
  assert.equal(names[0], 'START_HERE');
  assert.deepEqual(names.slice(1, 4), ['_export_info', '_data_dictionary', '_field_dictionary']);
  assert.equal(
    names.some((name) => name === '_export_state' || name.startsWith('__gads_export_')),
    false,
  );
  assert.equal(api.buildFieldDictionaryRows(manifest).length - 1, 914);
  assert.equal(workbook.active.getName(), 'START_HERE');
  assert.equal(workbook.sheets.every((sheet) => Boolean(sheet.tabColor)), true);
  const representativeColors = [
    'campaign', 'campaign_inventory', 'rsa_assets', 'geo_targets',
    'negative_keywords_all', 'user_list_performance', 'change_history',
  ].map((name) => workbook.getSheetByName(name).tabColor);
  assert.equal(new Set(representativeColors).size, representativeColors.length);
});

test('manifest engine records source-read boundaries and exposes them per tab', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign');
  const state = api.createRunState({
    version: api.VERSION,
    accountId: '123-456-7890',
    spreadsheetId: 'sheet-123',
    configSignature: 'config-abc',
  }, NOW_MS, {
    aggregate: { start: '2026-05-28', end: '2026-08-25' },
    weekly: { start: '2026-05-28', end: '2026-08-25' },
    change: { start: '2026-07-29', end: '2026-08-25' },
  }, [job.id]);
  let clock = NOW_MS;
  const adapter = {
    remainingSeconds() { return 1_000; },
    saveState() {},
    writeInfo() {},
    startJob() { return 'stage-campaign'; },
    getChunkCount() { return 2; },
    getChunkStartRow() { return 2; },
    rollbackChunk() {},
    runChunk() { clock += 1_000; return 1; },
    commitJob() {},
    abortJob() {},
    finalizeWorkbook() {},
    publishWorkbook() {},
    hasPriorFinal() { return false; },
    clearState() {},
    nowMs() { clock += 100; return clock; },
  };

  const result = api.runManifestEngine(state, [job], adapter, 180);
  const tab = result.tabs.campaign;
  assert.equal(result.status, 'COMPLETE');
  assert.ok(tab.sourceReadStartedAtMs >= NOW_MS);
  assert.ok(tab.sourceReadCompletedAtMs >= tab.sourceReadStartedAtMs);

  const infoRows = api.buildExportInfoRows(result, [job]);
  const header = infoRows.find((row) => row[0] === 'tab');
  const campaign = infoRows.find((row) => row[0] === 'campaign');
  const startedColumn = header.indexOf('source_read_started_at');
  const completedColumn = header.indexOf('source_read_completed_at');
  assert.equal(campaign[startedColumn], new Date(tab.sourceReadStartedAtMs).toISOString());
  assert.equal(campaign[completedColumn], new Date(tab.sourceReadCompletedAtMs).toISOString());
});

test('_export_info explains that tabs are sequential non-atomic snapshots', () => {
  const { api } = loadExporter();
  const rows = api.buildExportInfoRows(completedState(api), []);
  const row = rows.find((candidate) => candidate[0] === 'snapshot_semantics');

  assert.ok(row, 'missing snapshot_semantics metadata row');
  const warning = String(row[1] || '');
  assert.match(warning, /sequential/i);
  assert.match(warning, /non[- ]atomic|not (?:an )?atomic/i);
  assert.match(warning, /campaign.*authoritative|authoritative.*campaign|campaign.*baseline/i);
});

test('date-only fields and Google Ads datetimes retain explicit text semantics', () => {
  const { api } = loadExporter();
  const weekly = fieldDictionaryObject(api, 'campaign_weekly', 'segments.week');
  const inventory = fieldDictionaryObject(api, 'campaign_inventory', 'campaign.start_date_time');
  const change = fieldDictionaryObject(api, 'change_history', 'change_date_time');

  assert.equal(weekly.data_type, 'iso_date_text');
  assert.match(String(weekly.unit), /YYYY-MM-DD|ISO 8601/i);
  for (const metadata of [inventory, change]) {
    assert.equal(metadata.data_type, 'google_ads_datetime_text');
    assert.match(String(metadata.unit), /customer\.time_zone|account time zone/i);
    assert.match(String(metadata.unit), /no (?:embedded )?(?:UTC )?offset|without.*offset/i);
  }
});

test('data dictionaries retain DSA and sparse-week interpretation warnings', () => {
  const { api } = loadExporter();
  const dictionary = dictionaryObjects(api);
  const dsaWarning = String(dictionary.ad_to_lp_map.google_side_limitations);
  assert.match(dsaWarning, /Dynamic Search Ad|\bDSA\b/i);
  assert.match(dsaWarning, /landing_pages/i);

  for (const tab of ['campaign_weekly', 'ad_group_weekly', 'pmax_asset_group_weekly']) {
    const warning = String(dictionary[tab].google_side_limitations);
    const dateRange = String(dictionary[tab].date_range);
    assert.match(dateRange, /same.*90 complete days|90 complete days.*same/i);
    assert.doesNotMatch(dateRange, /13 complete|full weeks/i);
    assert.match(warning, /partial.*(?:first|boundary).*week|boundary weeks.*partial/i);
    assert.match(warning, /missing|sparse|fewer/i);
    assert.match(warning, /zero|no metric row|no activity/i);
    assert.match(warning, /not .*export (?:failure|error)|does not (?:mean|indicate).*fail/i);
  }
});

test('change history discloses the Google 30-day source limit behind its 28-day window', () => {
  const { api } = loadExporter();
  const dictionary = dictionaryObjects(api);

  assert.match(String(dictionary.change_history.date_range), /28 complete days/i);
  assert.match(
    String(dictionary.change_history.google_side_limitations),
    /Google.*(?:past|within).*30 days|30-day.*Google/i,
  );
});

test('field dictionary keeps long provenance columns readable', () => {
  const harness = createPersistentRichTextHarness();
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === '_field_dictionary');
  const rows = api.buildFieldDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  const sheet = harness.createSheet(
    '_field_dictionary',
    rows.slice(0, 2).map((row) => row.map(() => '')),
  );
  sheet.getRange(1, 1, 2, headers.length).setValues(rows.slice(0, 2).map((row) => (
    Array.from(row, (value) => (typeof value === 'string' && value !== '' ? `'${value}` : value))
  )));

  api.formatReportSheet(sheet, job, 'USD');

  function width(header) {
    return sheet.columnWidths[headers.indexOf(header) + 1];
  }
  assert.ok(width('field') >= 240);
  assert.ok(width('source_fields') >= 280);
  assert.ok(width('unit') >= 240);
  assert.ok(width('derivation') >= 320);
  assert.ok(width('blank_when') >= 280);
});

test('stable public source removes release-candidate language without changing workbook behavior', () => {
  const { source, context } = loadExporter();

  assert.match(source, /Google Ads Analysis Workbook/);
  assert.doesNotMatch(source, RELEASE_CANDIDATE_LANGUAGE);
  assert.equal(context.VERSION, 'v1.0.0');
  assert.equal(context.OUTPUT_SCHEMA_VERSION, 9);
  assert.equal(context.RUNTIME_CONTRACT_VERSION, 10);
});
