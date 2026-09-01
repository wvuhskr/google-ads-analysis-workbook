'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadExporter } = require('./load-exporter');

const SPREADSHEET_ID = 'sheet-1';
const ACCOUNT_ID = '123-456-7890';

function spreadsheetUrl() {
  return ['https://docs.google.com', 'spreadsheets', 'd', SPREADSHEET_ID, 'edit'].join('/');
}

function matrixSlice(source, row, column, rowCount, columnCount, stringify) {
  return Array.from({ length: rowCount }, (_unused, rowOffset) => (
    Array.from({ length: columnCount }, (_unusedColumn, columnOffset) => {
      const sourceRow = source[row - 1 + rowOffset] || [];
      const value = sourceRow[column - 1 + columnOffset];
      if (value === undefined || value === null) return '';
      return stringify ? String(value) : value;
    })
  ));
}

class FixtureSheet {
  constructor(name, rows = [], options = {}) {
    this.name = name;
    this.rows = rows.map((row) => Array.from(row));
    this.displayRows = (options.displayRows || rows).map((row) => Array.from(row));
    this.physicalLastRow = options.physicalLastRow;
    this.hidden = Boolean(options.hidden);
    this.drawings = options.drawings || [];
  }

  getName() { return this.name; }

  getLastRow() {
    return this.physicalLastRow === undefined ? this.rows.length : this.physicalLastRow;
  }

  getLastColumn() {
    return this.rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
  }

  getMaxRows() { return Math.max(1, this.getLastRow()); }

  getMaxColumns() { return Math.max(1, this.getLastColumn()); }

  isSheetHidden() { return this.hidden; }

  getDrawings() { return this.drawings.slice(); }

  getRange(row, column, rowCount = 1, columnCount = 1) {
    const values = matrixSlice(this.rows, row, column, rowCount, columnCount, false);
    const displayValues = matrixSlice(
      this.displayRows,
      row,
      column,
      rowCount,
      columnCount,
      true,
    );
    return {
      getValues() { return values.map((candidate) => candidate.slice()); },
      getDisplayValues() { return displayValues.map((candidate) => candidate.slice()); },
      getDisplayValue() { return displayValues[0][0]; },
    };
  }
}

class FixtureWorkbook {
  constructor(sheets) { this.sheets = sheets.slice(); }

  getSheets() { return this.sheets.slice(); }

  getSheetByName(name) {
    return this.sheets.find((sheet) => sheet.getName() === name) || null;
  }
}

function terminalState(api, status = 'COMPLETE') {
  const state = api.createRunState({
    version: api.VERSION,
    accountId: ACCOUNT_ID,
    spreadsheetId: SPREADSHEET_ID,
    configSignature: 'config-1',
  }, Date.UTC(2026, 7, 28, 12), {
    aggregate: { start: '2026-05-30', end: '2026-08-27' },
    weekly: { start: '2026-05-30', end: '2026-08-27' },
    change: { start: '2026-07-31', end: '2026-08-27' },
  }, []);
  state.status = status;
  state.accountName = 'Example Account';
  state.accountCurrencyCode = 'USD';
  return state;
}

function nativeInfoRows(api, manifest) {
  const rows = [
    [
      api.OWNER_KEY,
      api.VERSION,
      'output_schema_version',
      api.OUTPUT_SCHEMA_VERSION,
      'Confidential Google Ads account export',
      '',
      '',
      '',
    ],
    [
      'run_id',
      'run-1787928000000',
      'overall_status',
      'COMPLETE',
      'started_at',
      '2026-08-28T16:00:00.000Z',
      'workbook_status',
      'READY',
    ],
    [
      'account_id',
      ACCOUNT_ID,
      'account_name',
      'Example Account',
      'updated_at',
      '2026-08-28T16:05:00.000Z',
      'deliverable_type',
      'NATIVE_GOOGLE_SHEET',
    ],
    [
      'aggregate_range',
      '2026-05-30 through 2026-08-27',
      'weekly_range',
      '2026-05-30 through 2026-08-27',
      'change_range',
      '2026-07-31 through 2026-08-27',
      'last_complete_day',
      '2026-08-27',
    ],
    [
      'current_job', '', 'job_index', manifest.length, 'manifest_jobs', manifest.length,
      'workbook_url', spreadsheetUrl(),
    ],
    [
      'next_action',
      'The native Google Sheets workbook is complete with limited coverage.',
      'workbook_grid_cells',
      2500000,
      'cell_safety_headroom',
      6500000,
      'reporting_window',
      'LAST_90_COMPLETE_DAYS',
    ],
    [
      'snapshot_semantics',
      'Tabs are sequential, non-atomic query snapshots.',
      'runtime_contract_version',
      api.RUNTIME_CONTRACT_VERSION,
      'checkpoint_schema_version',
      1,
      'refresh_behavior',
      'Run main() again after completion to refresh the export.',
    ],
    [
      'tab',
      'status',
      'rows',
      'duration_seconds',
      'source_read_started_at',
      'source_read_completed_at',
      'prior_data_preserved',
      'limitation_or_error',
    ],
  ];

  manifest.forEach((job) => {
    const isSearchTerms = job.tab === 'search_terms';
    rows.push([
      job.tab,
      'OK',
      isSearchTerms ? 54321 : 0,
      1,
      '2026-08-28T16:00:00.000Z',
      '2026-08-28T16:00:01.000Z',
      'NO',
      '',
    ]);
  });
  return rows;
}

function nativeWorkbookFixture(api) {
  const manifest = api.getManifestDefinition();
  const infoRows = nativeInfoRows(api, manifest);
  const displayInfoRows = infoRows.map((row) => row.map((value) => String(value ?? '')));
  displayInfoRows.find((row) => row[0] === 'search_terms')[2] = '54,321';

  const sheetsByName = {
    START_HERE: new FixtureSheet('START_HERE', [['Google Ads Analysis Workbook']]),
    _export_info: new FixtureSheet('_export_info', infoRows, { displayRows: displayInfoRows }),
  };
  manifest.forEach((job) => {
    const rows = [Array.from(api.headersForJob(job))];
    sheetsByName[job.tab] = new FixtureSheet(job.tab, rows, {
      physicalLastRow: job.tab === 'search_terms' ? 54322 : undefined,
    });
  });
  const outputSheets = Array.from(
    api.preferredTabOrder(manifest),
    (name) => sheetsByName[name],
  );
  const checkpoint = new FixtureSheet(
    '_export_state',
    [[api.OWNER_KEY, api.VERSION], ['run_id', 'run-1787928000000']],
    { hidden: true },
  );
  return {
    manifest,
    workbook: new FixtureWorkbook(outputSheets.concat(checkpoint)),
  };
}

function nonManagerReport() {
  let read = false;
  return {
    rows() {
      return {
        hasNext() { return !read; },
        next() { read = true; return { 'customer.manager': false }; },
      };
    },
  };
}

test('new checkpoints contain no distribution or XLSX artifact state', () => {
  const { api } = loadExporter();
  const state = terminalState(api);

  assert.deepEqual(
    Object.keys(state).filter((key) => /^distribution/i.test(key)),
    [],
  );
});

test('_export_info identifies native readiness and separates optional local XLSX sanitization', () => {
  const { api } = loadExporter();

  for (const status of ['COMPLETE', 'COMPLETE_WITH_LIMITATIONS']) {
    const rows = api.buildExportInfoRows(terminalState(api, status), []);
    const instruction = String(rows.find((row) => row[0] === 'next_action')[1] || '');
    const flattened = rows.flat().map(String);

    assert.match(instruction, /Google Sheets? workbook/i, status);
    assert.match(instruction, /complete/i, status);
    if (status === 'COMPLETE') assert.match(instruction, /ready for analysis/i);
    if (status === 'COMPLETE_WITH_LIMITATIONS') {
      assert.match(instruction, /limited coverage/i);
      assert.match(instruction, /review every LIMITED row/i);
    }
    assert.match(instruction, /XLSX.*sanitiz|sanitiz.*XLSX/i, status);
    assert.doesNotMatch(instruction, /distribution|artifact retry|retry.*artifact/i, status);
    assert.equal(
      flattened.some((value) => /^distribution_/i.test(value)),
      false,
      status,
    );
  }
});

test('_export_info keeps the workbook non-ready during final validation', () => {
  const { api } = loadExporter();
  const rows = api.buildExportInfoRows(terminalState(api, 'FINALIZING'), []);

  assert.equal(rows[1][3], 'FINALIZING');
  assert.equal(rows[1][7], 'IN_PROGRESS');
  assert.equal(rows[0][3], api.OUTPUT_SCHEMA_VERSION);
  assert.equal(rows[6][3], api.RUNTIME_CONTRACT_VERSION);
  assert.match(String(rows[5][1]), /finaliz|resume/i);
});

test('START_HERE identifies the native Google Sheet and has no conversion status', () => {
  const { api } = loadExporter();
  const state = terminalState(api);
  const model = api.buildStartHereModel(state, [], api.getManifestDefinition());
  const metadata = api.buildStartHereRows(model)[2].map(String);

  assert.equal(Object.prototype.hasOwnProperty.call(model.account, 'distributionStatus'), false);
  assert.equal(metadata.includes('Deliverable: This Google Sheet'), true);
  assert.equal(metadata.some((value) => /^Data status: Complete$/i.test(value)), true);
  assert.equal(metadata.some((value) => /XLSX status|distribution|conversion snapshot/i.test(value)), false);
});

test('Preview runs account diagnostics without ScriptApp, UrlFetchApp, or DriveApp', () => {
  const events = [];
  const { context } = loadExporter({
    SpreadsheetApp: {
      openById(id) {
        assert.equal(id, SPREADSHEET_ID);
        return new FixtureWorkbook([new FixtureSheet('Sheet1')]);
      },
    },
    AdsApp: {
      getExecutionInfo() { return { isPreview() { return true; } }; },
      currentAccount() { return { getTimeZone() { return 'America/New_York'; } }; },
      report(query) {
        const accountCheck = query.includes('customer.manager');
        events.push(accountCheck ? 'account' : 'diagnostic');
        let read = false;
        return {
          rows() {
            return {
              hasNext() { return accountCheck && !read; },
              next() { read = true; return { 'customer.manager': false }; },
            };
          },
        };
      },
    },
    Utilities: { formatDate() { return '2026-08-28'; } },
    Logger: { log() {} },
  });
  context.CONFIG.SPREADSHEET_URL =
    spreadsheetUrl();

  const results = context.main();

  assert.equal(events[0], 'account');
  assert.equal(events.includes('diagnostic'), true);
  assert.equal(results.length > 0, true);
  assert.equal(
    results.some((result) => result.probe === 'automatic_xlsx_transport'),
    false,
  );
});

test('resetExportState refuses Preview before locking or opening the workbook', () => {
  let lockAttempts = 0;
  let workbookOpens = 0;
  const { context } = loadExporter({
    AdsApp: {
      getExecutionInfo() { return { isPreview() { return true; } }; },
    },
    LockService: {
      getScriptLock() { lockAttempts += 1; throw new Error('LOCK_REACHED'); },
    },
    SpreadsheetApp: {
      openById() { workbookOpens += 1; throw new Error('WORKBOOK_REACHED'); },
    },
  });
  context.CONFIG.SPREADSHEET_URL =
    spreadsheetUrl();
  context.CONFIG.ALLOW_RESET = true;

  assert.throws(() => context.resetExportState(), /Preview|Run/i);
  assert.equal(lockAttempts, 0);
  assert.equal(workbookOpens, 0);
});

test('a fresh live export reaches the campaign query without XLSX runtime services', () => {
  const reportQueries = [];
  const blank = new FixtureSheet('Sheet1');
  const workbook = new FixtureWorkbook([blank]);
  const { context } = loadExporter({
    SpreadsheetApp: {
      openById(id) { assert.equal(id, SPREADSHEET_ID); return workbook; },
    },
    Utilities: {
      formatDate() { return '2026-08-28'; },
      sleep() {},
    },
    AdsApp: {
      getExecutionInfo() {
        return { isPreview() { return false; }, getRemainingTime() { return 1000; } };
      },
      currentAccount() {
        return {
          getCustomerId() { return ACCOUNT_ID; },
          getName() { return 'Example Account'; },
          getCurrencyCode() { return 'USD'; },
          getTimeZone() { return 'America/New_York'; },
        };
      },
      report(query) {
        reportQueries.push(query);
        if (query.includes('customer.manager')) return nonManagerReport();
        throw new Error('SENTINEL_CAMPAIGN_QUERY_REACHED');
      },
    },
    Logger: { log() {} },
  });
  context.CONFIG.SPREADSHEET_URL =
    spreadsheetUrl();

  assert.throws(() => context.main(), /SENTINEL_CAMPAIGN_QUERY_REACHED/);
  assert.equal(reportQueries.length, 2);
  assert.match(reportQueries[0], /customer\.manager/);
  assert.match(reportQueries[1], /campaign\.id/);
});

test('a completed-workbook rerun fails closed before and during checkpoint creation', () => {
  const events = [];
  const infoRowsByStatus = new Map();
  const workbook = { getSheets() { return []; } };
  const { api, context } = loadExporter({
    SpreadsheetApp: { flush() { events.push('flush'); } },
    Utilities: { formatDate() { return '2026-08-28'; } },
    AdsApp: {
      getExecutionInfo() { return { isPreview() { return false; } }; },
      currentAccount() {
        return {
          getCustomerId() { return ACCOUNT_ID; },
          getName() { return 'Example Account'; },
          getCurrencyCode() { return 'USD'; },
          getTimeZone() { return 'America/New_York'; },
        };
      },
    },
    Logger: { log() {} },
  });
  context.CONFIG.SPREADSHEET_URL =
    spreadsheetUrl();
  context.validateRuntimeConfig_ = () => {};
  context.extractSpreadsheetId_ = () => SPREADSHEET_ID;
  context.openSpreadsheetRuntime_ = () => workbook;
  context.summarizeWorkbookRuntime_ = () => [
    { name: '_export_info', marker: api.OWNER_KEY, blank: false },
  ];
  context.assertWorkbookOwnership_ = () => 'owned';
  context.getManifestDefinition_ = () => [];
  context.validateManifest_ = () => {};
  context.materialConfigSignature_ = () => 'config-1';
  context.loadStateSheetRuntime_ = () => null;
  context.assertOwnedWorkbookRestartIdentityRuntime_ = () => true;
  context.assertAdvertiserAccountRuntime_ = () => true;
  context.fetchCampaignIdsRuntime_ = () => ['1'];
  context.writeExportInfoRuntime_ = (_spreadsheet, state, manifest) => {
    events.push(`info:${state.status}`);
    infoRowsByStatus.set(
      state.status,
      api.buildExportInfoRows(state, manifest),
    );
  };
  context.writeStartHereRuntime_ = (_spreadsheet, state, _manifest, options) => {
    assert.equal(options.progressOnly, true);
    events.push(`progress:${state.status}`);
  };
  context.writeCampaignStateRuntime_ = () => {
    events.push('checkpoint');
    throw new Error('SENTINEL_CHECKPOINT_WRITE_FAILED');
  };

  assert.throws(() => context.main(), /SENTINEL_CHECKPOINT_WRITE_FAILED/);
  assert.deepEqual(events, [
    'info:RUNNING',
    'flush',
    'progress:RUNNING',
    'flush',
    'checkpoint',
    'info:COMPLETE_WITH_ERRORS',
    'flush',
    'progress:COMPLETE_WITH_ERRORS',
    'flush',
  ]);

  const runningRows = infoRowsByStatus.get('RUNNING');
  const runningAction = String(runningRows.find((row) => row[0] === 'next_action')[1]);
  assert.match(runningAction, /missing.*checkpoint|checkpoint.*missing/i);
  assert.match(runningAction, /resetExportState|ALLOW_RESET/i);

  const failedRows = infoRowsByStatus.get('COMPLETE_WITH_ERRORS');
  assert.equal(failedRows[1][3], 'COMPLETE_WITH_ERRORS');
  assert.notEqual(failedRows[1][7], 'READY');
  const failedAction = String(failedRows.find((row) => row[0] === 'next_action')[1]);
  assert.match(failedAction, /no resumable checkpoint|resetExportState|ALLOW_RESET/i);
});

test('a resumable live export reaches checkpoint validation without XLSX runtime services', () => {
  const seed = loadExporter().api;
  const manifest = seed.getManifestDefinition();
  const state = seed.createRunState({
    version: seed.VERSION,
    outputSchemaVersion: seed.OUTPUT_SCHEMA_VERSION,
    runtimeContractVersion: seed.RUNTIME_CONTRACT_VERSION,
    accountId: ACCOUNT_ID,
    spreadsheetId: SPREADSHEET_ID,
    configSignature: seed.materialConfigSignature(manifest),
  }, Date.now(), {
    aggregate: { start: '2026-05-30', end: '2026-08-27' },
    weekly: { start: '2026-05-30', end: '2026-08-27' },
    change: { start: '2026-07-31', end: '2026-08-27' },
  }, ['deliberately-wrong']);
  state.status = 'PAUSED';

  const cells = new Map([
    ['1:1', seed.OWNER_KEY],
    ['2:2', state.runId],
    ['4:1', 'state_json'],
    ['4:2', JSON.stringify(state)],
  ]);
  const checkpoint = {
    getName() { return '_export_state'; },
    getLastRow() { return 4; },
    getLastColumn() { return 2; },
    getRange(row, column) {
      return { getDisplayValue() { return String(cells.get(`${row}:${column}`) || ''); } };
    },
    isSheetHidden() { return true; },
  };
  const workbook = new FixtureWorkbook([checkpoint]);
  let adsQueryCount = 0;
  const { context } = loadExporter({
    SpreadsheetApp: {
      openById(id) { assert.equal(id, SPREADSHEET_ID); return workbook; },
    },
    Utilities: { sleep() {} },
    AdsApp: {
      getExecutionInfo() { return { isPreview() { return false; } }; },
      currentAccount() {
        return {
          getCustomerId() { return ACCOUNT_ID; },
          getName() { return 'Example Account'; },
          getCurrencyCode() { return 'USD'; },
          getTimeZone() { return 'America/New_York'; },
        };
      },
      report(query) {
        adsQueryCount += 1;
        assert.match(query, /customer\.manager/);
        return nonManagerReport();
      },
    },
    Logger: { log() {} },
  });
  context.CONFIG.SPREADSHEET_URL =
    spreadsheetUrl();

  assert.throws(() => context.main(), /different job manifest/i);
  assert.equal(adsQueryCount, 0, 'incompatible checkpoints must fail before Ads queries');
});

test('resume reads the campaign count as a typed value rather than a formatted display string', () => {
  const { api } = loadExporter({ SpreadsheetApp: { flush() {} } });
  const ids = Array.from({ length: 1234 }, (_unused, index) => String(index + 1));
  const state = { runId: 'run-1' };
  let hidden = false;
  let hideAttempts = 0;
  const sheet = {
    getLastRow() { return ids.length + 4; },
    isSheetHidden() { return hidden; },
    hideSheet() { hideAttempts += 1; hidden = true; },
    getRange(row, column, rowCount = 1) {
      if (row === 1 && column === 1) {
        return { getDisplayValue() { return api.OWNER_KEY; } };
      }
      if (row === 2 && column === 2) {
        return { getDisplayValue() { return state.runId; } };
      }
      if (row === 3 && column === 2) {
        return {
          getValue() { return ids.length; },
          getDisplayValue() { return '1,234'; },
        };
      }
      if (row === 5 && column === 1) {
        return {
          getDisplayValues() {
            return ids.slice(0, rowCount).map((id) => [id]);
          },
        };
      }
      throw new Error(`Unexpected range ${row}:${column}`);
    },
  };
  const workbook = { getSheetByName() { return sheet; } };

  assert.deepEqual(Array.from(api.readCampaignStateRuntime(workbook, state)), ids);
  assert.equal(hidden, true);
  assert.equal(hideAttempts, 1);
  assert.deepEqual(Array.from(api.readCampaignStateRuntime(workbook, state)), ids);
  assert.equal(hideAttempts, 1, 'an already-hidden checkpoint should not be hidden again');
});

test('native finalization refreshes and validates the workbook before clearing its checkpoint', () => {
  const { source } = loadExporter();
  const adapterStart = source.indexOf('function createRuntimeAdapter_(');
  const finalizeStart = source.indexOf('finalizeWorkbook: function(state, manifest)', adapterStart);
  const publishStart = source.indexOf('publishWorkbook: function(state, manifest)', finalizeStart);
  const clearStart = source.indexOf('clearState: function()', publishStart);
  assert.notEqual(adapterStart, -1, 'missing runtime adapter');
  assert.notEqual(finalizeStart, -1, 'missing native finalization');
  assert.notEqual(publishStart, -1, 'missing terminal native publication');
  assert.notEqual(clearStart, -1, 'missing checkpoint cleanup');

  const finalizeBody = source.slice(finalizeStart, publishStart);
  const startHere = finalizeBody.indexOf('writeStartHereRuntime_(');
  const layout = finalizeBody.indexOf('finalizeWorkbookLayout_(');
  const finalInfo = finalizeBody.indexOf('writeExportInfoRuntime_(');
  const validation = finalizeBody.indexOf('validateNativeWorkbookRuntime_(');
  const flush = finalizeBody.indexOf('SpreadsheetApp.flush()');
  for (const [label, index] of [
    ['START_HERE write', startHere],
    ['layout', layout],
    ['terminal _export_info refresh', finalInfo],
    ['native validation', validation],
    ['flush', flush],
  ]) {
    assert.notEqual(index, -1, `missing ${label}`);
  }
  assert.ok(startHere < layout, 'START_HERE must be written before final layout');
  assert.ok(layout < finalInfo, '_export_info must be refreshed after final layout');
  assert.ok(finalInfo < validation, 'terminal metadata must be refreshed before validation');
  assert.ok(validation < flush, 'validated native state must be flushed before checkpoint cleanup');
  assert.match(finalizeBody, /allowCheckpoint:\s*true,\s*allowFinalizing:\s*true/);

  const publishBody = source.slice(publishStart, source.indexOf('writeFailureSummary:', publishStart));
  const publishValidation = publishBody.indexOf('validateNativeWorkbookRuntime_(');
  const terminalValidation = publishBody.indexOf(
    'validateNativeWorkbookRuntime_(',
    publishValidation + 1,
  );
  const publishStartHere = publishBody.indexOf('writeStartHereRuntime_(');
  const publishInfo = publishBody.indexOf('writeExportInfoRuntime_(');
  const publishFlush = publishBody.indexOf('SpreadsheetApp.flush()');
  assert.ok(
    publishValidation < publishStartHere && publishStartHere < publishInfo &&
      publishInfo < publishFlush && publishFlush < terminalValidation,
    'checkpoint-free FINALIZING validation must pass before terminal metadata is published',
  );
  assert.notEqual(terminalValidation, -1, 'published terminal metadata must be validated');
  assert.match(publishBody, /allowFinalizing:\s*true/);
  assert.doesNotMatch(publishBody, /allowCheckpoint:\s*true/);
  assert.match(publishBody, /preserveOwnerMarker:\s*true/);

  const preservingWriterStart = source.indexOf('function writeMatrixPreservingOwnerRuntime_(');
  const preservingWriterEnd = source.indexOf('\nfunction ', preservingWriterStart + 1);
  assert.notEqual(preservingWriterStart, -1, 'missing owner-preserving metadata writer');
  assert.doesNotMatch(
    source.slice(preservingWriterStart, preservingWriterEnd),
    /clearContents\(/,
    'terminal metadata writer must never erase the only surviving owner marker',
  );
  const exportInfoWriterStart = source.indexOf('function writeExportInfoRuntime_(');
  const exportInfoWriterEnd = source.indexOf('\nfunction ', exportInfoWriterStart + 1);
  assert.match(
    source.slice(exportInfoWriterStart, exportInfoWriterEnd),
    /getDisplayValue\(\)[\s\S]*OWNER_KEY/,
    'every rewrite of an already-owned _export_info sheet must preserve its owner marker',
  );

  const engineStart = source.indexOf('function runManifestEngine_(');
  const engineEnd = source.indexOf('// --------------------------------------------------------------------------', engineStart + 1);
  const engine = source.slice(engineStart, engineEnd);
  assert.ok(
    engine.indexOf('adapter.finalizeWorkbook(') < engine.indexOf('adapter.clearState()'),
    'checkpoint cleanup must follow final workbook validation',
  );
  assert.ok(
    engine.indexOf('adapter.clearState()') < engine.indexOf('adapter.publishWorkbook('),
    'READY publication must follow checkpoint cleanup',
  );
});

test('native validation accepts typed row counts even when display values contain commas', () => {
  const { api } = loadExporter();
  const { workbook, manifest } = nativeWorkbookFixture(api);

  assert.doesNotThrow(() => api.validateNativeWorkbookRuntime(
    workbook,
    SPREADSHEET_ID,
    manifest,
    ACCOUNT_ID,
    { allowCheckpoint: true },
  ));
});

test('native validation rejects a weekly range that differs from the exact 90-day range', () => {
  const { api } = loadExporter();
  const { workbook, manifest } = nativeWorkbookFixture(api);
  const rangeRow = workbook.getSheetByName('_export_info').rows
    .find((row) => row[0] === 'aggregate_range');
  rangeRow[3] = '2026-05-25 through 2026-08-23';

  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      workbook,
      SPREADSHEET_ID,
      manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /weekly range.*90-day|weekly.*aggregate/i,
  );
});

test('native validation enforces the disclosed 28-day Change Event exception', () => {
  const { api } = loadExporter();
  const { workbook, manifest } = nativeWorkbookFixture(api);
  const rangeRow = workbook.getSheetByName('_export_info').rows
    .find((row) => row[0] === 'aggregate_range');
  rangeRow[5] = '2026-07-30 through 2026-08-27';

  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      workbook,
      SPREADSHEET_ID,
      manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /change.*28 complete days|28-day.*change/i,
  );
});

test('native validation rejects declared row counts that differ from physical rows', () => {
  const { api } = loadExporter();
  const { workbook, manifest } = nativeWorkbookFixture(api);
  const info = workbook.getSheetByName('_export_info');
  info.rows.find((row) => row[0] === 'search_terms')[2] = 54320;

  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      workbook,
      SPREADSHEET_ID,
      manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /physical rows|row count/i,
  );
});

test('native validation requires declared row counts to be typed integer numbers', () => {
  const { api } = loadExporter();

  for (const invalidCount of ['', null, false, '0']) {
    const { workbook, manifest } = nativeWorkbookFixture(api);
    workbook.getSheetByName('_export_info').rows
      .find((row) => row[0] === 'campaign')[2] = invalidCount;
    assert.throws(
      () => api.validateNativeWorkbookRuntime(
        workbook,
        SPREADSHEET_ID,
        manifest,
        ACCOUNT_ID,
        { allowCheckpoint: true },
      ),
      /typed|declared row count|integer/i,
      String(invalidCount),
    );
  }
});

test('native validation rejects output headers that differ from the manifest', () => {
  const { api } = loadExporter();
  const { workbook, manifest } = nativeWorkbookFixture(api);
  workbook.getSheetByName('campaign').rows[0][0] = 'unexpected.header';

  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      workbook,
      SPREADSHEET_ID,
      manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /campaign.*headers|headers.*campaign/i,
  );
});

test('native validation requires exactly 41 visible output tabs in manifest order', () => {
  const { api } = loadExporter();
  const ordered = nativeWorkbookFixture(api);
  const checkpoint = ordered.workbook.sheets.pop();
  [ordered.workbook.sheets[4], ordered.workbook.sheets[5]] =
    [ordered.workbook.sheets[5], ordered.workbook.sheets[4]];
  ordered.workbook.sheets.push(checkpoint);
  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      ordered.workbook,
      SPREADSHEET_ID,
      ordered.manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /tab set|order|41/i,
  );

  const hidden = nativeWorkbookFixture(api);
  hidden.workbook.getSheetByName('campaign').hidden = true;
  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      hidden.workbook,
      SPREADSHEET_ID,
      hidden.manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /hidden|visible/i,
  );
});

test('native validation tolerates only the expected checkpoint during finalization', () => {
  const { api } = loadExporter();
  const missingCheckpoint = nativeWorkbookFixture(api);
  missingCheckpoint.workbook.sheets.pop();
  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      missingCheckpoint.workbook,
      SPREADSHEET_ID,
      missingCheckpoint.manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /checkpoint/i,
  );

  const fixture = nativeWorkbookFixture(api);
  fixture.workbook.sheets.push(new FixtureSheet(
    '__gads_export_stage__campaign',
    [['unexpected']],
    { hidden: true },
  ));

  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      fixture.workbook,
      SPREADSHEET_ID,
      fixture.manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /temporary|stage|tab set|undeclared/i,
  );
});

test('native prepublication validation accepts FINALIZING but never treats it as READY', () => {
  const { api } = loadExporter();
  const fixture = nativeWorkbookFixture(api);
  const statusRow = fixture.workbook.getSheetByName('_export_info').rows[1];
  statusRow[3] = 'FINALIZING';
  statusRow[7] = 'IN_PROGRESS';

  assert.doesNotThrow(() => api.validateNativeWorkbookRuntime(
    fixture.workbook,
    SPREADSHEET_ID,
    fixture.manifest,
    ACCOUNT_ID,
    { allowCheckpoint: true, allowFinalizing: true },
  ));
  assert.throws(
    () => api.validateNativeWorkbookRuntime(
      fixture.workbook,
      SPREADSHEET_ID,
      fixture.manifest,
      ACCOUNT_ID,
      { allowCheckpoint: true },
    ),
    /not terminal|FINALIZING/i,
  );
});

function restartWorkbook(api, status = 'COMPLETE', accountId = ACCOUNT_ID) {
  return new FixtureWorkbook([
    new FixtureSheet('_export_info', [
      [api.OWNER_KEY, api.VERSION, 'output_schema_version', api.OUTPUT_SCHEMA_VERSION],
      ['run_id', 'run-1', 'overall_status', status],
      ['account_id', accountId, 'account_name', 'Example Account'],
      ['tab', 'status', 'rows', 'duration_seconds'],
    ]),
  ]);
}

test('owned workbook restart accepts only the same advertiser account', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  assert.equal(
    api.assertOwnedWorkbookRestartIdentityRuntime(
      restartWorkbook(api),
      manifest,
      ACCOUNT_ID,
    ),
    true,
  );
  assert.throws(
    () => api.assertOwnedWorkbookRestartIdentityRuntime(
      restartWorkbook(api),
      manifest,
      '999-999-9999',
    ),
    /account|customer|advertiser/i,
  );
});

test('owned workbook restart refuses nonterminal metadata when its checkpoint is missing', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();

  for (const status of ['RUNNING', 'PAUSED', 'FINALIZING', 'COMPLETE_WITH_ERRORS', '']) {
    assert.throws(
      () => api.assertOwnedWorkbookRestartIdentityRuntime(
        restartWorkbook(api, status),
        manifest,
        ACCOUNT_ID,
      ),
      /checkpoint|nonterminal|reset/i,
      status || 'missing status',
    );
  }
});

test('owned workbook restart rejects undeclared content but permits a truly blank extra tab', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const blank = restartWorkbook(api);
  blank.sheets.push(new FixtureSheet('Sheet1'));
  assert.equal(
    api.assertOwnedWorkbookRestartIdentityRuntime(blank, manifest, ACCOUNT_ID),
    true,
  );

  const notes = restartWorkbook(api);
  notes.sheets.push(new FixtureSheet('Notes', [['do not overwrite']]));
  assert.throws(
    () => api.assertOwnedWorkbookRestartIdentityRuntime(notes, manifest, ACCOUNT_ID),
    /undeclared|Notes|tab/i,
  );

  const drawing = restartWorkbook(api);
  drawing.sheets.push(new FixtureSheet('DrawingOnly', [], { drawings: [{}] }));
  assert.throws(
    () => api.assertOwnedWorkbookRestartIdentityRuntime(drawing, manifest, ACCOUNT_ID),
    /undeclared|DrawingOnly|tab/i,
  );
});

test('owned workbook restart rejects a hidden declared sheet', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const workbook = restartWorkbook(api);
  workbook.sheets.push(new FixtureSheet(
    'START_HERE',
    [['Google Ads Analysis Workbook']],
    { hidden: true },
  ));

  assert.throws(
    () => api.assertOwnedWorkbookRestartIdentityRuntime(workbook, manifest, ACCOUNT_ID),
    /hidden/i,
  );
});

test('RESET metadata may restart before an advertiser identity has been written', () => {
  const { api } = loadExporter();
  assert.equal(
    api.assertOwnedWorkbookRestartIdentityRuntime(
      restartWorkbook(api, 'RESET', ''),
      api.getManifestDefinition(),
      '999-999-9999',
    ),
    true,
  );
});
