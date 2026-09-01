'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadExporter } = require('./load-exporter');

const NOW_MS = Date.UTC(2026, 7, 25, 12, 0, 0);
const OUTPUT_SCHEMA_VERSION = 9;

function identity(outputSchemaVersion) {
  const value = {
    version: 'v1.0.0',
    accountId: '123-456-7890',
    spreadsheetId: 'sheet-123',
    configSignature: 'config-abc',
  };
  if (outputSchemaVersion !== undefined) value.outputSchemaVersion = outputSchemaVersion;
  return value;
}

function frozenRanges() {
  return {
    aggregate: { start: '2026-05-27', end: '2026-08-24' },
    weekly: { start: '2026-05-27', end: '2026-08-24' },
    change: { start: '2026-07-28', end: '2026-08-24' },
  };
}

function dictionaryContract(api, tab) {
  const rows = api.buildDataDictionaryRows(api.getManifestDefinition());
  const headers = rows[0];
  const row = rows.slice(1).find((candidate) => candidate[0] === tab);
  assert.ok(row, `missing dictionary row for ${tab}`);
  return Object.fromEntries(headers.map((header, index) => [header, row[index]]));
}

function fieldDictionaryContract(api, tab, field) {
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

function metadataValue(rows, key) {
  for (const row of rows) {
    for (let index = 0; index < row.length - 1; index += 1) {
      if (row[index] === key) return row[index + 1];
    }
  }
  return undefined;
}

function successfulAdapter() {
  return {
    remainingSeconds() { return 1_000; },
    saveState() {},
    writeInfo() {},
    startJob(job) { return `stage-${job.id}`; },
    getChunkCount() { return 1; },
    getChunkStartRow() { return 2; },
    rollbackChunk() {},
    runChunk() { return 3; },
    commitJob() {},
    abortJob() {},
    finalizeWorkbook() {},
    publishWorkbook() {},
    hasPriorFinal() { return false; },
    clearState() {},
    nowMs() { return NOW_MS + 300_000; },
  };
}

test('landing_pages tells marketers not to use URL rows as authoritative totals', () => {
  const { api } = loadExporter();
  const warning = dictionaryContract(api, 'landing_pages').google_side_limitations;
  const requiredGuidance = {
    'non-additive': /non-additive/i,
    'may not reconcile': /may not reconcile/i,
    'do not sum': /do not sum/i,
    'use campaign for authoritative totals': /use\s+`?campaign`?\s+for\s+authoritative\s+totals/i,
  };
  const missing = Object.entries(requiredGuidance)
    .filter(([, pattern]) => !pattern.test(String(warning)))
    .map(([label]) => label);

  assert.deepEqual(missing, [], `landing_pages warning is missing: ${missing.join(', ')}`);
});

test('a successful landing_pages job remains OK despite its interpretation warning', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'landing_pages');
  const state = api.createRunState(identity(), NOW_MS, frozenRanges(), [job.id]);

  const result = api.runManifestEngine(state, [job], successfulAdapter(), 180);
  const infoRows = api.buildExportInfoRows(result, [job]);
  const infoRow = infoRows.find((row) => row[0] === 'landing_pages');

  assert.equal(result.status, 'COMPLETE');
  assert.equal(result.tabs.landing_pages.status, 'OK');
  assert.equal(result.tabs.landing_pages.limitation, '');
  const statusHeader = infoRows.find((row) => row[0] === 'tab');
  assert.equal(infoRow[statusHeader.indexOf('limitation_or_error')], '');
});

test('new run state records the output schema version used to create its tabs', () => {
  const { api } = loadExporter();
  const state = api.createRunState(identity(), NOW_MS, frozenRanges(), ['campaign_geo']);

  assert.equal(api.OUTPUT_SCHEMA_VERSION, OUTPUT_SCHEMA_VERSION);
  assert.equal(state.outputSchemaVersion, OUTPUT_SCHEMA_VERSION);
  assert.equal(typeof api.RUNTIME_CONTRACT_VERSION, 'number');
  assert.equal(api.RUNTIME_CONTRACT_VERSION, 10);
  assert.equal(state.runtimeContractVersion, api.RUNTIME_CONTRACT_VERSION);
});

test('_export_info exposes the output schema version for downstream provenance', () => {
  const { api } = loadExporter();
  const state = api.createRunState(identity(), NOW_MS, frozenRanges(), ['campaign_geo']);
  state.outputSchemaVersion = OUTPUT_SCHEMA_VERSION;
  state.status = 'COMPLETE';

  const rows = api.buildExportInfoRows(state, []);

  assert.deepEqual(Array.from(rows[0].slice(0, 4)), [
    api.OWNER_KEY,
    api.VERSION,
    'output_schema_version',
    OUTPUT_SCHEMA_VERSION,
  ]);
  assert.equal(metadataValue(rows, 'output_schema_version'), OUTPUT_SCHEMA_VERSION);
  assert.equal(metadataValue(rows, 'runtime_contract_version'), api.RUNTIME_CONTRACT_VERSION);
});

test('_export_info reserves G2:H7 for native-workbook metadata', () => {
  const { api } = loadExporter();
  const current = api.createRunState(identity(), NOW_MS, frozenRanges(), []);
  current.status = 'COMPLETE';
  const rows = api.buildExportInfoRows(current, []);
  const metadata = Array.from(rows.slice(1, 7), (row) => Array.from(row.slice(6, 8)));

  assert.deepEqual(metadata.slice(0, 5), [
    ['workbook_status', 'READY'],
    ['deliverable_type', 'NATIVE_GOOGLE_SHEET'],
    ['last_complete_day', '2026-08-24'],
    ['workbook_url', ['https://docs.google.com', 'spreadsheets', 'd', 'sheet-123', 'edit'].join('/')],
    ['reporting_window', 'LAST_90_COMPLETE_DAYS'],
  ]);
  assert.equal(metadata[5][0], 'refresh_behavior');
  assert.match(String(metadata[5][1]), /run main\(\).*fresh export/i);
  assert.equal(rows.flat().some((value) => /distribution_/i.test(String(value))), false);
  assert.equal(rows[7][0], 'tab');
});

test('resume rejects a checkpoint created under a different runtime contract', () => {
  const { api } = loadExporter();
  const currentIdentity = identity();
  const state = api.createRunState(currentIdentity, NOW_MS, frozenRanges(), ['campaign_geo']);

  state.runtimeContractVersion = api.RUNTIME_CONTRACT_VERSION + 1;
  assert.throws(
    () => api.assertStateCompatible(state, currentIdentity, NOW_MS + 1_000, 24),
    /runtime contract version/i,
  );
  assert.equal(
    JSON.parse(api.materialConfigSignature(api.getManifestDefinition())).runtimeContractVersion,
    api.RUNTIME_CONTRACT_VERSION,
  );
});

test('RESET metadata never claims preserved tabs already use the current output schema', () => {
  const { api } = loadExporter();
  const state = api.createRunState(identity(), NOW_MS, frozenRanges(), ['campaign_geo']);
  state.status = 'RESET';

  const rows = api.buildExportInfoRows(state, []);
  const nextAction = rows.find((row) => row[0] === 'next_action');

  assert.equal(metadataValue(rows, 'overall_status'), 'RESET');
  assert.equal(metadataValue(rows, 'output_schema_version'), '');
  assert.match(nextAction[1], /preserved/i);
  assert.match(nextAction[1], /older output schema|prior output schema/i);
  assert.match(nextAction[1], /run main\(\)/i);
  assert.doesNotMatch(nextAction[1], /finished successfully|coverage is complete/i);
});

test('mixed prior-tab preservation never claims one realized workbook output schema', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign_geo');
  const state = api.createRunState(identity(), NOW_MS, frozenRanges(), [job.id]);
  state.status = 'COMPLETE_WITH_ERRORS';
  state.tabs.campaign_geo = {
    status: 'ERROR_PREVIOUS_PRESERVED',
    rows: 0,
    durationMs: 500,
    error: 'injected formatting failure',
    limitation: '',
    priorPreserved: true,
  };

  const rows = api.buildExportInfoRows(state, [job]);
  const nextAction = rows.find((row) => row[0] === 'next_action');
  const jobRow = rows.find((row) => row[0] === 'campaign_geo');

  assert.equal(metadataValue(rows, 'output_schema_version'), '');
  const statusHeader = rows.find((row) => row[0] === 'tab');
  assert.equal(jobRow[statusHeader.indexOf('prior_data_preserved')], 'YES');
  assert.match(nextAction[1], /prior|preserved/i);
  assert.match(nextAction[1], /schema/i);
});

test('limited optional prior-tab preservation explicitly warns about mixed schemas', () => {
  const { api } = loadExporter();
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign_geo');
  const state = api.createRunState(identity(), NOW_MS, frozenRanges(), [job.id]);
  state.status = 'COMPLETE_WITH_LIMITATIONS';
  state.tabs.campaign_geo = {
    status: 'LIMITED',
    rows: 0,
    durationMs: 500,
    error: 'injected optional formatting failure',
    limitation: '',
    priorPreserved: true,
  };

  const rows = api.buildExportInfoRows(state, [job]);
  const nextAction = rows.find((row) => row[0] === 'next_action');

  assert.equal(metadataValue(rows, 'output_schema_version'), '');
  assert.match(nextAction[1], /prior|preserved/i);
  assert.match(nextAction[1], /schema/i);
});

test('output schema is withheld while a workbook is still RUNNING or PAUSED', () => {
  const { api } = loadExporter();
  for (const status of ['RUNNING', 'PAUSED']) {
    const state = api.createRunState(identity(), NOW_MS, frozenRanges(), ['campaign_geo']);
    state.status = status;
    const rows = api.buildExportInfoRows(state, []);
    assert.equal(
      metadataValue(rows, 'output_schema_version'),
      '',
      `${status} workbook may still contain prior-schema final tabs`,
    );
  }
});

test('resume rejects checkpoint state that predates output-schema provenance', () => {
  const { api } = loadExporter();
  const currentIdentity = identity();
  const priorState = api.createRunState(currentIdentity, NOW_MS, frozenRanges(), ['campaign_geo']);
  delete priorState.outputSchemaVersion;

  assert.throws(
    () => api.assertStateCompatible(priorState, currentIdentity, NOW_MS + 1_000, 24),
    /output schema|schema version|schema revision/i,
  );
});

test('resume rejects checkpoint state from a different output schema version', () => {
  const { api } = loadExporter();
  const currentIdentity = identity();
  const priorState = api.createRunState(currentIdentity, NOW_MS, frozenRanges(), ['campaign_geo']);
  priorState.outputSchemaVersion = OUTPUT_SCHEMA_VERSION - 1;

  assert.throws(
    () => api.assertStateCompatible(priorState, currentIdentity, NOW_MS + 1_000, 24),
    /output schema|schema version|schema revision/i,
  );
});

test('resume accepts checkpoint state from the same output schema version', () => {
  const { api } = loadExporter();
  const currentIdentity = identity();
  const currentState = api.createRunState(currentIdentity, NOW_MS, frozenRanges(), ['campaign_geo']);
  currentState.outputSchemaVersion = OUTPUT_SCHEMA_VERSION;

  assert.doesNotThrow(
    () => api.assertStateCompatible(currentState, currentIdentity, NOW_MS + 1_000, 24),
  );
});

test('campaign_geo declares a canonical criterion-ID key instead of a display-name key', () => {
  const { api } = loadExporter();
  const contract = dictionaryContract(api, 'campaign_geo');
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign_geo');
  const outputHeaders = Array.from(api.headersForJob(job));

  assert.equal(
    contract.keys,
    'campaign.id, geographic_view.location_type, geographic_view.country_criterion_id, ' +
      'geo_target_most_specific_location_criterion_id, geo_target_state_criterion_id',
  );
  assert.equal(outputHeaders.includes('segments.geo_target_most_specific_location'), true);
  assert.equal(outputHeaders.includes('segments.geo_target_state'), true);
  assert.equal(outputHeaders.includes('geo_target_most_specific_location_criterion_id'), true);
  assert.equal(outputHeaders.includes('geo_target_state_criterion_id'), true);
  for (const field of [
    'geo_target_most_specific_location_criterion_id',
    'geo_target_state_criterion_id',
  ]) {
    const definition = fieldDictionaryContract(api, 'campaign_geo', field);
    assert.equal(definition.data_type, 'identifier');
    assert.match(String(definition.unit), /text-preserved ID/i);
    assert.equal(definition.is_key, true);
  }
});

test('campaign_geo dictionary explains its ID grain and warns that location names repeat', () => {
  const { api } = loadExporter();
  const contract = dictionaryContract(api, 'campaign_geo');
  const requiredGuidance = {
    'ID-based row grain': /campaign.*location type.*country.*most-specific.*criterion.*state.*criterion/i,
    'most-specific ID identity': /geo_target_most_specific_location_criterion_id.*(?:canonical|identity)|(?:canonical|identity).*geo_target_most_specific_location_criterion_id/i,
    'state ID identity': /geo_target_state_criterion_id.*(?:canonical|identity)|(?:canonical|identity).*geo_target_state_criterion_id/i,
    'repeated display labels': /names?.*(display )?labels?.*may repeat|display labels?.*may repeat/i,
    'not configured targeting': /not (?:the same as|identical to) configured targeting/i,
  };
  const values = {
    'ID-based row grain': contract.row_grain,
    'most-specific ID identity': contract.google_side_limitations,
    'state ID identity': contract.google_side_limitations,
    'repeated display labels': contract.google_side_limitations,
    'not configured targeting': contract.google_side_limitations,
  };
  const missing = Object.entries(requiredGuidance)
    .filter(([label, pattern]) => !pattern.test(String(values[label])))
    .map(([label]) => label);

  assert.deepEqual(missing, [], `campaign_geo dictionary is missing: ${missing.join(', ')}`);
});
