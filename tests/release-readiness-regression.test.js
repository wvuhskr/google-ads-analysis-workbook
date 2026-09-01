'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const {
  loadExporter,
  createPersistentRichTextHarness,
} = require('./load-exporter');

function tabDictionaryRowsForTab(api, tab) {
  const rows = api.buildDataDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  return rows.slice(1)
    .filter((row) => row[headers.indexOf('tab')] === tab)
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index]])));
}

function fieldDictionaryRowsForTab(api, tab) {
  const rows = api.buildFieldDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  return rows.slice(1)
    .filter((row) => row[headers.indexOf('tab')] === tab)
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index]])));
}

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

function createCapturingWriterSheet(width) {
  const writes = [];
  let lastRow = 0;
  const sheet = {
    getLastRow() { return lastRow; },
    getMaxRows() { return 100; },
    getMaxColumns() { return width; },
    getParent() { return spreadsheet; },
    insertRowsAfter() {},
    insertColumnsAfter() {},
    getRange(row, column, rowCount, columnCount) {
      assert.equal(column, 1);
      assert.equal(columnCount, width);
      return {
        setValues(values) {
          writes.push(values.map((candidate) => candidate.slice()));
          lastRow = Math.max(lastRow, row + rowCount - 1);
        },
      };
    },
  };
  const spreadsheet = { getSheets() { return [sheet]; } };
  return { sheet, writes };
}

function createCapturingMatrixSheet(width) {
  const writes = [];
  const sheet = {
    maxRows: 20,
    maxColumns: width,
    getMaxRows() { return this.maxRows; },
    getMaxColumns() { return this.maxColumns; },
    getParent() { return spreadsheet; },
    clearContents() {},
    insertRowsAfter(_after, count) { this.maxRows += count; },
    insertColumnsAfter(_after, count) { this.maxColumns += count; },
    setFrozenRows() {},
    getRange(_row, _column, rowCount, columnCount) {
      assert.equal(columnCount, width);
      return {
        setValues(values) {
          assert.equal(values.length, rowCount);
          writes.push(JSON.parse(JSON.stringify(values)));
        },
      };
    },
  };
  const spreadsheet = { getSheets() { return [sheet]; } };
  return { sheet, writes };
}

function createEngineState(api) {
  return api.createRunState(
    {
      version: api.VERSION,
      outputSchemaVersion: api.OUTPUT_SCHEMA_VERSION,
      accountId: '000-000-0000',
      spreadsheetId: 'sheet-123',
      configSignature: 'config-abc',
    },
    0,
    { aggregate: {}, weekly: {}, change: {} },
    ['job-one'],
  );
}

function validSafetyConfig() {
  return {
    INCLUDE_SENSITIVE_CHANGE_DETAILS: false,
    DIAGNOSTICS_LOG_SAMPLE_ROWS: false,
    HIDE_RAW_MICROS_COLUMNS: true,
    ALLOW_RESET: false,
    FREEZE_CONTEXT_COLUMNS: 0,
    MIN_REMAINING_SECONDS: 180,
    MIN_COMMIT_REMAINING_SECONDS: 360,
  };
}

test('production source contains no rejected conversion or distribution subsystem', () => {
  const { source } = loadExporter();

  assert.doesNotMatch(source, /\bScriptApp\b|\bUrlFetchApp\b|\bDriveApp\b/);
  assert.doesNotMatch(source, /drive\/v3\/files|AUTOMATIC_XLSX|RETRY_DISTRIBUTION/);
  assert.doesNotMatch(source, /createDistribution|distribution_/);
  assert.doesNotMatch(
    source,
    /distribution(?:Status|FileUrl|FileName|Bytes|CreatedAtMs|Error)/,
  );
});

test('live export calls the manifest engine directly with no post-export artifact action', () => {
  const { source } = loadExporter();
  assert.doesNotMatch(source, /RUN_MODE|CLEAN_XLSX|exportSanitizedXlsx/);

  const exportStart = source.indexOf('function runExport_(');
  const exportEnd = source.indexOf('function runtimeWriterOptions_(', exportStart);
  const exportSource = source.slice(exportStart, exportEnd);
  const advertiserCheckIndex = exportSource.indexOf('assertAdvertiserAccountRuntime_(');
  const fetchIndex = exportSource.indexOf('fetchCampaignIdsRuntime_(');
  const engineIndex = exportSource.indexOf('runManifestEngine_(');

  assert.notEqual(exportStart, -1, 'missing runExport_');
  assert.notEqual(advertiserCheckIndex, -1, 'missing advertiser-account guard');
  assert.notEqual(engineIndex, -1, 'runExport_ must invoke the resumable manifest engine');
  assert.ok(advertiserCheckIndex < fetchIndex, 'advertiser validation must precede account queries');
  assert.ok(fetchIndex < engineIndex, 'a fresh run must freeze campaign scope before execution');
  assert.doesNotMatch(exportSource, /resolveLiveRunAction_|runAutomatic|createDistribution/);
  assert.match(exportSource, /Native Google Sheet ready/);
});

function createEngineAdapter(options = {}) {
  const events = [];
  let nowMs = 0;
  const chunkCount = Number(options.chunkCount || 1);
  const remaining = (options.remaining || []).slice();
  const adapter = {
    events,
    remainingSeconds() { return remaining.length ? remaining.shift() : 1_000; },
    saveState(state) {
      events.push(['save', state.status, state.jobIndex, state.chunkIndex, state.chunkInProgress]);
    },
    writeInfo(state) { events.push(['info', state.status, state.jobIndex, state.chunkIndex]); },
    startJob(job) { events.push(['start', job.id]); return `stage-${job.id}`; },
    getChunkCount() { return chunkCount; },
    getChunkStartRow(_job, state) { return 2 + Number(state.tabs['job-one'].rows || 0); },
    rollbackChunk(job, state) { events.push(['rollback', job.id, state.chunkStartRow]); },
    runChunk(job, _state, index) {
      events.push(['chunk', job.id, index]);
      nowMs += Number(options.chunkDurationMs || 0);
      return 1;
    },
    commitJob(job) { events.push(['commit', job.id]); },
    abortJob(job) { events.push(['abort', job.id]); },
    finalizeWorkbook() { events.push(['finalize']); },
    publishWorkbook(state) { events.push(['publish', state.status]); },
    hasPriorFinal() { return false; },
    clearState() { events.push(['clear']); },
    nowMs() { return nowMs; },
  };
  return adapter;
}

test('sheet-write encoding quote-prefixes every non-empty string without changing typed values', () => {
  // Break caught: returning raw text allows Google Sheets to create implicit
  // links; double-prefixing a literal apostrophe changes its display.
  const { api } = loadExporter();
  assert.equal(
    typeof api.encodeSheetCellForWrite,
    'function',
    'the plain-text write policy must be exposed as a pure verification seam',
  );

  const cases = [
    ['campaign.id', "'campaign.id"],
    ['www.example.test', "'www.example.test"],
    ['fictional-plumbing.example', "'fictional-plumbing.example"],
    ['https://example.test/path?a=1,b=2', "'https://example.test/path?a=1,b=2"],
    ['=IMPORTXML("https://example.test")', "'=IMPORTXML(\"https://example.test\")"],
    ["'literal-leading-apostrophe", "''literal-leading-apostrophe"],
    ['', ''],
    [null, ''],
    [42, 42],
    [false, false],
  ];
  for (const [input, expected] of cases) {
    assert.equal(api.encodeSheetCellForWrite(input), expected, String(input));
  }
});

test('row-buffer writes use quote-prefix encoding for body text and URL strings', () => {
  // Break caught: implementing an exported helper without routing row-buffer
  // writes through it leaves body cells vulnerable to implicit links.
  const { api } = loadExporter();
  const { sheet, writes } = createCapturingWriterSheet(5);
  const buffer = api.createRowBuffer(
    sheet,
    ['field', 'domain', 'keyword', 'url', 'count'],
    { batchRows: 10, cellLimit: 1_000, retries: 1, sleep() {} },
  );

  buffer.push([
    'campaign.id',
    'www.example.test',
    'fictional-plumbing.example',
    'https://example.test/landing-page',
    12,
  ]);
  buffer.flush();

  assert.deepEqual(JSON.parse(JSON.stringify(writes)), [[[
    "'campaign.id",
    "'www.example.test",
    "'fictional-plumbing.example",
    "'https://example.test/landing-page",
    12,
  ]]]);
});

test('matrix writer quote-prefixes headers, metadata, and URL strings', () => {
  // Break caught: body rows can remain protected while writeMatrixRuntime
  // silently regresses headers, _export_info, and checkpoint initialization.
  const { api } = loadExporter({ SpreadsheetApp: { flush() {} } });
  const { sheet, writes } = createCapturingMatrixSheet(4);

  api.writeMatrixRuntime(sheet, [
    ['campaign.id', 'domain', 'url', 'count'],
    ['000000001', 'www.example.test', 'https://example.test/path', 3],
  ]);

  assert.deepEqual(writes, [[
    ["'campaign.id", "'domain", "'url", "'count"],
    ["'000000001", "'www.example.test", "'https://example.test/path", 3],
  ]]);
});

test('quote-prefixed report headers remain plain text after formatting', () => {
  // Break caught: a later rich-text rewrite clears quotePrefix, recreating the
  // native formula and automatic-link behavior that the Sheet contract forbids.
  const harness = createPersistentRichTextHarness({ richTextClearsQuotePrefix: true });
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === 'campaign');
  const headers = Array.from(api.headersForJob(job));
  const sheet = harness.createSheet('campaign', [headers.map(() => '')]);
  sheet.getRange(1, 1, 1, headers.length).setValues([
    headers.map((value) => (value === '' ? '' : `'${value}`)),
  ]);

  api.formatReportSheet(sheet, job, 'USD');

  const column = headers.indexOf('campaign.id') + 1;
  const cell = sheet.ensureCell(1, column);
  assert.equal(cell.value, 'campaign.id');
  assert.equal(cell.quotePrefix, true);
  assert.deepEqual(linkedCharacters(cell.richText), []);
});

test('quote-prefixed field-dictionary cells remain plain text after formatting', () => {
  // Break caught: treating the entire dictionary as rich text strips the
  // native plain-text encoding from dotted field names in body rows.
  const harness = createPersistentRichTextHarness({ richTextClearsQuotePrefix: true });
  const { api } = loadExporter({ SpreadsheetApp: harness.SpreadsheetApp });
  const job = api.getManifestDefinition().find((candidate) => candidate.tab === '_field_dictionary');
  assert.ok(job, 'normalized field dictionary must be a manifest job');
  const headers = Array.from(api.headersForJob(job));
  const values = [headers, headers.map((header) => (
    header === 'field' ? 'campaign.id' : `metadata_${header}`
  ))];
  const sheet = harness.createSheet('_field_dictionary', values.map((row) => row.map(() => '')));
  sheet.getRange(1, 1, values.length, headers.length).setValues(values.map((row) => (
    row.map((value) => (value === '' ? '' : `'${value}`))
  )));

  api.formatReportSheet(sheet, job, 'USD');

  for (let row = 1; row <= values.length; row += 1) {
    for (let column = 1; column <= headers.length; column += 1) {
      const cell = sheet.ensureCell(row, column);
      assert.equal(cell.value, values[row - 1][column - 1]);
      assert.equal(cell.quotePrefix, true, `${row},${column} lost quotePrefix`);
      assert.deepEqual(linkedCharacters(cell.richText), []);
    }
  }
});

test('tab catalog points to the normalized field dictionary instead of packing schemas', () => {
  // Break caught: retaining a packed column_schema blob leaves metadata hard to
  // filter even if a second field catalog is added.
  const { api } = loadExporter();
  const rows = api.buildDataDictionaryRows(api.getManifestDefinition());
  const headers = Array.from(rows[0]);
  const catalogJob = api.getManifestDefinition().find((job) => job.tab === '_data_dictionary');

  assert.deepEqual(headers, Array.from(api.headersForJob(catalogJob)));
  assert.notEqual(headers.indexOf('field_dictionary_reference'), -1);
  assert.equal(headers.includes('column_schema'), false, 'packed column_schema must be removed');
  const tabIndex = headers.indexOf('tab');
  const referenceIndex = headers.indexOf('field_dictionary_reference');
  for (const row of rows.slice(1)) {
    const isMixedLayout = ['START_HERE', '_export_info'].includes(row[tabIndex]);
    assert.equal(
      row[referenceIndex],
      isMixedLayout ? '' : '_field_dictionary',
      `${row[tabIndex]} field-catalog reference`,
    );
  }
});

test('field dictionary is normalized to exactly one row per declared field', () => {
  // Break caught: packing every field into one column_schema cell makes the
  // dictionary unfilterable and permits field documentation to drift.
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  assert.equal(
    typeof api.buildFieldDictionaryRows,
    'function',
    'normalized field metadata must have a pure builder',
  );
  const rows = api.buildFieldDictionaryRows(manifest);
  const headers = Array.from(rows[0]);
  const fieldJob = manifest.find((job) => job.tab === '_field_dictionary');

  assert.deepEqual(headers, Array.from(api.headersForJob(fieldJob)));

  for (const header of [
    'tab', 'column_ordinal', 'field', 'source_fields', 'data_type',
    'unit', 'is_key', 'is_derived', 'derivation', 'blank_when',
  ]) {
    assert.notEqual(headers.indexOf(header), -1, `missing normalized dictionary column ${header}`);
  }

  const tabIndex = headers.indexOf('tab');
  const fieldIndex = headers.indexOf('field');
  const ordinalIndex = headers.indexOf('column_ordinal');
  const allPairs = new Set();
  for (const row of rows.slice(1)) {
    const tab = String(row[tabIndex] || '');
    const field = String(row[fieldIndex] || '');
    assert.notEqual(tab, '', 'dictionary tab must not be blank');
    assert.notEqual(field, '', `${tab} dictionary field must not be blank`);
    const pair = `${tab}\u0000${field}`;
    assert.equal(allPairs.has(pair), false, `duplicate dictionary field ${tab}.${field}`);
    allPairs.add(pair);
  }

  for (const job of manifest) {
    const entries = rows.slice(1)
      .filter((row) => row[tabIndex] === job.tab)
      .sort((left, right) => Number(left[ordinalIndex]) - Number(right[ordinalIndex]));
    const actual = Array.from(entries, (row) => row[fieldIndex]);
    const expected = Array.from(api.headersForJob(job));
    assert.deepEqual(actual, expected, `${job.tab} dictionary fields must match its output headers`);
    assert.deepEqual(
      Array.from(entries, (row) => row[ordinalIndex]),
      expected.map((_, index) => index + 1),
      `${job.tab} ordinals must reconstruct exact output order`,
    );
  }
  assert.equal(
    rows.slice(1).some((row) => row[tabIndex] === '_export_info'),
    false,
    '_export_info is a mixed-layout metadata sheet, not a rectangular field table',
  );
});

test('field dictionary preserves type, unit, key, derived, and blankability semantics', () => {
  // Break caught: normalizing by merely splitting text loses the structured
  // semantics marketers need to interpret and join fields safely.
  const { api } = loadExporter();
  const campaign = fieldDictionaryRowsForTab(api, 'campaign');
  const byField = Object.fromEntries(campaign.map((row) => [row.field, row]));

  assert.equal(byField['campaign.id'].data_type, 'identifier');
  assert.match(String(byField['campaign.id'].unit), /text-preserved ID/i);
  assert.equal(byField['campaign.id'].is_key, true);
  assert.match(String(byField['campaign.id'].blank_when), /(?:not applicable|unavailable)/i);
  assert.equal(byField['campaign.id'].is_derived, false);
  assert.equal(byField['metrics.cost_micros'].data_type, 'number');
  assert.match(String(byField['metrics.cost_micros'].unit), /micro-units/i);
  assert.equal(byField['metrics.cost_micros'].is_key, false);
  assert.match(String(byField.cost.unit), /account currency/i);
  assert.equal(byField.cost.is_derived, true);
  assert.match(String(byField.cost.source_fields), /metrics\.cost_micros/i);
  assert.match(String(byField.cost.derivation), /1,?000,?000|micro/i);
});

test('field dictionary documents custom joins, normalized outputs, unions, and concrete units', () => {
  const { api } = loadExporter();
  const geo = Object.fromEntries(
    fieldDictionaryRowsForTab(api, 'campaign_geo').map((row) => [row.field, row]),
  );
  const landing = Object.fromEntries(
    fieldDictionaryRowsForTab(api, 'ad_to_lp_map').map((row) => [row.field, row]),
  );
  const negatives = Object.fromEntries(
    fieldDictionaryRowsForTab(api, 'negative_keywords_all').map((row) => [row.field, row]),
  );
  const fieldDictionary = Object.fromEntries(
    fieldDictionaryRowsForTab(api, '_field_dictionary').map((row) => [row.field, row]),
  );
  const proximity = Object.fromEntries(
    fieldDictionaryRowsForTab(api, 'geo_proximity_targets').map((row) => [row.field, row]),
  );
  const audience = Object.fromEntries(
    fieldDictionaryRowsForTab(api, 'user_list_performance').map((row) => [row.field, row]),
  );
  const extensions = Object.fromEntries(
    fieldDictionaryRowsForTab(api, 'asset_extensions').map((row) => [row.field, row]),
  );

  assert.equal(geo.geo_country_name.is_derived, true);
  assert.match(String(geo.geo_country_name.source_fields), /country_criterion_id.*geo_target_constant\.name/i);
  assert.equal(landing.domain.is_derived, true);
  assert.match(String(landing.domain.source_fields), /final_url_norm/i);
  assert.equal(negatives.source.is_derived, true);
  assert.match(String(negatives.source.source_fields), /neg_keywords_campaign/i);
  assert.equal(fieldDictionary.column_ordinal.is_derived, true);
  assert.match(String(fieldDictionary.column_ordinal.unit), /1-based/i);
  assert.match(String(proximity['campaign_criterion.proximity.radius'].unit), /radius_units/i);
  assert.match(
    String(proximity['campaign_criterion.proximity.geo_point.latitude_in_micro_degrees'].unit),
    /microdegrees/i,
  );
  for (const field of ['scope', 'user_list_resource', 'criterion_id']) {
    assert.equal(audience[field].is_derived, true, `user_list_performance.${field}`);
    assert.match(String(audience[field].derivation), /scope|CAMPAIGN|AD_GROUP/i);
    assert.doesNotMatch(String(audience[field].unit), /^Google-reported/i);
  }
  for (const field of ['scope', 'field_type', 'association_status', 'source', 'asset.text']) {
    assert.equal(extensions[field].is_derived, true, `asset_extensions.${field}`);
    assert.notEqual(String(extensions[field].source_fields || ''), field);
    assert.doesNotMatch(String(extensions[field].unit), /^Google-reported/i);
  }
  assert.match(String(extensions['asset.text'].source_fields), /callout_asset\.callout_text/i);
  assert.match(String(extensions['asset.text'].source_fields), /sitelink_asset\.link_text/i);
  assert.match(String(audience['ad_group.id'].blank_when), /CAMPAIGN scope/i);
});

test('campaign geography warns against reconciling or summing inferred rows as totals', () => {
  // Break caught: ID-grain guidance alone still invites marketers to sum an
  // inferred geography view and compare it to authoritative campaign totals.
  const { api } = loadExporter();
  const warnings = tabDictionaryRowsForTab(api, 'campaign_geo')
    .map((row) => String(row.google_side_limitations || ''));
  const warning = warnings.find(Boolean) || '';

  assert.match(warning, /may not reconcile/i);
  assert.match(warning, /do not sum/i);
  assert.match(warning, /use\s+`?campaign`?\s+for\s+authoritative\s+totals/i);
});

test('export info documents conservative grid-safety overhead', () => {
  const { api } = loadExporter();
  const info = tabDictionaryRowsForTab(api, '_export_info')[0];
  assert.match(String(info.google_side_limitations), /conservative safety/i);
  assert.match(String(info.google_side_limitations), /checkpoint|blank-sheet/i);
});

test('per-job campaign chunk planning isolates expensive geography work', () => {
  // Break caught: applying the default 25-campaign batch to expensive geography
  // work creates a larger uncheckpointed unit than the runtime safety contract.
  const { api } = loadExporter();
  assert.equal(
    typeof api.chunkCampaignIdsForJob,
    'function',
    'runtime chunk planning must have one pure source of truth',
  );
  const manifest = api.getManifestDefinition();
  const campaignGeo = manifest.find((job) => job.tab === 'campaign_geo');
  const campaign = manifest.find((job) => job.tab === 'campaign');
  const unchunked = manifest.find((job) => job.tab === 'conversion_actions');
  const ids = Array.from({ length: 61 }, (_, index) => String(index + 1));

  const geoChunks = api.chunkCampaignIdsForJob(campaignGeo, ids, 25);
  const ordinaryChunks = api.chunkCampaignIdsForJob(campaign, ids, 25);
  const unchunkedPlan = api.chunkCampaignIdsForJob(unchunked, ids, 25);

  assert.equal(campaignGeo.campaignChunkSize, 10);
  assert.deepEqual(Array.from(geoChunks, (chunk) => chunk.length), [10, 10, 10, 10, 10, 10, 1]);
  assert.deepEqual(Array.from(ordinaryChunks, (chunk) => chunk.length), [25, 25, 11]);
  assert.deepEqual(Array.from(unchunkedPlan), [null]);
  assert.deepEqual(
    Array.from(geoChunks.flat()),
    Array.from(api.chunkCampaignIds(ids, 61)[0]),
    'chunk planning must preserve the exact deduplicated campaign universe',
  );
});

test('resume identity fingerprints per-job chunk plans and exact output contracts', () => {
  const { api } = loadExporter();
  const manifest = api.getManifestDefinition();
  const baseline = api.materialConfigSignature(manifest);
  const changedChunkPlan = manifest.map((job) => (
    job.tab === 'campaign_geo' ? { ...job, campaignChunkSize: 25 } : job
  ));
  const changedHeader = manifest.map((job) => (
    job.tab === 'campaign'
      ? { ...job, columns: Array.from(job.columns).concat([{ field: 'new.field', type: 'text', header: 'new.field' }]) }
      : job
  ));

  assert.notEqual(api.materialConfigSignature(changedChunkPlan), baseline);
  assert.notEqual(api.materialConfigSignature(changedHeader), baseline);

  const identity = {
    version: api.VERSION,
    outputSchemaVersion: api.OUTPUT_SCHEMA_VERSION,
    accountId: '000-000-0000',
    spreadsheetId: 'sheet-123',
    configSignature: baseline,
  };
  const state = api.createRunState(identity, 1_000, {}, manifest.map((job) => job.id));
  assert.throws(
    () => api.assertStateCompatible(
      state,
      { ...identity, configSignature: api.materialConfigSignature(changedChunkPlan) },
      2_000,
      24,
    ),
    /configuration/i,
  );
});

test('runtime safety configuration rejects quoted booleans and unsafe reserves', () => {
  const { api } = loadExporter();
  const baseline = validSafetyConfig();

  assert.equal(api.validateRuntimeSafetyConfig(baseline), true);
  for (const key of [
    'INCLUDE_SENSITIVE_CHANGE_DETAILS',
    'DIAGNOSTICS_LOG_SAMPLE_ROWS',
    'HIDE_RAW_MICROS_COLUMNS',
    'ALLOW_RESET',
  ]) {
    assert.throws(
      () => api.validateRuntimeSafetyConfig({ ...baseline, [key]: 'false' }),
      /without quotes/i,
      key,
    );
  }
  assert.throws(
    () => api.validateRuntimeSafetyConfig({ ...baseline, MIN_COMMIT_REMAINING_SECONDS: 60 }),
    /at least MIN_REMAINING_SECONDS/i,
  );
  assert.throws(
    () => api.validateRuntimeSafetyConfig({ ...baseline, MIN_COMMIT_REMAINING_SECONDS: 999_999 }),
    /execution limit/i,
  );
  assert.throws(
    () => api.validateRuntimeSafetyConfig({ ...baseline, FREEZE_CONTEXT_COLUMNS: -1 }),
    /nonnegative integer/i,
  );
  assert.throws(
    () => api.validateRuntimeSafetyConfig({ ...baseline, FREEZE_CONTEXT_COLUMNS: 2.5 }),
    /nonnegative integer/i,
  );
  assert.throws(
    () => api.validateRuntimeSafetyConfig({ ...baseline, MIN_REMAINING_SECONDS: true }),
    /must be numbers/i,
  );
  assert.throws(
    () => api.validateRuntimeSafetyConfig({
      ...baseline,
      MIN_REMAINING_SECONDS: true,
      MIN_COMMIT_REMAINING_SECONDS: true,
    }),
    /must be numbers/i,
  );
  assert.throws(
    () => api.validateRuntimeSafetyConfig({ ...baseline, FREEZE_CONTEXT_COLUMNS: true }),
    /must be a number/i,
  );
  assert.throws(
    () => api.validateRuntimeSafetyConfig({ ...baseline, FREEZE_CONTEXT_COLUMNS: '6' }),
    /must be a number/i,
  );
  assert.throws(
    () => api.runManifestEngine(
      createEngineState(api),
      [{ id: 'job-one', tab: 'job-one', required: true }],
      createEngineAdapter(),
      180,
      60,
      120,
    ),
    /ordinary reserve/i,
  );
});

test('Preview and Export reject boolean time reserves before querying the account', () => {
  for (const isPreview of [true, false]) {
    let accountQueryCount = 0;
    const { context } = loadExporter({
      AdsApp: {
        getExecutionInfo() { return { isPreview() { return isPreview; } }; },
        report() {
          accountQueryCount += 1;
          throw new Error('SENTINEL_ACCOUNT_QUERY_REACHED');
        },
      },
    });
    context.CONFIG.MIN_REMAINING_SECONDS = true;
    context.CONFIG.MIN_COMMIT_REMAINING_SECONDS = true;

    assert.throws(() => context.main(), /must be numbers/i, isPreview ? 'Preview' : 'Export');
    assert.equal(accountQueryCount, 0, isPreview ? 'Preview' : 'Export');
  }
});

test('runtime saves every chunk while throttling status-sheet refreshes to 120 seconds', () => {
  // Break caught: coupling user-facing status rewrites to transactional saves
  // adds a full formatted-sheet rewrite after every chunk and job.
  const { api } = loadExporter();
  const state = createEngineState(api);
  const adapter = createEngineAdapter({ chunkCount: 5, chunkDurationMs: 30_000 });
  const manifest = [{ id: 'job-one', tab: 'job-one', required: true, chunked: true }];

  const result = api.runManifestEngine(state, manifest, adapter, 180, 360, 120);

  assert.equal(result.status, 'COMPLETE');
  const completedChunkSaves = adapter.events
    .filter(([name, status, , , inProgress]) => name === 'save' && status === 'RUNNING' && !inProgress)
    .map(([, , , chunkIndex]) => chunkIndex);
  for (const completedChunk of [1, 2, 3, 4, 5]) {
    assert.equal(
      completedChunkSaves.includes(completedChunk),
      true,
      `chunk ${completedChunk} must have a durable checkpoint`,
    );
  }
  assert.deepEqual(
    adapter.events.filter(([name]) => name === 'info').map((event) => event.slice(1)),
    [
      ['RUNNING', 0, 0],
      ['RUNNING', 0, 4],
      ['FINALIZING', 1, 0],
    ],
    'status UI refreshes at invocation start, after 120 seconds, and before terminal publication',
  );
  assert.deepEqual(adapter.events.at(-1), ['publish', 'COMPLETE']);
});

test('runtime forces a PAUSED status refresh even before the throttle interval elapses', () => {
  // Break caught: throttling without a forced pause write leaves marketers with
  // stale RUNNING metadata even though the durable checkpoint is resumable.
  const { api } = loadExporter();
  const state = createEngineState(api);
  const adapter = createEngineAdapter({
    chunkCount: 2,
    chunkDurationMs: 10_000,
    remaining: [1_000, 1_000, 100],
  });
  const manifest = [{ id: 'job-one', tab: 'job-one', required: true, chunked: true }];

  const result = api.runManifestEngine(state, manifest, adapter, 180, 360, 120);

  assert.equal(result.status, 'PAUSED');
  const startIndex = adapter.events.findIndex(([name]) => name === 'start');
  const initialInfoIndex = adapter.events.findIndex((event) => (
    event[0] === 'info' && event[1] === 'RUNNING' && event[3] === 0
  ));
  assert.ok(initialInfoIndex >= 0 && initialInfoIndex < startIndex, 'start refresh must precede job work');
  assert.deepEqual(adapter.events.filter(([name]) => name === 'info').at(-1).slice(1), [
    'PAUSED', 0, 1,
  ]);
});

test('runtime pauses below the separate commit reserve and resumes without rerunning the chunk', () => {
  // Break caught: using the ordinary 180-second reserve for a potentially large
  // format/commit step can time out after all report rows were already staged.
  const { api } = loadExporter();
  const state = createEngineState(api);
  const manifest = [{ id: 'job-one', tab: 'job-one', required: true, chunked: true }];
  const first = createEngineAdapter({
    chunkCount: 1,
    remaining: [1_000, 1_000, 300],
  });

  const paused = api.runManifestEngine(state, manifest, first, 180, 360, 120);

  assert.equal(paused.status, 'PAUSED');
  assert.equal(paused.jobIndex, 0);
  assert.equal(paused.chunkIndex, 1);
  assert.equal(first.events.some(([name]) => name === 'commit'), false);

  const resumed = createEngineAdapter({ chunkCount: 1, remaining: [1_000, 1_000] });
  const completed = api.runManifestEngine(paused, manifest, resumed, 180, 360, 120);

  assert.equal(completed.status, 'COMPLETE');
  assert.equal(resumed.events.some(([name]) => name === 'chunk'), false);
  assert.equal(resumed.events.filter(([name]) => name === 'commit').length, 1);
});
