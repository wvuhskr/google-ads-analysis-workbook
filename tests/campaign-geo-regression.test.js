'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { loadExporter } = require('./load-exporter');

const AGGREGATE_RANGES = {
  aggregate: { start: '2026-05-28', end: '2026-08-25' },
};

const CAMPAIGN_ID = '100200300';

const EXPECTED_GEO_KEY_HEADERS = [
  'campaign.id',
  'geographic_view.location_type',
  'geographic_view.country_criterion_id',
  'geo_target_most_specific_location_criterion_id',
  'geo_target_state_criterion_id',
];

function requireFunction(api, name) {
  assert.equal(
    typeof api[name],
    'function',
    `exporter.${name} must expose the campaign geography behavior`,
  );
  return api[name];
}

function findJob(api, tab) {
  const job = Array.from(api.getManifestDefinition()).find((candidate) => candidate.tab === tab);
  assert.ok(job, `${tab} must remain in the export manifest`);
  return job;
}

function iteratorFor(rows) {
  let index = 0;
  return {
    hasNext() { return index < rows.length; },
    next() {
      assert.ok(index < rows.length, 'iterator read past the supplied report rows');
      const row = rows[index];
      index += 1;
      return row;
    },
  };
}

function reportFor(rows) {
  return { rows: () => iteratorFor(rows) };
}

function displayedSheetValue(value) {
  return typeof value === 'string' && value.startsWith("'") ? value.slice(1) : value;
}

function idsFromLookupQuery(query) {
  const match = /WHERE\s+geo_target_constant\.id\s+IN\s*\(([^)]*)\)/i.exec(query);
  assert.ok(match, `expected an ID-bounded geo_target_constant query, received:\n${query}`);
  if (!match[1].trim()) return [];
  return match[1].split(',').map((value) => value.trim().replace(/^['"]|['"]$/g, ''));
}

class MemorySheet {
  constructor(headers) {
    this.data = [headers.slice()];
    this.maxRows = 1;
    this.maxColumns = Math.max(1, headers.length);
    this.parent = { getSheets: () => [this] };
  }

  getParent() { return this.parent; }
  getMaxRows() { return this.maxRows; }
  getMaxColumns() { return this.maxColumns; }
  getLastRow() { return this.data.length; }
  insertRowsAfter(_after, count) { this.maxRows += count; }
  insertColumnsAfter(_after, count) { this.maxColumns += count; }

  getRange(row, column, rowCount, columnCount) {
    assert.ok(row >= 1 && column >= 1 && rowCount >= 1 && columnCount >= 1);
    assert.ok(row + rowCount - 1 <= this.maxRows, 'write exceeds memory sheet row bounds');
    assert.ok(column + columnCount - 1 <= this.maxColumns, 'write exceeds memory sheet column bounds');
    return {
      setValues: (values) => {
        assert.equal(values.length, rowCount);
        values.forEach((valuesRow, rowOffset) => {
          assert.equal(valuesRow.length, columnCount);
          const targetIndex = row - 1 + rowOffset;
          if (!this.data[targetIndex]) this.data[targetIndex] = [];
          valuesRow.forEach((value, columnOffset) => {
            this.data[targetIndex][column - 1 + columnOffset] = displayedSheetValue(value);
          });
        });
        return this;
      },
    };
  }
}

function writerOptions() {
  return {
    batchRows: 2,
    cellLimit: 100000,
    retries: 1,
    sleep() {},
  };
}

function campaignGeoRow(overrides) {
  return {
    'customer.id': '0000000000',
    'customer.descriptive_name': 'Synthetic Fixture Account',
    'customer.currency_code': 'USD',
    'customer.time_zone': 'America/New_York',
    'campaign.id': CAMPAIGN_ID,
    'campaign.name': 'Northstar Demand Gen',
    'campaign.status': 'ENABLED',
    'campaign.advertising_channel_type': 'DEMAND_GEN',
    'campaign.advertising_channel_sub_type': 'DEMAND_GEN',
    'geographic_view.location_type': 'LOCATION_OF_PRESENCE',
    'geographic_view.country_criterion_id': '2840',
    'segments.geo_target_most_specific_location': '1023191',
    'segments.geo_target_state': '21167',
    'metrics.impressions': '10',
    'metrics.clicks': '0',
    'metrics.ctr': '0',
    'metrics.average_cpc': '0',
    'metrics.cost_micros': '250000',
    'metrics.conversions': '0',
    'metrics.conversions_value': '0',
    'metrics.all_conversions': '0',
    'metrics.all_conversions_value': '0',
    ...overrides,
  };
}

const GEO_LOOKUPS = {
  2840: {
    'geo_target_constant.id': '2840',
    'geo_target_constant.resource_name': 'geoTargetConstants/2840',
    'geo_target_constant.name': 'United States',
    'geo_target_constant.canonical_name': 'United States',
    'geo_target_constant.country_code': 'US',
    'geo_target_constant.target_type': 'Country',
    'geo_target_constant.status': 'ENABLED',
  },
  1023191: {
    'geo_target_constant.id': '1023191',
    'geo_target_constant.resource_name': 'geoTargetConstants/1023191',
    'geo_target_constant.name': 'New York',
    'geo_target_constant.canonical_name': 'New York, New York, United States',
    'geo_target_constant.country_code': 'US',
    'geo_target_constant.target_type': 'City',
    'geo_target_constant.status': 'ENABLED',
  },
  21167: {
    'geo_target_constant.id': '21167',
    'geo_target_constant.resource_name': 'geoTargetConstants/21167',
    'geo_target_constant.name': 'New York',
    'geo_target_constant.canonical_name': 'New York, United States',
    'geo_target_constant.country_code': 'US',
    'geo_target_constant.target_type': 'State',
    'geo_target_constant.status': 'ENABLED',
  },
  1015010: {
    'geo_target_constant.id': '1015010',
    'geo_target_constant.resource_name': 'geoTargetConstants/1015010',
    'geo_target_constant.name': 'Hyde Park',
    'geo_target_constant.canonical_name': 'Hyde Park, Florida, United States',
    'geo_target_constant.country_code': 'US',
    'geo_target_constant.target_type': 'City',
    'geo_target_constant.status': 'ENABLED',
  },
  1015011: {
    'geo_target_constant.id': '1015011',
    'geo_target_constant.resource_name': 'geoTargetConstants/1015011',
    'geo_target_constant.name': 'Hyde Park',
    'geo_target_constant.canonical_name': 'Hyde Park, Florida, United States',
    'geo_target_constant.country_code': 'US',
    'geo_target_constant.target_type': 'Neighborhood',
    'geo_target_constant.status': 'ENABLED',
  },
  21142: {
    'geo_target_constant.id': '21142',
    'geo_target_constant.resource_name': 'geoTargetConstants/21142',
    'geo_target_constant.name': 'Florida',
    'geo_target_constant.canonical_name': 'Florida, United States',
    'geo_target_constant.country_code': 'US',
    'geo_target_constant.target_type': 'State',
    'geo_target_constant.status': 'ENABLED',
  },
};

function geoRuntime(mainRows, lookupRows = GEO_LOOKUPS) {
  const calls = [];
  return {
    calls,
    sleep() {},
    report(query, options) {
      calls.push({ query, options });
      if (/FROM\s+geographic_view\b/i.test(query)) return reportFor(mainRows);
      if (/FROM\s+geo_target_constant\b/i.test(query)) {
        const rows = idsFromLookupQuery(query)
          .map((id) => lookupRows[id])
          .filter(Boolean);
        return reportFor(rows);
      }
      throw new Error(`unexpected report query in campaign geography test:\n${query}`);
    },
  };
}

function dataRowsAsObjects(sheet) {
  const headers = sheet.data[0];
  return sheet.data.slice(1).map((values) => Object.fromEntries(
    headers.map((header, index) => [header, values[index]]),
  ));
}

test('only campaign_geo requests unresolved geographic IDs from AdsApp.report', () => {
  const { api } = loadExporter();
  const runCampaignGeoChunk = requireFunction(api, 'runCampaignGeoChunk');
  const geoJob = findJob(api, 'campaign_geo');
  const geoSheet = new MemorySheet(Array.from(api.headersForJob(geoJob)));
  const geoReportRuntime = geoRuntime([]);

  runCampaignGeoChunk(
    geoJob,
    AGGREGATE_RANGES,
    [CAMPAIGN_ID],
    geoSheet,
    geoReportRuntime,
    writerOptions(),
  );

  const geographicCall = geoReportRuntime.calls.find((call) => /FROM\s+geographic_view\b/i.test(call.query));
  assert.ok(geographicCall, 'campaign_geo must execute its geographic_view query');
  assert.equal(geographicCall.options.resolveGeoNames, false);

  const campaignJob = findJob(api, 'campaign');
  const campaignSheet = new MemorySheet(Array.from(api.headersForJob(campaignJob)));
  const ordinaryCalls = [];
  api.runGaqlChunk(
    campaignJob,
    AGGREGATE_RANGES,
    [CAMPAIGN_ID],
    campaignSheet,
    {
      sleep() {},
      report(query, options) {
        ordinaryCalls.push({ query, options });
        return reportFor([]);
      },
    },
    writerOptions(),
  );

  assert.equal(ordinaryCalls.length, 1);
  assert.notEqual(ordinaryCalls[0].options && ordinaryCalls[0].options.resolveGeoNames, false);
});

test('normalizes numeric and resource-shaped geo IDs without precision loss', () => {
  const { api } = loadExporter();
  const normalizeGeoCriterionId = requireFunction(api, 'normalizeGeoCriterionId');

  assert.equal(normalizeGeoCriterionId(2840), '2840');
  assert.equal(normalizeGeoCriterionId('1023191'), '1023191');
  assert.equal(normalizeGeoCriterionId('geoTargetConstants/1023191'), '1023191');
  assert.equal(
    normalizeGeoCriterionId('900719925474099312345'),
    '900719925474099312345',
  );
  assert.equal(
    normalizeGeoCriterionId('geoTargetConstants/900719925474099312345'),
    '900719925474099312345',
  );
  assert.equal(normalizeGeoCriterionId(''), '');
});

test('deduplicates geo lookup IDs and bounds every GAQL lookup batch at 500', () => {
  const { api } = loadExporter();
  const buildGeoTargetLookupQueries = requireFunction(api, 'buildGeoTargetLookupQueries');
  const ids = [];
  for (let id = 1000000; id <= 1001000; id += 1) ids.push(String(id));
  ids.push('1000000', 'geoTargetConstants/1000500');

  const queries = Array.from(buildGeoTargetLookupQueries(ids, 500));
  assert.equal(queries.length, 3);
  assert.deepEqual(queries.map((query) => idsFromLookupQuery(query).length), [500, 500, 1]);

  const emittedIds = queries.flatMap(idsFromLookupQuery);
  assert.equal(emittedIds.length, 1001);
  assert.equal(new Set(emittedIds).size, 1001);
  assert.ok(emittedIds.includes('1000000'));
  assert.ok(emittedIds.includes('1001000'));
});

test('same-name New York and Hyde Park rows remain distinct through stable ID keys', () => {
  const { api } = loadExporter();
  const runCampaignGeoChunk = requireFunction(api, 'runCampaignGeoChunk');
  const job = findJob(api, 'campaign_geo');
  const sheet = new MemorySheet(Array.from(api.headersForJob(job)));
  const runtime = geoRuntime([
    campaignGeoRow({
      'segments.geo_target_most_specific_location': '1023191',
      'segments.geo_target_state': '21167',
      'metrics.impressions': '8',
      'metrics.clicks': '1',
      'metrics.ctr': '0.125',
      'metrics.average_cpc': '1500000',
      'metrics.cost_micros': '1500000',
    }),
    campaignGeoRow({
      'segments.geo_target_most_specific_location': '21167',
      'segments.geo_target_state': '21167',
      'metrics.impressions': '16',
      'metrics.clicks': '4',
      'metrics.ctr': '0.25',
      'metrics.average_cpc': '700000',
      'metrics.cost_micros': '2800000',
    }),
    campaignGeoRow({
      'segments.geo_target_most_specific_location': '1015010',
      'segments.geo_target_state': '21142',
      'metrics.impressions': '5',
      'metrics.clicks': '0',
      'metrics.ctr': '0',
      'metrics.average_cpc': '0',
      'metrics.cost_micros': '0',
    }),
    campaignGeoRow({
      'segments.geo_target_most_specific_location': '1015011',
      'segments.geo_target_state': '21142',
      'metrics.impressions': '3',
      'metrics.clicks': '0',
      'metrics.ctr': '0',
      'metrics.average_cpc': '0',
      'metrics.cost_micros': '0',
    }),
  ]);

  const result = runCampaignGeoChunk(
    job,
    AGGREGATE_RANGES,
    [CAMPAIGN_ID],
    sheet,
    runtime,
    writerOptions(),
  );

  assert.deepEqual(JSON.parse(JSON.stringify(result)), {
    rows: 4,
    status: 'OK',
    limitation: '',
  });
  const rows = dataRowsAsObjects(sheet);
  assert.deepEqual(rows.map((row) => row['segments.geo_target_most_specific_location']), [
    'New York',
    'New York',
    'Hyde Park',
    'Hyde Park',
  ]);
  assert.deepEqual(rows.map((row) => row['metrics.impressions']), [8, 16, 5, 3]);

  const keys = rows.map((row) => EXPECTED_GEO_KEY_HEADERS.map((header) => row[header]).join('|'));
  assert.deepEqual(keys, [
    `${CAMPAIGN_ID}|LOCATION_OF_PRESENCE|2840|1023191|21167`,
    `${CAMPAIGN_ID}|LOCATION_OF_PRESENCE|2840|21167|21167`,
    `${CAMPAIGN_ID}|LOCATION_OF_PRESENCE|2840|1015010|21142`,
    `${CAMPAIGN_ID}|LOCATION_OF_PRESENCE|2840|1015011|21142`,
  ]);
  assert.equal(new Set(keys).size, 4);
});

test('missing geo labels preserve stable IDs and metrics while marking the chunk LIMITED', () => {
  const { api } = loadExporter();
  const runCampaignGeoChunk = requireFunction(api, 'runCampaignGeoChunk');
  const job = findJob(api, 'campaign_geo');
  const sheet = new MemorySheet(Array.from(api.headersForJob(job)));
  const runtime = geoRuntime([
    campaignGeoRow({
      'segments.geo_target_most_specific_location': '7777777777',
      'segments.geo_target_state': '21142',
      'metrics.impressions': '9',
      'metrics.clicks': '3',
      'metrics.ctr': '0.3333333333333333',
      'metrics.average_cpc': '411522.3333333333',
      'metrics.cost_micros': '1234567',
      'metrics.conversions': '1',
      'metrics.conversions_value': '25',
      'metrics.all_conversions': '1',
      'metrics.all_conversions_value': '25',
    }),
  ]);

  const result = runCampaignGeoChunk(
    job,
    AGGREGATE_RANGES,
    [CAMPAIGN_ID],
    sheet,
    runtime,
    writerOptions(),
  );

  assert.equal(result.rows, 1);
  assert.equal(result.status, 'LIMITED');
  assert.ok(String(result.limitation).length > 0);
  const [row] = dataRowsAsObjects(sheet);
  assert.equal(row.geo_target_most_specific_location_criterion_id, '7777777777');
  assert.equal(row['segments.geo_target_most_specific_location'], '');
  assert.equal(row['metrics.impressions'], 9);
  assert.equal(row['metrics.clicks'], 3);
  assert.equal(row['metrics.cost_micros'], 1234567);
  assert.equal(row['metrics.conversions'], 1);
});

test('a recognized unsupported geo label lookup preserves raw IDs as LIMITED output', () => {
  const { api } = loadExporter();
  const runCampaignGeoChunk = requireFunction(api, 'runCampaignGeoChunk');
  const job = findJob(api, 'campaign_geo');
  const sheet = new MemorySheet(Array.from(api.headersForJob(job)));
  const mainRows = [campaignGeoRow({
    'segments.geo_target_most_specific_location': '7777777777',
    'segments.geo_target_state': '21142',
    'metrics.impressions': '12',
    'metrics.clicks': '4',
    'metrics.cost_micros': '2222222',
  })];
  const runtime = {
    sleep() {},
    report(query, options) {
      if (/FROM\s+geographic_view\b/i.test(query)) {
        assert.equal(options.resolveGeoNames, false);
        return reportFor(mainRows);
      }
      if (/FROM\s+geo_target_constant\b/i.test(query)) {
        throw new Error('QueryError.UNSUPPORTED_RESOURCE');
      }
      throw new Error(`unexpected report query:\n${query}`);
    },
  };

  const result = runCampaignGeoChunk(
    job,
    AGGREGATE_RANGES,
    [CAMPAIGN_ID],
    sheet,
    runtime,
    writerOptions(),
  );

  assert.equal(result.rows, 1);
  assert.equal(result.status, 'LIMITED');
  assert.match(result.limitation, /lookup|metadata|readable/i);
  const [row] = dataRowsAsObjects(sheet);
  assert.equal(row.geo_target_most_specific_location_criterion_id, '7777777777');
  assert.equal(row.geo_target_state_criterion_id, '21142');
  assert.equal(row['metrics.impressions'], 12);
  assert.equal(row['metrics.clicks'], 4);
  assert.equal(row['metrics.cost_micros'], 2222222);
});

test('a transient geo label lookup failure remains terminal instead of LIMITED', () => {
  const { api } = loadExporter();
  const runCampaignGeoChunk = requireFunction(api, 'runCampaignGeoChunk');
  const job = findJob(api, 'campaign_geo');
  const sheet = new MemorySheet(Array.from(api.headersForJob(job)));
  const runtime = {
    sleep() {},
    report(query) {
      if (/FROM\s+geographic_view\b/i.test(query)) {
        return reportFor([campaignGeoRow()]);
      }
      if (/FROM\s+geo_target_constant\b/i.test(query)) {
        throw new Error('resource temporarily unavailable');
      }
      throw new Error(`unexpected report query:\n${query}`);
    },
  };

  assert.throws(
    () => runCampaignGeoChunk(
      job,
      AGGREGATE_RANGES,
      [CAMPAIGN_ID],
      sheet,
      runtime,
      writerOptions(),
    ),
    /temporarily unavailable/i,
  );
});

test('duplicate canonical campaign geography keys fail closed before publication', () => {
  const { api } = loadExporter();
  const runCampaignGeoChunk = requireFunction(api, 'runCampaignGeoChunk');
  const job = findJob(api, 'campaign_geo');
  const sheet = new MemorySheet(Array.from(api.headersForJob(job)));
  const runtime = geoRuntime([
    campaignGeoRow({ 'metrics.impressions': '6' }),
    campaignGeoRow({ 'metrics.impressions': '21', 'metrics.clicks': '2' }),
  ]);

  assert.throws(
    () => runCampaignGeoChunk(
      job,
      AGGREGATE_RANGES,
      [CAMPAIGN_ID],
      sheet,
      runtime,
      writerOptions(),
    ),
    /duplicate canonical geographic key/i,
  );
});

test('zero campaign geography rows complete OK without issuing lookup queries', () => {
  const { api } = loadExporter();
  const runCampaignGeoChunk = requireFunction(api, 'runCampaignGeoChunk');
  const job = findJob(api, 'campaign_geo');
  const sheet = new MemorySheet(Array.from(api.headersForJob(job)));
  const runtime = geoRuntime([]);

  const result = runCampaignGeoChunk(
    job,
    AGGREGATE_RANGES,
    [CAMPAIGN_ID],
    sheet,
    runtime,
    writerOptions(),
  );

  assert.deepEqual(JSON.parse(JSON.stringify(result)), {
    rows: 0,
    status: 'OK',
    limitation: '',
  });
  assert.equal(runtime.calls.filter((call) => /FROM\s+geo_target_constant\b/i.test(call.query)).length, 0);
  assert.equal(sheet.data.length, 1);
});

test('campaign_geo declares the exact stable ID key and every key is an output header', () => {
  const { api } = loadExporter();
  const job = findJob(api, 'campaign_geo');
  const headers = Array.from(api.headersForJob(job));
  const keyFields = Array.from(job.dictionary.keyFields);

  assert.deepEqual(keyFields, EXPECTED_GEO_KEY_HEADERS);
  assert.equal(new Set(keyFields).size, EXPECTED_GEO_KEY_HEADERS.length);
  keyFields.forEach((field) => assert.ok(headers.includes(field), `${field} must be an output header`));
  assert.ok(headers.includes('segments.geo_target_most_specific_location'));
  assert.ok(headers.includes('segments.geo_target_state'));
  assert.equal(api.validateManifest(api.getManifestDefinition()), true);
});

test('diagnostics probe raw campaign geography IDs and geo target label lookup compatibility', () => {
  const { api } = loadExporter();
  const probes = Array.from(api.diagnosticProbes({ start: '2026-08-19', end: '2026-08-25' }));
  const byName = Object.fromEntries(probes.map((probe) => [probe.name, probe]));

  assert.ok(byName.campaign_geo_raw_ids, 'diagnostics must probe unresolved campaign geography IDs');
  assert.equal(byName.campaign_geo_raw_ids.reportOptions.resolveGeoNames, false);
  assert.match(byName.campaign_geo_raw_ids.query, /FROM\s+geographic_view\b/);
  assert.match(byName.campaign_geo_raw_ids.query, /segments\.geo_target_most_specific_location/);
  assert.match(byName.campaign_geo_raw_ids.query, /segments\.geo_target_state/);

  assert.ok(byName.geo_target_constant_lookup, 'diagnostics must probe the readable-label lookup resource');
  assert.match(byName.geo_target_constant_lookup.query, /FROM\s+geo_target_constant\b/);
  assert.match(byName.geo_target_constant_lookup.query, /geo_target_constant\.id/);
  assert.match(byName.geo_target_constant_lookup.query, /geo_target_constant\.name/);
  assert.match(byName.geo_target_constant_lookup.query, /geo_target_constant\.canonical_name/);
});
