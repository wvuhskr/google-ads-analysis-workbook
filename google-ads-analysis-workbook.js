/**
 * SPDX-License-Identifier: MIT
 * Google Ads Analysis Workbook v1.0.0, reviewed for Google Ads API v25
 *
 * Creates a documented, resumable Google Sheet for one individual advertiser
 * account at a time. The workbook is structured for authorized human or
 * LLM-assisted analysis. The exporter is read-only with respect to Google Ads
 * and writes output and checkpoints to the configured Google Sheet.
 *
 * Normal entry point: main(). Preview automatically runs read-only diagnostics;
 * every live main() run exports or resumes the workbook. Optional direct entry
 * points: runDiagnostics() and resetExportState(). The completed native Google
 * Sheets workbook is the primary deliverable. A manual XLSX download is
 * supported only after the bundled local sanitizer validates the converted
 * OOXML package; Google Ads Scripts does not create or transport that file.
 *
 * This is not an MCC exporter, editable account backup, real-time or atomic
 * snapshot, automated optimizer, exhaustive export of every Google Ads
 * resource, direct AI integration, or anonymization tool. Do not edit, rename,
 * or delete workbook tabs while an export is running or resumable; the
 * transactional recovery contract assumes exporter-owned tabs remain intact
 * between main() invocations.
 *
 * Confidentiality: output includes account performance, search terms,
 * targeting, audience, and configuration data. Store and share it only under
 * the advertiser's authorization and your organization's data policy.
 */

var VERSION = 'v1.0.0';
var OWNER_KEY = 'google-ads-analysis-workbook';
var STATE_SCHEMA_VERSION = 1;
var OUTPUT_SCHEMA_VERSION = 9;
// Increment whenever executable behavior can change emitted rows without an
// accompanying VERSION or OUTPUT_SCHEMA_VERSION change. It prevents a paused
// checkpoint from resuming under a different implementation contract.
var RUNTIME_CONTRACT_VERSION = 10;
var STATE_SHEET_NAME = '_export_state';
var START_HERE_SHEET_NAME = 'START_HERE';
var INFO_SHEET_NAME = '_export_info';
var DICTIONARY_SHEET_NAME = '_data_dictionary';
var FIELD_DICTIONARY_SHEET_NAME = '_field_dictionary';
var STAGE_PREFIX = '__gads_export_stage__';
var BACKUP_PREFIX = '__gads_export_backup__';
var START_HERE_TAB_GROUPS = {
  performance: [
    'campaign', 'campaign_weekly', 'imp_share', 'keywords', 'search_terms', 'ads',
    'landing_pages', 'campaign_device_network', 'ad_schedule', 'campaign_geo',
    'ad_group_weekly', 'pmax_asset_group_weekly', 'ad_group', 'quality_score_keywords'
  ],
  structure: [
    'campaign_inventory', 'ad_group_inventory', 'keyword_inventory', 'ad_inventory',
    'ad_to_lp_map'
  ],
  creative: [
    'rsa_assets', 'demandgen_assets', 'pmax_asset_groups', 'pmax_assets', 'asset_extensions'
  ],
  targeting: ['geo_targets', 'geo_proximity_targets'],
  negatives: [
    'neg_keywords_campaign', 'neg_keywords_ad_group', 'neg_keywords_shared',
    'neg_keyword_shared_links', 'neg_keyword_account_links', 'negative_keywords_all'
  ],
  audiences: ['pmax_audience_signals', 'user_list_performance'],
  governance: [
    'conversion_actions', 'conversion_action_config', 'change_history', INFO_SHEET_NAME,
    DICTIONARY_SHEET_NAME, FIELD_DICTIONARY_SHEET_NAME
  ]
};
var GEO_TARGET_LOOKUP_BATCH_SIZE = 500;
var GOOGLE_ADS_MAX_EXECUTION_SECONDS = 1800;
var CONFIG = {
  SPREADSHEET_URL: 'INSERT-GOOGLE-SHEETS-URL-HERE',
  API_VERSION: 'v25',
  THROTTLE_MS: 700,
  BATCH_ROWS: 2000,
  CAMPAIGN_CHUNK_SIZE: 25,
  MIN_REMAINING_SECONDS: 180,
  MIN_COMMIT_REMAINING_SECONDS: 360,
  INFO_REFRESH_INTERVAL_SECONDS: 120,
  MAX_RESUME_AGE_HOURS: 24,
  INCLUDE_SENSITIVE_CHANGE_DETAILS: false,
  CHANGE_HISTORY_CLIENT_TYPES: ['GOOGLE_ADS_WEB_CLIENT'],
  DIAGNOSTICS_LOG_SAMPLE_ROWS: false,
  WORKBOOK_CELL_SAFETY_LIMIT: 9000000,
  FREEZE_CONTEXT_COLUMNS: 0,
  HIDE_RAW_MICROS_COLUMNS: true,
  ALLOW_RESET: false
};

// --------------------------------------------------------------------------
// Pure date, value, state, and validation helpers
// --------------------------------------------------------------------------

function parseYmdUtc_(ymd) {
  var match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || ''));
  if (!match) throw new Error('Expected YYYY-MM-DD; received: ' + ymd);
  var date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  if (formatYmdUtc_(date) !== ymd) throw new Error('Invalid calendar date: ' + ymd);
  return date;
}

function formatYmdUtc_(date) {
  function two_(value) { return value < 10 ? '0' + value : String(value); }
  return date.getUTCFullYear() + '-' + two_(date.getUTCMonth() + 1) + '-' + two_(date.getUTCDate());
}

function addDaysYmd_(ymd, days) {
  var date = parseYmdUtc_(ymd);
  date.setUTCDate(date.getUTCDate() + Number(days));
  return formatYmdUtc_(date);
}

function buildFrozenRanges_(todayYmd) {
  var yesterday = addDaysYmd_(todayYmd, -1);
  var aggregate = { start: addDaysYmd_(yesterday, -89), end: yesterday };
  return {
    aggregate: aggregate,
    weekly: { start: aggregate.start, end: aggregate.end },
    change: { start: addDaysYmd_(yesterday, -27), end: yesterday }
  };
}

function encodeSheetCellForWrite_(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number' || typeof value === 'boolean') return value;
  var text = String(value);
  // Sheets consumes the first apostrophe as a literal-text marker, preserves the
  // displayed value, and prevents automatic formulas or links. Do not rewrite
  // these cells through RichText later because that removes quotePrefix.
  return text === '' ? '' : "'" + text;
}

function encodeSheetRowForWrite_(row, headersOrWidth) {
  if (!Array.isArray(row)) throw new Error('Report row must be an array.');
  var expected = Array.isArray(headersOrWidth) ? headersOrWidth.length : Number(headersOrWidth);
  if (row.length !== expected) {
    throw new Error('Report row width ' + row.length + ' did not match expected ' + expected + '.');
  }
  return row.map(encodeSheetCellForWrite_);
}

function stableStringify_(value) {
  if (value === null || typeof value !== 'object') return JSON.stringify(value);
  if (Array.isArray(value)) return '[' + value.map(stableStringify_).join(',') + ']';
  var keys = Object.keys(value).sort();
  return '{' + keys.map(function(key) {
    return JSON.stringify(key) + ':' + stableStringify_(value[key]);
  }).join(',') + '}';
}

function compactStateForStorage_(state, maxCharacters) {
  var limit = Number(maxCharacters || 45000);
  if (!isFinite(limit) || limit < 1000) throw new Error('Invalid checkpoint character limit.');
  var compact = JSON.parse(JSON.stringify(state));
  Object.keys(compact.tabs || {}).forEach(function(tab) {
    var result = compact.tabs[tab];
    if (!result.error) delete result.error;
    if (!result.limitation) delete result.limitation;
    if (!result.partialLimited) delete result.partialLimited;
    if (!result.priorPreserved) delete result.priorPreserved;
    if (!result.durationMs) delete result.durationMs;
    if (!result.sourceReadStartedAtMs) delete result.sourceReadStartedAtMs;
    if (!result.sourceReadCompletedAtMs) delete result.sourceReadCompletedAtMs;
  });
  if (JSON.stringify(compact).length <= limit) return compact;
  [160, 80, 40, 12, 0].some(function(messageLimit) {
    Object.keys(compact.tabs || {}).forEach(function(tab) {
      var result = compact.tabs[tab];
      ['error', 'limitation'].forEach(function(field) {
        if (!result[field]) return;
        if (messageLimit === 0) delete result[field];
        else result[field] = String(result[field]).substring(0, messageLimit);
      });
    });
    return JSON.stringify(compact).length <= limit;
  });
  if (JSON.stringify(compact).length > limit) {
    Object.keys(compact.tabs || {}).forEach(function(tab) {
      delete compact.tabs[tab].durationMs;
    });
  }
  if (JSON.stringify(compact).length > limit) {
    throw new Error('Resume checkpoint cannot fit within the reserved Google Sheets state cell.');
  }
  return compact;
}

function normalizeTabResult_(result) {
  result.status = result.status || 'RUNNING';
  result.rows = Number(result.rows || 0);
  result.durationMs = Number(result.durationMs || 0);
  result.sourceReadStartedAtMs = Number(result.sourceReadStartedAtMs || 0);
  result.sourceReadCompletedAtMs = Number(result.sourceReadCompletedAtMs || 0);
  result.error = String(result.error || '');
  result.limitation = String(result.limitation || '');
  result.partialLimited = Boolean(result.partialLimited);
  result.priorPreserved = Boolean(result.priorPreserved);
  return result;
}

function validateManifest_(manifest) {
  if (!Array.isArray(manifest) || manifest.length === 0) {
    throw new Error('The export manifest must contain at least one job.');
  }
  var ids = {};
  var tabs = {};
  manifest.forEach(function(job) {
    if (!job || !job.id || !job.tab) throw new Error('Every manifest job requires id and tab.');
    if (ids[job.id]) throw new Error('Duplicate job id: ' + job.id);
    if (tabs[job.tab]) throw new Error('Duplicate final tab: ' + job.tab);
    ids[job.id] = true;
    tabs[job.tab] = true;
    if (Array.isArray(job.columns) && job.dictionary) {
      var headers = headersForJob_(job);
      var keyFields = job.dictionary.keyFields;
      if (!Array.isArray(keyFields) || !keyFields.length) {
        throw new Error('Manifest job ' + job.id + ' requires dictionary key fields.');
      }
      var seenKeys = {};
      keyFields.forEach(function(field) {
        if (seenKeys[field]) throw new Error('Duplicate dictionary key for ' + job.id + ': ' + field);
        if (headers.indexOf(field) < 0) {
          throw new Error('Dictionary key for ' + job.id + ' is not an output header: ' + field);
        }
        seenKeys[field] = true;
      });
    }
    if (job.campaignChunkSize !== undefined) {
      var campaignChunkSize = Number(job.campaignChunkSize);
      if (!job.chunked || !isFinite(campaignChunkSize) || campaignChunkSize < 1 ||
          Math.floor(campaignChunkSize) !== campaignChunkSize) {
        throw new Error('Manifest job ' + job.id + ' has an invalid campaignChunkSize.');
      }
    }
  });
  return true;
}

function assertWorkbookOwnership_(sheetSummaries) {
  var summaries = sheetSummaries || [];
  var owned = summaries.some(function(summary) {
    return (summary.name === INFO_SHEET_NAME || summary.name === STATE_SHEET_NAME) &&
      summary.marker === OWNER_KEY;
  });
  if (owned) return 'owned';
  if (summaries.every(function(summary) { return summary.blank; })) return 'blank';
  throw new Error(
    'The target workbook is populated but is not owned by ' + OWNER_KEY +
    '. Use a blank, dedicated workbook.'
  );
}

function createRunState_(identity, nowMs, ranges, manifestIds) {
  var outputSchemaVersion = identity.outputSchemaVersion === undefined ?
    OUTPUT_SCHEMA_VERSION : identity.outputSchemaVersion;
  var runtimeContractVersion = identity.runtimeContractVersion === undefined ?
    RUNTIME_CONTRACT_VERSION : identity.runtimeContractVersion;
  return {
    schemaVersion: STATE_SCHEMA_VERSION,
    outputSchemaVersion: outputSchemaVersion,
    runtimeContractVersion: runtimeContractVersion,
    runId: 'run-' + String(nowMs),
    version: identity.version,
    accountId: identity.accountId,
    spreadsheetId: identity.spreadsheetId,
    configSignature: identity.configSignature,
    startedAtMs: nowMs,
    updatedAtMs: nowMs,
    ranges: ranges,
    manifest: manifestIds.slice(),
    status: 'RUNNING',
    jobIndex: 0,
    currentJobId: '',
    chunkIndex: 0,
    chunkInProgress: false,
    chunkStartRow: 0,
    stageSheetName: '',
    jobPhase: '',
    tabs: {}
  };
}

function assertStateCompatible_(state, identity, nowMs, maxResumeAgeHours) {
  if (!state || state.schemaVersion !== STATE_SCHEMA_VERSION) {
    throw new Error('Saved export state has an unsupported schema. Reset it before continuing.');
  }
  var expectedOutputSchemaVersion = identity.outputSchemaVersion === undefined ?
    OUTPUT_SCHEMA_VERSION : identity.outputSchemaVersion;
  if (state.outputSchemaVersion !== expectedOutputSchemaVersion) {
    throw new Error(
      'Saved export state does not match the current output schema version. Reset it before continuing.'
    );
  }
  var expectedRuntimeContractVersion = identity.runtimeContractVersion === undefined ?
    RUNTIME_CONTRACT_VERSION : identity.runtimeContractVersion;
  if (state.runtimeContractVersion !== expectedRuntimeContractVersion) {
    throw new Error(
      'Saved export state does not match the current runtime contract version. Reset it before continuing.'
    );
  }
  [
    ['version', 'script version'],
    ['accountId', 'account'],
    ['spreadsheetId', 'workbook'],
    ['configSignature', 'configuration']
  ].forEach(function(pair) {
    if (state[pair[0]] !== identity[pair[0]]) {
      throw new Error('Saved export state does not match the current ' + pair[1] + '.');
    }
  });
  var maxAgeMs = Number(maxResumeAgeHours) * 60 * 60 * 1000;
  if (!isFinite(maxAgeMs) || maxAgeMs <= 0) throw new Error('MAX_RESUME_AGE_HOURS must be positive.');
  if ((Number(nowMs) - Number(state.startedAtMs)) > maxAgeMs) {
    throw new Error('Saved export state has expired. Set ALLOW_RESET=true and run resetExportState().');
  }
  return true;
}

function planReset_(allowReset, sheetNames) {
  if (allowReset !== true) {
    throw new Error('Reset refused. Set CONFIG.ALLOW_RESET=true before calling resetExportState().');
  }
  return (sheetNames || []).filter(function(name) {
    return name === STATE_SHEET_NAME ||
      name.indexOf(STAGE_PREFIX) === 0 ||
      name.indexOf(BACKUP_PREFIX) === 0;
  });
}

function chunkCampaignIds_(campaignIds, chunkSize) {
  var size = Number(chunkSize);
  if (!isFinite(size) || size < 1 || Math.floor(size) !== size) {
    throw new Error('CAMPAIGN_CHUNK_SIZE must be a positive integer.');
  }
  var unique = {};
  (campaignIds || []).forEach(function(id) {
    var value = String(id);
    if (!/^\d+$/.test(value)) throw new Error('Invalid campaign ID: ' + value);
    unique[value] = true;
  });
  var sorted = Object.keys(unique).sort(function(left, right) {
    if (left.length !== right.length) return left.length - right.length;
    return left < right ? -1 : (left > right ? 1 : 0);
  });
  var chunks = [];
  for (var index = 0; index < sorted.length; index += size) {
    chunks.push(sorted.slice(index, index + size));
  }
  return chunks;
}

function chunkCampaignIdsForJob_(job, campaignIds, defaultChunkSize) {
  if (!job || !job.chunked) return [null];
  var chunkSize = job.campaignChunkSize === undefined ? defaultChunkSize : job.campaignChunkSize;
  return chunkCampaignIds_(campaignIds, chunkSize);
}

function normalizeExecutionContractValue_(value) {
  if (value === undefined) return '[undefined]';
  if (typeof value === 'function') return String(value);
  if (value === null || typeof value !== 'object') return value;
  if (Array.isArray(value)) return value.map(normalizeExecutionContractValue_);
  var normalized = {};
  Object.keys(value).sort().forEach(function(key) {
    normalized[key] = normalizeExecutionContractValue_(value[key]);
  });
  return normalized;
}

function compactTextFingerprint_(text) {
  var input = String(text || '');
  var fnv = 2166136261;
  var djb = 5381;
  for (var index = 0; index < input.length; index++) {
    var code = input.charCodeAt(index);
    fnv ^= code;
    fnv += (fnv << 1) + (fnv << 4) + (fnv << 7) + (fnv << 8) + (fnv << 24);
    djb = ((djb << 5) + djb) ^ code;
  }
  function hex_(value) {
    return ('00000000' + (value >>> 0).toString(16)).slice(-8);
  }
  return input.length + '-' + hex_(fnv) + '-' + hex_(djb);
}

function manifestExecutionSignature_(manifest, defaultChunkSize) {
  var contract = (manifest || []).map(function(job) {
    return {
      definition: normalizeExecutionContractValue_(job),
      effectiveCampaignChunkSize: job && job.chunked ? Number(
        job.campaignChunkSize === undefined ? defaultChunkSize : job.campaignChunkSize
      ) : 0,
      headers: job && Array.isArray(job.columns) ? headersForJob_(job) : []
    };
  });
  return compactTextFingerprint_(stableStringify_(contract));
}

// --------------------------------------------------------------------------
// Declarative Google Ads report catalog
// --------------------------------------------------------------------------

function column_(field, type, header, metadata) {
  var definition = { field: field, type: type || 'text', header: header || field };
  Object.keys(metadata || {}).forEach(function(key) { definition[key] = metadata[key]; });
  return definition;
}

function sourceColumn_(field, type, header, metadata) {
  var details = metadata || {};
  if (details.isDerived === undefined) details.isDerived = true;
  if (!details.derivation) details.derivation = 'Joined by the exporter from a supporting Google Ads query';
  var column = column_(field, type, header, details);
  column.select = false;
  return column;
}

function derivedOutputColumn_(header, type, sourceFields, derivation) {
  return column_(header, type, header, {
    sourceFields: (sourceFields || []).slice(),
    isDerived: true,
    derivation: derivation || 'Composed by the exporter'
  });
}

function customerColumns_() {
  return [
    column_('customer.id', 'id'),
    column_('customer.descriptive_name'),
    column_('customer.currency_code'),
    column_('customer.time_zone')
  ];
}

function campaignColumns_() {
  return [
    column_('campaign.id', 'id'),
    column_('campaign.name'),
    column_('campaign.status'),
    column_('campaign.advertising_channel_type'),
    column_('campaign.advertising_channel_sub_type')
  ];
}

function adGroupColumns_() {
  return [
    column_('ad_group.id', 'id'),
    column_('ad_group.name'),
    column_('ad_group.status'),
    column_('ad_group.type')
  ];
}

function adGroupEligibilityLookup_() {
  return {
    resource: 'ad_group',
    keyFields: ['campaign.id', 'ad_group.id'],
    statusFields: ['campaign.status', 'ad_group.status']
  };
}

function assetGroupEligibilityLookup_() {
  return {
    resource: 'asset_group',
    keyFields: ['campaign.id', 'asset_group.id'],
    statusFields: ['campaign.status', 'asset_group.status'],
    where: ["campaign.advertising_channel_type = 'PERFORMANCE_MAX'"]
  };
}

function performanceColumns_() {
  return [
    column_('metrics.impressions', 'number'),
    column_('metrics.clicks', 'number'),
    column_('metrics.interactions', 'number'),
    column_('metrics.ctr', 'number'),
    column_('metrics.average_cpc', 'number'),
    column_('metrics.cost_micros', 'number'),
    column_('metrics.conversions', 'number'),
    column_('metrics.conversions_value', 'number'),
    column_('metrics.all_conversions', 'number'),
    column_('metrics.all_conversions_value', 'number')
  ];
}

function activityMetricFields_() {
  return [
    'metrics.impressions',
    'metrics.clicks',
    'metrics.interactions',
    'metrics.cost_micros',
    'metrics.conversions',
    'metrics.conversions_value',
    'metrics.all_conversions',
    'metrics.all_conversions_value'
  ];
}

function hasSurveyActivity_(row) {
  return activityMetricFields_().some(function(field) {
    return metricNumber_(row || {}, field) !== 0;
  });
}

function isInactiveStatus_(value) {
  if (value === false) return true;
  var normalized = String(value === null || value === undefined ? '' : value).toUpperCase();
  return normalized === 'FALSE' || normalized === 'PAUSED' ||
    normalized === 'REMOVED' || normalized === 'DISABLED';
}

function entityKey_(row, keyFields) {
  return (keyFields || []).map(function(field) {
    return String((row || {})[field] === undefined || (row || {})[field] === null ? '' : (row || {})[field]);
  }).join('|');
}

function indexRowsByEntityKey_(rows, keyFields) {
  var index = {};
  (rows || []).forEach(function(row) {
    var key = entityKey_(row, keyFields);
    if (!index[key]) index[key] = row;
  });
  return index;
}

function shouldIncludeSurveyEntity_(row, statusFields, activeKeys, keyFields) {
  var inactive = (statusFields || []).some(function(field) {
    return isInactiveStatus_((row || {})[field]);
  });
  if (!inactive || hasSurveyActivity_(row || {})) return true;
  if (activeKeys && keyFields && keyFields.length) {
    return activeKeys[entityKey_(row, keyFields)] === true;
  }
  return false;
}

function buildCampaignEligibilityQueries_(range) {
  if (!range || !range.start || !range.end) {
    throw new Error('Campaign eligibility requires the frozen aggregate date range.');
  }
  return {
    current: 'SELECT\n  campaign.id\nFROM campaign\n' +
      "WHERE campaign.status NOT IN ('PAUSED', 'REMOVED')\nORDER BY campaign.id",
    activity: 'SELECT\n  campaign.id,\n  ' + activityMetricFields_().join(',\n  ') +
      '\nFROM campaign\nWHERE segments.date BETWEEN \'' + range.start + "' AND '" + range.end +
      "'\nORDER BY campaign.id"
  };
}

function collectEligibleCampaignIds_(currentRows, activityRows) {
  var ids = [];
  (currentRows || []).forEach(function(row) { ids.push(String(row['campaign.id'])); });
  (activityRows || []).forEach(function(row) {
    if (hasSurveyActivity_(row)) ids.push(String(row['campaign.id']));
  });
  var chunks = chunkCampaignIds_(ids, Math.max(1, ids.length || 1));
  return chunks.length ? chunks[0] : [];
}

function standardDerivedColumns_() {
  return [
    {
      header: 'cost',
      sourceFields: ['metrics.cost_micros'],
      derivation: 'metrics.cost_micros divided by 1,000,000',
      compute: function(row) { return metricNumber_(row, 'metrics.cost_micros') / 1000000; }
    },
    {
      header: 'average_cpc',
      sourceFields: ['metrics.average_cpc'],
      derivation: 'metrics.average_cpc divided by 1,000,000',
      compute: function(row) {
        var raw = row['metrics.average_cpc'];
        if (raw === '' || raw === null || raw === undefined) return '';
        var number = Number(String(raw).replace(/,/g, ''));
        return isFinite(number) ? number / 1000000 : '';
      }
    },
    {
      header: 'cost_per_conversion',
      sourceFields: ['metrics.cost_micros', 'metrics.conversions'],
      derivation: 'Cost in account currency divided by metrics.conversions; blank when conversions are zero',
      compute: function(row) {
      var conversions = metricNumber_(row, 'metrics.conversions');
      return conversions ? (metricNumber_(row, 'metrics.cost_micros') / 1000000) / conversions : '';
      }
    },
    {
      header: 'conversion_rate',
      sourceFields: ['metrics.conversions', 'metrics.interactions'],
      derivation: 'metrics.conversions divided by metrics.interactions; blank when interactions are zero',
      compute: function(row) {
      var interactions = metricNumber_(row, 'metrics.interactions');
      return interactions ? metricNumber_(row, 'metrics.conversions') / interactions : '';
      }
    },
    {
      header: 'conversion_value_per_cost',
      sourceFields: ['metrics.conversions_value', 'metrics.cost_micros'],
      derivation: 'metrics.conversions_value divided by cost in account currency; blank when cost is zero',
      compute: function(row) {
      var cost = metricNumber_(row, 'metrics.cost_micros') / 1000000;
      return cost ? metricNumber_(row, 'metrics.conversions_value') / cost : '';
      }
    }
  ];
}

function microsDerivedColumns_(definitions) {
  return definitions.map(function(definition) {
    var field = definition[0];
    var header = definition[1];
    return {
      header: header,
      sourceFields: [field],
      derivation: field + ' divided by 1,000,000',
      compute: function(row) {
        var raw = row[field];
        if (raw === '' || raw === null || raw === undefined) return '';
        var number = Number(String(raw).replace(/,/g, ''));
        return isFinite(number) ? number / 1000000 : '';
      }
    };
  });
}

function reportJob_(definition) {
  definition.kind = definition.kind || 'gaql';
  definition.required = definition.required !== false;
  definition.derived = definition.derived || [];
  definition.where = definition.where || [];
  definition.dictionary = definition.dictionary || {};
  return definition;
}

function dictionaryKeyFields_() {
  return {
    campaign: ['campaign.id'],
    campaign_weekly: ['campaign.id', 'segments.week'],
    imp_share: ['campaign.id'],
    keywords: [
      'campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id',
      'segments.device', 'segments.ad_network_type'
    ],
    search_terms: [
      'campaign.id', 'ad_group.id', 'search_term_view.search_term',
      'segments.keyword.info.text', 'segments.keyword.info.match_type'
    ],
    ads: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
    rsa_assets: [
      'campaign.id', 'ad_group.id', 'ad_group_ad.ad.id', 'asset.id',
      'ad_group_ad_asset_view.field_type'
    ],
    demandgen_assets: [
      'campaign.id', 'ad_group.id', 'ad_group_ad.ad.id', 'asset.id',
      'ad_group_ad_asset_view.field_type'
    ],
    landing_pages: [
      'campaign.id', 'landing_page_view.unexpanded_final_url',
      'segments.ad_network_type', 'segments.device'
    ],
    campaign_device_network: ['campaign.id', 'segments.device', 'segments.ad_network_type'],
    ad_schedule: ['campaign.id', 'segments.day_of_week', 'segments.hour'],
    campaign_geo: [
      'campaign.id', 'geographic_view.location_type', 'geographic_view.country_criterion_id',
      'geo_target_most_specific_location_criterion_id', 'geo_target_state_criterion_id'
    ],
    ad_group_weekly: ['campaign.id', 'ad_group.id', 'segments.week'],
    pmax_asset_group_weekly: ['campaign.id', 'asset_group.id', 'segments.week'],
    ad_group: ['campaign.id', 'ad_group.id'],
    campaign_inventory: ['campaign.id'],
    ad_group_inventory: ['campaign.id', 'ad_group.id'],
    keyword_inventory: ['campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id'],
    ad_inventory: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
    ad_to_lp_map: ['campaign.id', 'ad_group.id', 'ad_id', 'url_source', 'final_url_raw'],
    geo_targets: ['campaign.id', 'campaign_criterion.criterion_id'],
    geo_proximity_targets: ['campaign.id', 'campaign_criterion.criterion_id'],
    neg_keywords_campaign: ['campaign.id', 'campaign_criterion.criterion_id'],
    neg_keywords_ad_group: ['campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id'],
    neg_keywords_shared: ['shared_set.id', 'shared_criterion.criterion_id'],
    neg_keyword_shared_links: ['campaign.id', 'shared_set.id'],
    neg_keyword_account_links: ['customer_negative_criterion.resource_name'],
    negative_keywords_all: [
      'source', 'scope', 'campaign.id', 'ad_group.id', 'criterion_id',
      'shared_set.id', 'keyword.text', 'keyword.match_type'
    ],
    conversion_actions: ['segments.conversion_action'],
    conversion_action_config: ['conversion_action.id'],
    quality_score_keywords: ['campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id'],
    pmax_asset_groups: ['campaign.id', 'asset_group.id'],
    pmax_assets: [
      'campaign.id', 'asset_group.id', 'asset.id', 'asset_group_asset.field_type'
    ],
    pmax_audience_signals: ['asset_group_signal.resource_name'],
    user_list_performance: [
      'scope', 'user_list_resource', 'criterion_id', 'campaign.id', 'ad_group.id'
    ],
    asset_extensions: ['scope', 'campaign.id', 'ad_group.id', 'asset.id', 'field_type'],
    change_history: ['change_event_resource_name'],
    _data_dictionary: ['tab'],
    _field_dictionary: ['tab', 'field']
  };
}

function getManifestDefinition_() {
  var customer = customerColumns_();
  var campaign = campaignColumns_();
  var adGroup = adGroupColumns_();
  var metrics = performanceColumns_();
  var derived = standardDerivedColumns_();
  var campaignBudgetDerived = microsDerivedColumns_([
    ['campaign_budget.amount_micros', 'campaign_budget.amount'],
    ['campaign_budget.recommended_budget_amount_micros', 'campaign_budget.recommended_budget_amount'],
    ['campaign_budget.total_amount_micros', 'campaign_budget.total_amount'],
    ['campaign.maximize_conversions.target_cpa_micros', 'campaign.maximize_conversions.target_cpa'],
    ['campaign.target_cpa.target_cpa_micros', 'campaign.target_cpa.target_cpa']
  ]);
  var campaignWeeklyDerived = microsDerivedColumns_([
    ['campaign_budget.amount_micros', 'campaign_budget.amount'],
    ['campaign.maximize_conversions.target_cpa_micros', 'campaign.maximize_conversions.target_cpa'],
    ['campaign.target_cpa.target_cpa_micros', 'campaign.target_cpa.target_cpa']
  ]);
  var adGroupTargetDerived = microsDerivedColumns_([
    ['ad_group.target_cpa_micros', 'ad_group.target_cpa']
  ]);
  var adGroupInventoryDerived = microsDerivedColumns_([
    ['ad_group.cpc_bid_micros', 'ad_group.cpc_bid'],
    ['ad_group.cpm_bid_micros', 'ad_group.cpm_bid'],
    ['ad_group.target_cpa_micros', 'ad_group.target_cpa']
  ]);
  var negativeUnionSources = [
    'neg_keywords_campaign', 'neg_keywords_ad_group', 'neg_keywords_shared',
    'neg_keyword_shared_links', 'neg_keyword_account_links'
  ];
  var negativeUnionDerivation =
    'Unified with explicit source and scope from the direct, shared-list, and account-list tabs';
  var manifest = [
    reportJob_({
      id: 'campaign', tab: 'campaign', resource: 'campaign', chunked: true, range: 'aggregate',
      columns: customer.concat(campaign, [
        column_('campaign_budget.id', 'id'), column_('campaign_budget.name'),
        column_('campaign_budget.amount_micros', 'number'), column_('campaign_budget.period'),
        column_('campaign_budget.has_recommended_budget', 'boolean'),
        column_('campaign_budget.recommended_budget_amount_micros', 'number'),
        column_('campaign_budget.total_amount_micros', 'number'),
        column_('campaign.maximize_conversions.target_cpa_micros', 'number'),
        column_('campaign.target_cpa.target_cpa_micros', 'number'),
        column_('campaign.bidding_strategy_type'), column_('campaign.optimization_score', 'number')
      ], metrics), derived: derived.concat(campaignBudgetDerived),
      orderBy: ['campaign.id'],
      dictionary: { purpose: 'Campaign performance', grain: 'campaign', keys: 'campaign.id' }
    }),
    reportJob_({
      id: 'campaign_weekly', tab: 'campaign_weekly', resource: 'campaign', chunked: true, range: 'weekly',
      columns: customer.concat(campaign, [
        column_('campaign_budget.amount_micros', 'number'),
        column_('campaign.maximize_conversions.target_cpa_micros', 'number'),
        column_('campaign.target_cpa.target_cpa_micros', 'number'),
        column_('segments.week', 'date')
      ], metrics), derived: derived.concat(campaignWeeklyDerived),
      orderBy: ['segments.week', 'campaign.id'],
      dictionary: { purpose: 'Weekly campaign trend', grain: 'campaign and week bucket within the 90-day window', keys: 'campaign.id, segments.week' }
    }),
    reportJob_({
      id: 'imp_share', tab: 'imp_share', resource: 'campaign', chunked: true, range: 'aggregate', required: false,
      columns: customer.concat(campaign, [
        column_('metrics.search_impression_share', 'number'),
        column_('metrics.search_top_impression_share', 'number'),
        column_('metrics.search_absolute_top_impression_share', 'number'),
        column_('metrics.search_budget_lost_impression_share', 'number'),
        column_('metrics.search_rank_lost_impression_share', 'number'),
        column_('metrics.content_impression_share', 'number'),
        column_('metrics.content_budget_lost_impression_share', 'number'),
        column_('metrics.content_rank_lost_impression_share', 'number')
      ]),
      where: ["campaign.advertising_channel_type IN ('SEARCH', 'DISPLAY')"],
      orderBy: ['campaign.id'],
      dictionary: { purpose: 'Search and display impression share', grain: 'campaign', keys: 'campaign.id' }
    }),
    reportJob_({
      id: 'keywords', tab: 'keywords', resource: 'keyword_view', chunked: true, range: 'aggregate',
      surveyStatusFields: ['campaign.status', 'ad_group.status', 'ad_group_criterion.status'],
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_criterion.criterion_id', 'id'),
        column_('ad_group_criterion.keyword.text'),
        column_('ad_group_criterion.keyword.match_type'),
        column_('ad_group_criterion.status'),
        column_('segments.device'), column_('segments.ad_network_type')
      ], metrics), derived: derived,
      orderBy: ['campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id'],
      dictionary: { purpose: 'Keyword performance by device and network', grain: 'keyword, device, and network', keys: 'ad_group.id, criterion_id' }
    }),
    reportJob_({
      id: 'search_terms', tab: 'search_terms', resource: 'search_term_view', chunked: true, range: 'aggregate',
      columns: customer.concat(campaign, adGroup, [
        column_('search_term_view.search_term'), column_('search_term_view.status'),
        column_('segments.keyword.info.text'), column_('segments.keyword.info.match_type')
      ], metrics), derived: derived,
      orderBy: ['campaign.id', 'ad_group.id', 'search_term_view.search_term'],
      dictionary: { purpose: 'Google-reported search queries', grain: 'search term and matched keyword context', keys: 'campaign.id, ad_group.id, search term' }
    }),
    reportJob_({
      id: 'ads', tab: 'ads', resource: 'ad_group_ad', chunked: true, range: 'aggregate',
      surveyStatusFields: ['campaign.status', 'ad_group.status', 'ad_group_ad.status'],
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_ad.ad.id', 'id'), column_('ad_group_ad.status'),
        column_('ad_group_ad.ad.name'), column_('ad_group_ad.ad.type'),
        column_('ad_group_ad.ad.final_urls'), column_('ad_group_ad.ad.final_mobile_urls'),
        column_('ad_group_ad.ad.tracking_url_template')
      ], metrics), derived: derived,
      orderBy: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
      dictionary: { purpose: 'Ad performance', grain: 'ad', keys: 'ad_group_ad.ad.id' }
    }),
    reportJob_({
      id: 'rsa_assets', tab: 'rsa_assets', resource: 'ad_group_ad_asset_view', chunked: true,
      range: 'aggregate', required: false,
      surveyStatusFields: [
        'campaign.status', 'ad_group.status', 'ad_group_ad.status', 'ad_group_ad_asset_view.enabled'
      ],
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_ad.ad.id', 'id'), column_('ad_group_ad.status'),
        column_('ad_group_ad_asset_view.field_type'),
        column_('ad_group_ad_asset_view.performance_label'), column_('ad_group_ad_asset_view.pinned_field'),
        column_('ad_group_ad_asset_view.enabled', 'boolean'), column_('asset.id', 'id'),
        column_('asset.type'), column_('asset.text_asset.text')
      ], metrics), derived: derived,
      where: [
        "campaign.advertising_channel_type = 'SEARCH'",
        "ad_group_ad.ad.type = 'RESPONSIVE_SEARCH_AD'"
      ],
      orderBy: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id', 'asset.id'],
      dictionary: { purpose: 'Responsive Search Ad asset performance', grain: 'RSA asset association', keys: 'ad id, asset.id, field_type' }
    }),
    reportJob_({
      id: 'demandgen_assets', tab: 'demandgen_assets', resource: 'ad_group_ad_asset_view',
      chunked: true, range: 'aggregate', required: false,
      surveyStatusFields: [
        'campaign.status', 'ad_group.status', 'ad_group_ad.status', 'ad_group_ad_asset_view.enabled'
      ],
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_ad.ad.id', 'id'), column_('ad_group_ad.status'),
        column_('ad_group_ad.ad.type'), column_('ad_group_ad_asset_view.enabled', 'boolean'),
        column_('ad_group_ad_asset_view.field_type'), column_('ad_group_ad_asset_view.performance_label'),
        column_('asset.id', 'id'), column_('asset.name'), column_('asset.type'),
        column_('asset.text_asset.text')
      ], metrics), derived: derived,
      where: ["campaign.advertising_channel_type = 'DEMAND_GEN'"],
      orderBy: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id', 'asset.id'],
      dictionary: { purpose: 'Demand Gen ad asset performance', grain: 'ad asset association', keys: 'ad id, asset.id, field_type' }
    }),
    reportJob_({
      id: 'landing_pages', tab: 'landing_pages', resource: 'landing_page_view', chunked: true, range: 'aggregate',
      columns: customer.concat(campaign, [
        column_('landing_page_view.unexpanded_final_url'),
        column_('segments.ad_network_type'), column_('segments.device')
      ], metrics), derived: derived,
      orderBy: ['campaign.id', 'landing_page_view.unexpanded_final_url'],
      dictionary: { purpose: 'Landing-page performance', grain: 'campaign, URL, network, and device', keys: 'campaign.id, URL, network, device' }
    }),
    reportJob_({
      id: 'campaign_device_network', tab: 'campaign_device_network', resource: 'campaign',
      chunked: true, range: 'aggregate',
      columns: customer.concat(campaign, [column_('segments.device'), column_('segments.ad_network_type')], metrics),
      derived: derived, orderBy: ['campaign.id', 'segments.device', 'segments.ad_network_type'],
      dictionary: { purpose: 'Campaign performance by device and network', grain: 'campaign, device, and network', keys: 'campaign.id, device, network' }
    }),
    reportJob_({
      id: 'ad_schedule', tab: 'ad_schedule', resource: 'campaign', chunked: true, range: 'aggregate',
      columns: customer.concat(campaign, [column_('segments.day_of_week'), column_('segments.hour', 'number')], metrics),
      derived: derived, orderBy: ['campaign.id', 'segments.day_of_week', 'segments.hour'],
      dictionary: { purpose: 'Campaign daypart performance', grain: 'campaign, weekday, and hour', keys: 'campaign.id, day, hour' }
    }),
    reportJob_({
      id: 'campaign_geo', tab: 'campaign_geo', kind: 'campaign_geo',
      resource: 'geographic_view', chunked: true, campaignChunkSize: 10,
      range: 'aggregate', required: false,
      columns: customer.concat(campaign, [
        column_('geographic_view.location_type'), column_('geographic_view.country_criterion_id', 'id'),
        sourceColumn_('geo_country_name', 'text', '', {
          sourceFields: ['geographic_view.country_criterion_id', 'geo_target_constant.name'],
          derivation: 'Current geo_target_constant.name looked up by country criterion ID'
        }),
        sourceColumn_('geo_country_code', 'text', '', {
          sourceFields: ['geographic_view.country_criterion_id', 'geo_target_constant.country_code'],
          derivation: 'Current geo_target_constant.country_code looked up by country criterion ID'
        }),
        sourceColumn_(
          'geo_target_most_specific_location_name',
          'text',
          'segments.geo_target_most_specific_location',
          {
            sourceFields: [
              'segments.geo_target_most_specific_location',
              'geo_target_constant.name'
            ],
            derivation: 'Current geo_target_constant.name looked up by most-specific criterion ID'
          }
        ),
        column_(
          'segments.geo_target_most_specific_location',
          'id',
          'geo_target_most_specific_location_criterion_id'
        ),
        sourceColumn_('geo_target_most_specific_canonical_name', 'text', '', {
          sourceFields: [
            'segments.geo_target_most_specific_location',
            'geo_target_constant.canonical_name'
          ],
          derivation: 'Current canonical location name looked up by most-specific criterion ID'
        }),
        sourceColumn_('geo_target_most_specific_target_type', 'text', '', {
          sourceFields: [
            'segments.geo_target_most_specific_location',
            'geo_target_constant.target_type'
          ],
          derivation: 'Current geo target type looked up by most-specific criterion ID'
        }),
        sourceColumn_('geo_target_state_name', 'text', 'segments.geo_target_state', {
          sourceFields: ['segments.geo_target_state', 'geo_target_constant.name'],
          derivation: 'Current geo_target_constant.name looked up by state criterion ID'
        }),
        column_('segments.geo_target_state', 'id', 'geo_target_state_criterion_id'),
        sourceColumn_('geo_target_state_canonical_name', 'text', '', {
          sourceFields: ['segments.geo_target_state', 'geo_target_constant.canonical_name'],
          derivation: 'Current canonical state name looked up by state criterion ID'
        })
      ], metrics), derived: derived,
      orderBy: [
        'campaign.id', 'segments.geo_target_most_specific_location', 'segments.geo_target_state'
      ],
      dictionary: {
        purpose: 'Geographic performance with stable criterion IDs and current readable labels',
        grain: 'campaign, location type, country criterion, most-specific criterion, and state criterion',
        keys: 'campaign.id, location type, country ID, most-specific ID, state ID'
      }
    }),
    reportJob_({
      id: 'ad_group_weekly', tab: 'ad_group_weekly', resource: 'ad_group', chunked: true, range: 'weekly',
      surveyStatusFields: ['campaign.status', 'ad_group.status'],
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group.target_cpa_micros', 'number'), column_('segments.week', 'date')
      ], metrics),
      derived: derived.concat(adGroupTargetDerived), orderBy: ['segments.week', 'campaign.id', 'ad_group.id'],
      dictionary: { purpose: 'Weekly ad-group trend', grain: 'ad group and week bucket within the 90-day window', keys: 'ad_group.id, segments.week' }
    }),
    reportJob_({
      id: 'pmax_asset_group_weekly', tab: 'pmax_asset_group_weekly', resource: 'asset_group',
      chunked: true, range: 'weekly', required: false,
      surveyStatusFields: ['campaign.status', 'asset_group.status'],
      columns: customer.concat(campaign, [
        column_('asset_group.id', 'id'), column_('asset_group.name'), column_('asset_group.status'),
        column_('segments.week', 'date')
      ], metrics), derived: derived,
      where: ["campaign.advertising_channel_type = 'PERFORMANCE_MAX'"],
      orderBy: ['segments.week', 'campaign.id', 'asset_group.id'],
      dictionary: { purpose: 'Weekly Performance Max asset-group trend', grain: 'asset group and week bucket within the 90-day window', keys: 'asset_group.id, segments.week' }
    }),
    reportJob_({
      id: 'ad_group', tab: 'ad_group', resource: 'ad_group', chunked: true, range: 'aggregate',
      surveyStatusFields: ['campaign.status', 'ad_group.status'],
      columns: customer.concat(campaign, adGroup, [column_('ad_group.target_cpa_micros', 'number')], metrics, [
        column_('metrics.search_impression_share', 'number'),
        column_('metrics.search_rank_lost_impression_share', 'number')
      ]), derived: derived.concat(adGroupTargetDerived), orderBy: ['campaign.id', 'ad_group.id'],
      dictionary: { purpose: 'Ad-group performance', grain: 'ad group', keys: 'ad_group.id' }
    }),
    reportJob_({
      id: 'campaign_inventory', tab: 'campaign_inventory', resource: 'campaign', chunked: true,
      columns: customer.concat(campaign, [
        column_('campaign.start_date_time', 'date'), column_('campaign.end_date_time', 'date'),
        column_('campaign.serving_status'), column_('campaign.bidding_strategy_type'),
        column_('campaign_budget.id', 'id'), column_('campaign_budget.name'),
        column_('campaign_budget.amount_micros', 'number'), column_('campaign_budget.period'),
        column_('campaign_budget.has_recommended_budget', 'boolean'),
        column_('campaign_budget.recommended_budget_amount_micros', 'number'),
        column_('campaign_budget.total_amount_micros', 'number'),
        column_('campaign.maximize_conversions.target_cpa_micros', 'number'),
        column_('campaign.target_cpa.target_cpa_micros', 'number')
      ]), derived: campaignBudgetDerived, orderBy: ['campaign.id'],
      dictionary: { purpose: 'Campaign configuration for currently non-inactive campaigns and historical campaigns with surveyed activity', grain: 'campaign', keys: 'campaign.id' }
    }),
    reportJob_({
      id: 'ad_group_inventory', tab: 'ad_group_inventory', resource: 'ad_group', chunked: true,
      activityLookup: adGroupEligibilityLookup_(),
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group.cpc_bid_micros', 'number'), column_('ad_group.cpm_bid_micros', 'number'),
        column_('ad_group.target_cpa_micros', 'number')
      ]), derived: adGroupInventoryDerived, orderBy: ['campaign.id', 'ad_group.id'],
      dictionary: { purpose: 'Ad-group configuration for enabled entities and inactive entities with surveyed activity', grain: 'ad group', keys: 'ad_group.id' }
    }),
    reportJob_({
      id: 'keyword_inventory', tab: 'keyword_inventory', resource: 'ad_group_criterion', chunked: true,
      activityLookup: {
        resource: 'keyword_view',
        keyFields: ['campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id'],
        statusFields: ['campaign.status', 'ad_group.status', 'ad_group_criterion.status']
      },
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_criterion.criterion_id', 'id'), column_('ad_group_criterion.status'),
        column_('ad_group_criterion.negative', 'boolean'), column_('ad_group_criterion.keyword.text'),
        column_('ad_group_criterion.keyword.match_type'), column_('ad_group_criterion.final_urls')
      ]), where: ["ad_group_criterion.type = 'KEYWORD'"],
      orderBy: ['campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id'],
      dictionary: { purpose: 'Keyword configuration for enabled entities and inactive entities with surveyed activity', grain: 'keyword criterion', keys: 'ad_group.id, criterion_id' }
    }),
    reportJob_({
      id: 'ad_inventory', tab: 'ad_inventory', resource: 'ad_group_ad', chunked: true,
      activityLookup: {
        resource: 'ad_group_ad',
        keyFields: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
        statusFields: ['campaign.status', 'ad_group.status', 'ad_group_ad.status']
      },
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_ad.ad.id', 'id'), column_('ad_group_ad.status'),
        column_('ad_group_ad.ad.name'), column_('ad_group_ad.ad.type'),
        column_('ad_group_ad.ad.final_urls'), column_('ad_group_ad.ad.final_mobile_urls')
      ]), orderBy: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
      dictionary: { purpose: 'Ad configuration for enabled entities and inactive entities with surveyed activity', grain: 'ad', keys: 'ad id' }
    }),
    reportJob_({
      id: 'ad_to_lp_map', tab: 'ad_to_lp_map', kind: 'ad_to_lp_map', chunked: true,
      columns: [
        column_('customer.id', 'id'), column_('customer.descriptive_name'),
        column_('customer.currency_code'), column_('customer.time_zone'),
        column_('campaign.id', 'id'), column_('campaign.name'),
        column_('ad_group.id', 'id'), column_('ad_group.name'),
        column_('ad_id', 'id', '', { sourceFields: ['ad_group_ad.ad.id'] }),
        derivedOutputColumn_(
          'final_url_raw', 'text',
          ['ad_group_ad.ad.final_urls', 'ad_group_ad.ad.final_mobile_urls'],
          'One URL emitted per parsed final or final-mobile URL value'
        ),
        derivedOutputColumn_(
          'final_url_norm', 'text', ['final_url_raw'],
          'Normalized scheme, host, path, and non-tracking query parameters'
        ),
        derivedOutputColumn_(
          'domain', 'text', ['final_url_norm'],
          'Lowercase host parsed from final_url_norm'
        ),
        derivedOutputColumn_(
          'expanded_url_raw', 'text', [],
          'Reserved field; not populated by v1.0.0'
        ),
        derivedOutputColumn_(
          'expanded_url_norm', 'text', ['expanded_url_raw'],
          'Reserved field; not populated by v1.0.0'
        ),
        derivedOutputColumn_(
          'url_source', 'text',
          ['ad_group_ad.ad.final_urls', 'ad_group_ad.ad.final_mobile_urls'],
          'FINAL or MOBILE according to the Google Ads source field'
        ),
        column_('tracking_url_template', 'text', '', {
          sourceFields: ['ad_group_ad.ad.tracking_url_template']
        })
      ],
      dictionary: { purpose: 'Explicit normalized ad-to-landing-page mapping', grain: 'ad and final URL', keys: 'ad_id, url_source, final_url_raw' }
    }),
    reportJob_({
      id: 'geo_targets', tab: 'geo_targets', resource: 'campaign_criterion', chunked: true, required: false,
      columns: customer.concat(campaign, [
        column_('campaign_criterion.criterion_id', 'id'), column_('campaign_criterion.status'),
        column_('campaign_criterion.negative', 'boolean'),
        column_('campaign_criterion.location.geo_target_constant')
      ]), where: ["campaign_criterion.type = 'LOCATION'"],
      orderBy: ['campaign.id', 'campaign_criterion.criterion_id'],
      dictionary: { purpose: 'Positive and excluded campaign location targets', grain: 'campaign location criterion', keys: 'campaign.id, criterion_id' }
    }),
    reportJob_({
      id: 'geo_proximity_targets', tab: 'geo_proximity_targets', resource: 'campaign_criterion',
      chunked: true, required: false,
      columns: customer.concat(campaign, [
        column_('campaign_criterion.criterion_id', 'id'), column_('campaign_criterion.status'),
        column_('campaign_criterion.negative', 'boolean'),
        column_('campaign_criterion.proximity.radius', 'number'),
        column_('campaign_criterion.proximity.radius_units'),
        column_('campaign_criterion.proximity.geo_point.latitude_in_micro_degrees', 'number'),
        column_('campaign_criterion.proximity.geo_point.longitude_in_micro_degrees', 'number'),
        column_('campaign_criterion.proximity.address.street_address'),
        column_('campaign_criterion.proximity.address.city_name'),
        column_('campaign_criterion.proximity.address.province_name'),
        column_('campaign_criterion.proximity.address.province_code'),
        column_('campaign_criterion.proximity.address.postal_code'),
        column_('campaign_criterion.proximity.address.country_code')
      ]), where: ["campaign_criterion.type = 'PROXIMITY'"],
      orderBy: ['campaign.id', 'campaign_criterion.criterion_id'],
      dictionary: { purpose: 'Campaign radius and proximity targets', grain: 'campaign proximity criterion', keys: 'campaign.id, criterion_id' }
    }),
    reportJob_({
      id: 'neg_keywords_campaign', tab: 'neg_keywords_campaign', resource: 'campaign_criterion', chunked: true,
      columns: customer.concat(campaign, [
        column_('campaign_criterion.criterion_id', 'id'), column_('campaign_criterion.status'),
        column_('campaign_criterion.keyword.text'), column_('campaign_criterion.keyword.match_type')
      ]), where: [
        "campaign_criterion.type = 'KEYWORD'", 'campaign_criterion.negative = TRUE',
        "campaign_criterion.status NOT IN ('PAUSED', 'REMOVED')"
      ],
      orderBy: ['campaign.id', 'campaign_criterion.criterion_id'],
      dictionary: { purpose: 'Direct campaign negative keywords', grain: 'campaign negative keyword', keys: 'campaign.id, criterion_id' }
    }),
    reportJob_({
      id: 'neg_keywords_ad_group', tab: 'neg_keywords_ad_group', resource: 'ad_group_criterion', chunked: true,
      activityLookup: adGroupEligibilityLookup_(),
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_criterion.criterion_id', 'id'), column_('ad_group_criterion.status'),
        column_('ad_group_criterion.keyword.text'), column_('ad_group_criterion.keyword.match_type')
      ]), where: [
        "ad_group_criterion.type = 'KEYWORD'", 'ad_group_criterion.negative = TRUE',
        "ad_group_criterion.status NOT IN ('PAUSED', 'REMOVED')"
      ],
      orderBy: ['campaign.id', 'ad_group.id', 'ad_group_criterion.criterion_id'],
      dictionary: { purpose: 'Direct ad-group negative keywords', grain: 'ad-group negative keyword', keys: 'ad_group.id, criterion_id' }
    }),
    reportJob_({
      id: 'neg_keywords_shared', tab: 'neg_keywords_shared', resource: 'shared_criterion',
      chunked: false, required: false,
      columns: customer.concat([
        column_('shared_set.id', 'id'), column_('shared_set.name'), column_('shared_set.type'),
        column_('shared_set.status'), column_('shared_criterion.criterion_id', 'id'),
        column_('shared_criterion.keyword.text'), column_('shared_criterion.keyword.match_type')
      ]),
      where: [
        "shared_set.type IN ('NEGATIVE_KEYWORDS', 'ACCOUNT_LEVEL_NEGATIVE_KEYWORDS')",
        "shared_set.status != 'REMOVED'"
      ],
      orderBy: ['shared_set.id', 'shared_criterion.criterion_id'],
      dictionary: { purpose: 'Shared and account-level negative-list criteria', grain: 'shared-set keyword criterion', keys: 'shared_set.id, criterion_id' }
    }),
    reportJob_({
      id: 'neg_keyword_shared_links', tab: 'neg_keyword_shared_links', resource: 'campaign_shared_set',
      chunked: true, required: false,
      columns: customer.concat(campaign, [
        column_('campaign_shared_set.resource_name'), column_('campaign_shared_set.status'),
        column_('shared_set.id', 'id'), column_('shared_set.name'), column_('shared_set.type'),
        column_('shared_set.status')
      ]), where: [
        "shared_set.type = 'NEGATIVE_KEYWORDS'",
        "campaign_shared_set.status != 'REMOVED'",
        "shared_set.status != 'REMOVED'"
      ],
      orderBy: ['campaign.id', 'shared_set.id'],
      dictionary: { purpose: 'Campaign links to shared negative keyword lists', grain: 'campaign and shared set', keys: 'campaign.id, shared_set.id' }
    }),
    reportJob_({
      id: 'neg_keyword_account_links', tab: 'neg_keyword_account_links',
      resource: 'customer_negative_criterion', chunked: false, required: false,
      columns: customer.concat([
        column_('customer_negative_criterion.resource_name'),
        column_('customer_negative_criterion.type'),
        column_('customer_negative_criterion.negative_keyword_list.shared_set')
      ]), where: ["customer_negative_criterion.type = 'NEGATIVE_KEYWORD_LIST'"],
      dictionary: { purpose: 'Account links to account-level negative keyword lists', grain: 'customer negative-list link', keys: 'resource_name' }
    }),
    reportJob_({
      id: 'negative_keywords_all', tab: 'negative_keywords_all', kind: 'negative_union',
      chunked: false, columns: [
        derivedOutputColumn_('source', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('scope', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('list_name', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('campaign.id', 'id', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('campaign.name', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('ad_group.id', 'id', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('ad_group.name', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('criterion_id', 'id', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('keyword.text', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('keyword.match_type', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('status', 'text', negativeUnionSources, negativeUnionDerivation),
        derivedOutputColumn_('shared_set.id', 'id', negativeUnionSources, negativeUnionDerivation)
      ],
      dictionary: { purpose: 'Unified direct, shared, and account-level negatives', grain: 'applicable negative keyword', keys: 'source, scope, entity IDs, keyword' }
    }),
    reportJob_({
      id: 'conversion_actions', tab: 'conversion_actions', resource: 'customer',
      chunked: false, range: 'aggregate',
      columns: customer.concat([
        column_('segments.conversion_action'), column_('segments.conversion_action_name'),
        column_('segments.conversion_action_category'), column_('metrics.conversions', 'number'),
        column_('metrics.conversions_value', 'number'), column_('metrics.all_conversions', 'number'),
        column_('metrics.all_conversions_value', 'number')
      ]), orderBy: ['segments.conversion_action_name'],
      dictionary: { purpose: 'Conversion totals by action', grain: 'conversion action', keys: 'segments.conversion_action' }
    }),
    reportJob_({
      id: 'conversion_action_config', tab: 'conversion_action_config', resource: 'conversion_action',
      chunked: false,
      columns: customer.concat([
        column_('conversion_action.id', 'id'), column_('conversion_action.name'),
        column_('conversion_action.status'), column_('conversion_action.type'),
        column_('conversion_action.category'), column_('conversion_action.origin'),
        column_('conversion_action.primary_for_goal', 'boolean'),
        column_('conversion_action.include_in_conversions_metric', 'boolean'),
        column_('conversion_action.counting_type'), column_('conversion_action.value_settings.default_value', 'number'),
        column_('conversion_action.value_settings.default_currency_code'),
        column_('conversion_action.value_settings.always_use_default_value', 'boolean')
      ]), orderBy: ['conversion_action.id'],
      dictionary: { purpose: 'Conversion action configuration and goal inclusion', grain: 'conversion action', keys: 'conversion_action.id' }
    }),
    reportJob_({
      id: 'quality_score_keywords', tab: 'quality_score_keywords', kind: 'quality_score',
      chunked: true, range: 'aggregate', required: false,
      columns: customer.concat(campaign, adGroup, [
        column_('ad_group_criterion.criterion_id', 'id'), column_('ad_group_criterion.keyword.text'),
        column_('ad_group_criterion.keyword.match_type'), column_('ad_group_criterion.status'),
        column_('ad_group_criterion.quality_info.quality_score', 'number'),
        column_('ad_group_criterion.quality_info.creative_quality_score'),
        column_('ad_group_criterion.quality_info.post_click_quality_score'),
        column_('ad_group_criterion.quality_info.search_predicted_ctr')
      ], metrics), derived: derived,
      where: ['Positive keyword criteria with aggregate-window impressions greater than zero'],
      dictionary: { purpose: 'Keyword Quality Score components joined to performance', grain: 'keyword criterion', keys: 'campaign.id, ad_group.id, criterion_id' }
    }),
    reportJob_({
      id: 'pmax_asset_groups', tab: 'pmax_asset_groups', kind: 'entity_performance_inventory', resource: 'asset_group',
      chunked: true, range: 'aggregate', required: false,
      eligibility: assetGroupEligibilityLookup_(),
      columns: customer.concat(campaign, [
        column_('asset_group.id', 'id'), column_('asset_group.name'), column_('asset_group.status'),
        column_('asset_group.ad_strength'), column_('asset_group.final_urls'),
        column_('asset_group.final_mobile_urls')
      ], metrics), derived: derived,
      where: ["campaign.advertising_channel_type = 'PERFORMANCE_MAX'"],
      orderBy: ['campaign.id', 'asset_group.id'],
      dictionary: { purpose: 'Performance Max asset-group performance', grain: 'asset group', keys: 'asset_group.id' }
    }),
    reportJob_({
      id: 'pmax_assets', tab: 'pmax_assets', resource: 'asset_group_asset',
      chunked: true, range: 'aggregate', required: false,
      surveyStatusFields: ['campaign.status', 'asset_group.status', 'asset_group_asset.status'],
      columns: customer.concat(campaign, [
        column_('asset_group.id', 'id'), column_('asset_group.name'), column_('asset_group.status'),
        column_('asset_group_asset.field_type'), column_('asset_group_asset.primary_status'),
        column_('asset_group_asset.primary_status_reasons'),
        column_('asset_group_asset.status'), column_('asset_group_asset.source'),
        column_('asset_group_asset.policy_summary.approval_status'),
        column_('asset_group_asset.policy_summary.review_status'), column_('asset.id', 'id'),
        column_('asset.name'), column_('asset.type'), column_('asset.text_asset.text')
      ], metrics), derived: derived,
      where: ["campaign.advertising_channel_type = 'PERFORMANCE_MAX'"],
      orderBy: ['campaign.id', 'asset_group.id', 'asset.id'],
      dictionary: { purpose: 'Performance Max asset inventory, serving status, policy status, and metrics', grain: 'asset-group asset association', keys: 'asset_group.id, asset.id, field_type' }
    }),
    reportJob_({
      id: 'pmax_audience_signals', tab: 'pmax_audience_signals', resource: 'asset_group_signal',
      chunked: true, required: false,
      eligibleSource: {
        tab: 'pmax_asset_groups',
        keyFields: ['campaign.id', 'asset_group.id'],
        enrichFields: ['asset_group.status']
      },
      columns: customer.concat(campaign, [
        column_('asset_group.id', 'id'), column_('asset_group.name'), sourceColumn_(
          'asset_group.status', 'text', '', {
            sourceFields: ['pmax_asset_groups.asset_group.status'],
            derivation: 'Joined from the eligible pmax_asset_groups parent row'
          }
        ),
        column_('asset_group_signal.resource_name'),
        column_('asset_group_signal.audience.audience'),
        column_('asset_group_signal.search_theme.text'),
        column_('asset_group_signal.approval_status'),
        column_('asset_group_signal.disapproval_reasons')
      ]), where: ["campaign.advertising_channel_type = 'PERFORMANCE_MAX'"],
      orderBy: ['campaign.id', 'asset_group.id'],
      dictionary: { purpose: 'Performance Max audience and search-theme signals; signals are hints, not targeting restrictions', grain: 'asset-group signal', keys: 'asset_group_signal.resource_name' }
    }),
    reportJob_({
      id: 'user_list_performance', tab: 'user_list_performance', kind: 'audience_performance',
      chunked: true, required: false, range: 'aggregate', scopes: ['CAMPAIGN', 'AD_GROUP'],
      columns: [
        derivedOutputColumn_(
          'scope', 'text', ['campaign_audience_view', 'ad_group_audience_view'],
          'Exporter-assigned CAMPAIGN or AD_GROUP label identifying the source audience view'
        ),
        derivedOutputColumn_(
          'user_list_resource', 'text', [
            'campaign_criterion.user_list.user_list',
            'ad_group_criterion.user_list.user_list'
          ],
          'Scope-aware alias of the exact Google user-list resource name'
        ),
        derivedOutputColumn_(
          'criterion_id', 'id', [
            'campaign_criterion.criterion_id',
            'ad_group_criterion.criterion_id'
          ],
          'Scope-aware alias of the campaign or ad-group audience criterion ID'
        ),
        column_('campaign.id', 'id'), column_('campaign.name'),
        column_('ad_group.id', 'id', '', {
          blankWhen: 'CAMPAIGN scope; populated only for AD_GROUP rows'
        }),
        column_('ad_group.name', 'text', '', {
          blankWhen: 'CAMPAIGN scope; populated only for AD_GROUP rows'
        })
      ].concat(metrics), derived: derived,
      dictionary: {
        purpose: '90-day campaign- and ad-group-level user-list performance keyed by the exact Google resource name',
        grain: 'scope, user list, and applicable entity',
        keys: 'scope, resource, entity IDs'
      }
    }),
    reportJob_({
      id: 'asset_extensions', tab: 'asset_extensions', kind: 'asset_extensions',
      chunked: false, required: false, range: 'aggregate',
      columns: [derivedOutputColumn_(
        'scope', 'text', ['customer_asset', 'campaign_asset', 'ad_group_asset'],
        'Exporter-assigned CUSTOMER, CAMPAIGN, or AD_GROUP label identifying the association resource'
      )].concat(customer, [
        column_('campaign.id', 'id'), column_('campaign.name'),
        column_('ad_group.id', 'id'), column_('ad_group.name'),
        derivedOutputColumn_(
          'field_type', 'text', [
            'customer_asset.field_type', 'campaign_asset.field_type', 'ad_group_asset.field_type'
          ],
          'Scope-aware alias of the association field type'
        ),
        derivedOutputColumn_(
          'association_status', 'text', [
            'customer_asset.status', 'campaign_asset.status', 'ad_group_asset.status'
          ],
          'Scope-aware alias of the association status'
        ),
        derivedOutputColumn_(
          'source', 'text', [
            'customer_asset.source', 'campaign_asset.source', 'ad_group_asset.source'
          ],
          'Scope-aware alias of the association source'
        ),
        column_('asset.id', 'id'), column_('asset.name'), column_('asset.type'),
        derivedOutputColumn_(
          'asset.text', 'text', [
            'asset.text_asset.text', 'asset.callout_asset.callout_text',
            'asset.sitelink_asset.link_text', 'asset.structured_snippet_asset.header',
            'asset.structured_snippet_asset.values', 'asset.promotion_asset.promotion_target',
            'asset.call_asset.phone_number'
          ],
          'First available readable value from the supported asset subtype fields'
        ),
        column_('asset.sitelink_asset.description1'), column_('asset.sitelink_asset.description2'),
        column_('asset.structured_snippet_asset.header'),
        column_('asset.structured_snippet_asset.values'),
        column_('asset.promotion_asset.promotion_target'),
        column_('asset.call_asset.phone_number')
      ], metrics), derived: derived,
      where: ['Associations with nonzero activity in the frozen aggregate survey'],
      dictionary: { purpose: 'Account, campaign, and ad-group asset association performance', grain: 'entity and asset association', keys: 'scope, entity IDs, asset.id, field_type' }
    }),
    reportJob_({
      id: 'change_history', tab: 'change_history', kind: 'change_history',
      chunked: false, required: false, range: 'change',
      columns: changeHistoryColumns_(CONFIG.INCLUDE_SENSITIVE_CHANGE_DETAILS),
      dictionary: { purpose: 'Account change events with privacy-safe fields by default', grain: 'change event', keys: 'change_event.resource_name' }
    }),
    reportJob_({
      id: 'data_dictionary', tab: DICTIONARY_SHEET_NAME, kind: 'data_dictionary',
      chunked: false, required: true,
      columns: [
        derivedOutputColumn_('tab', 'text', ['export manifest'], 'Final output tab name'),
        derivedOutputColumn_('purpose', 'text', ['export manifest'], 'Tab purpose declared by the exporter'),
        derivedOutputColumn_('row_grain', 'text', ['export manifest'], 'Row grain declared by the exporter'),
        derivedOutputColumn_('date_range', 'text', ['export manifest'], 'Frozen range contract for the tab'),
        derivedOutputColumn_('material_filters', 'text', ['export manifest'], 'Material inclusion filters summarized from the job contract'),
        derivedOutputColumn_('sensitive_data', 'text', ['export manifest'], 'Sensitivity classification assigned by the exporter'),
        derivedOutputColumn_('google_side_limitations', 'text', ['export manifest'], 'Google-side and interpretation limitations assigned by the exporter'),
        derivedOutputColumn_('keys', 'text', ['export manifest'], 'Declared composite key fields'),
        derivedOutputColumn_('field_dictionary_reference', 'text', ['export manifest'], 'Field-level metadata tab for rectangular outputs')
      ],
      dictionary: { purpose: 'Definitions and limitations for every output tab', grain: 'output tab', keys: 'tab' }
    }),
    reportJob_({
      id: 'field_dictionary', tab: FIELD_DICTIONARY_SHEET_NAME, kind: 'field_dictionary',
      chunked: false, required: true,
      columns: [
        derivedOutputColumn_('tab', 'text', ['export manifest'], 'Final rectangular output tab name'),
        derivedOutputColumn_('column_ordinal', 'number', ['export manifest'], '1-based position in the output tab'),
        derivedOutputColumn_('field', 'text', ['export manifest'], 'Exact output header'),
        derivedOutputColumn_('source_fields', 'text', ['export manifest'], 'Google or exporter fields used to populate the output field'),
        derivedOutputColumn_('data_type', 'text', ['export manifest'], 'Normalized output data type'),
        derivedOutputColumn_('unit', 'text', ['export manifest'], 'Display or analytical unit'),
        derivedOutputColumn_('is_key', 'boolean', ['export manifest'], 'TRUE when the field participates in the declared tab key'),
        derivedOutputColumn_('is_derived', 'boolean', ['export manifest'], 'TRUE when the exporter joins, computes, normalizes, or composes the field'),
        derivedOutputColumn_('derivation', 'text', ['export manifest'], 'Exporter derivation or join rule when applicable'),
        derivedOutputColumn_('blank_when', 'text', ['export manifest'], 'Known general blankability condition')
      ],
      dictionary: {
        purpose: 'Normalized field-level definitions for every rectangular output tab',
        grain: 'output tab and field', keys: 'tab, field'
      }
    })
  ];
  var keyFields = dictionaryKeyFields_();
  manifest.forEach(function(job) {
    job.dictionary.keyFields = (keyFields[job.tab] || []).slice();
    job.dictionary.keys = job.dictionary.keyFields.join(', ');
  });
  return manifest;
}

function buildGaqlQuery_(job, ranges, campaignIds) {
  if (!job || (job.kind !== 'gaql' && job.kind !== 'campaign_geo')) {
    throw new Error('Job is not a declarative GAQL job.');
  }
  var fields = job.columns.filter(function(column) {
    return column.select !== false;
  }).map(function(column) { return column.field; });
  var clauses = (job.where || []).slice();
  if (job.range) {
    var range = ranges[job.range];
    if (!range) throw new Error('Missing frozen date range: ' + job.range);
    clauses.push("segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'");
  }
  if (job.chunked) {
    if (!campaignIds || !campaignIds.length) throw new Error('Campaign chunk is empty for ' + job.id + '.');
    var safeIds = campaignIds.map(function(id) {
      var text = String(id);
      if (!/^\d+$/.test(text)) throw new Error('Unsafe campaign ID: ' + text);
      return text;
    });
    clauses.push('campaign.id IN (' + safeIds.join(', ') + ')');
  }
  var query = 'SELECT\n  ' + fields.join(',\n  ') + '\nFROM ' + job.resource;
  if (clauses.length) query += '\nWHERE ' + clauses.join('\n  AND ');
  if (job.orderBy && job.orderBy.length) query += '\nORDER BY ' + job.orderBy.join(', ');
  return query;
}

function buildActivityLookupQuery_(job, range, campaignIds) {
  var lookup = job && job.activityLookup;
  if (!lookup || !lookup.resource || !lookup.keyFields || !lookup.keyFields.length) {
    throw new Error('Job does not define a valid activity lookup.');
  }
  if (!range || !range.start || !range.end) {
    throw new Error('Activity lookup requires the frozen aggregate date range.');
  }
  var fields = lookup.keyFields.concat(activityMetricFields_());
  var clauses = (lookup.where || []).slice();
  clauses.push("segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'");
  clauses.push('campaign.id IN (' + safeCampaignIdList_(campaignIds, job.id + ' activity') + ')');
  return 'SELECT\n  ' + fields.join(',\n  ') + '\nFROM ' + lookup.resource +
    '\nWHERE ' + clauses.join('\n  AND ') +
    '\nORDER BY ' + lookup.keyFields.join(', ');
}

function buildEntityPerformanceQueries_(job, range, campaignIds) {
  var eligibility = job && job.eligibility;
  if (!job || job.kind !== 'entity_performance_inventory' || !eligibility ||
      !eligibility.keyFields || !eligibility.keyFields.length) {
    throw new Error('Job does not define an entity performance inventory contract.');
  }
  if (!range || !range.start || !range.end) {
    throw new Error('Entity performance inventory requires the frozen aggregate date range.');
  }
  var selected = (job.columns || []).filter(function(column) { return column.select !== false; });
  var currentFields = selected.filter(function(column) {
    return column.field.indexOf('metrics.') !== 0;
  }).map(function(column) { return column.field; });
  var activityFields = selected.map(function(column) { return column.field; });
  var baseClauses = (job.where || eligibility.where || []).slice();
  var campaignClause = 'campaign.id IN (' +
    safeCampaignIdList_(campaignIds, job.id + ' entity performance') + ')';
  var currentClauses = baseClauses.concat([campaignClause]);
  var activityClauses = baseClauses.concat([
    "segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'",
    campaignClause
  ]);
  var order = (job.orderBy && job.orderBy.length) ? job.orderBy : eligibility.keyFields;
  return {
    current: 'SELECT\n  ' + currentFields.join(',\n  ') + '\nFROM ' + job.resource +
      '\nWHERE ' + currentClauses.join('\n  AND ') + '\nORDER BY ' + order.join(', '),
    activity: 'SELECT\n  ' + activityFields.join(',\n  ') + '\nFROM ' + job.resource +
      '\nWHERE ' + activityClauses.join('\n  AND ') + '\nORDER BY ' + order.join(', ')
  };
}

function buildEntityPerformanceRows_(job, currentRows, activityRows) {
  var eligibility = job && job.eligibility;
  if (!eligibility || !eligibility.keyFields || !eligibility.statusFields) {
    throw new Error('Entity performance rows require key and status fields.');
  }
  var additiveFields = [
    'metrics.impressions', 'metrics.clicks', 'metrics.interactions', 'metrics.cost_micros',
    'metrics.conversions', 'metrics.conversions_value',
    'metrics.all_conversions', 'metrics.all_conversions_value'
  ];
  var currentByKey = {};
  var activityBaseByKey = {};
  var totalsByKey = {};
  var order = [];
  (currentRows || []).forEach(function(row) {
    var key = entityKey_(row, eligibility.keyFields);
    if (!currentByKey[key]) {
      currentByKey[key] = row;
      order.push(key);
    }
  });
  (activityRows || []).forEach(function(row) {
    var key = entityKey_(row, eligibility.keyFields);
    if (!activityBaseByKey[key]) activityBaseByKey[key] = row;
    if (!totalsByKey[key]) {
      totalsByKey[key] = {};
      additiveFields.forEach(function(field) { totalsByKey[key][field] = 0; });
    }
    additiveFields.forEach(function(field) {
      totalsByKey[key][field] += metricNumber_(row, field);
    });
    if (order.indexOf(key) < 0) order.push(key);
  });

  var output = [];
  order.forEach(function(key) {
    var current = currentByKey[key];
    var activityBase = activityBaseByKey[key];
    var totals = totalsByKey[key] || {};
    var joined = {};
    var base = current || activityBase || {};
    Object.keys(base).forEach(function(field) { joined[field] = base[field]; });
    if (activityBase) {
      Object.keys(activityBase).forEach(function(field) {
        if (joined[field] === '' || joined[field] === null || joined[field] === undefined) {
          joined[field] = activityBase[field];
        }
      });
    }
    additiveFields.forEach(function(field) {
      joined[field] = totals[field] || 0;
    });
    joined['metrics.ctr'] = joined['metrics.impressions'] ?
      joined['metrics.clicks'] / joined['metrics.impressions'] : 0;
    joined['metrics.average_cpc'] = joined['metrics.clicks'] ?
      joined['metrics.cost_micros'] / joined['metrics.clicks'] : '';
    if (!shouldIncludeSurveyEntity_(joined, eligibility.statusFields)) return;
    output.push(mapGaqlRow_(job, joined));
  });
  return output;
}

function metricNumber_(row, field) {
  var value = row[field];
  if (value === '' || value === null || value === undefined) return 0;
  var number = Number(String(value).replace(/,/g, ''));
  return isFinite(number) ? number : 0;
}

function coerceReportValue_(column, value) {
  if (value === null || value === undefined || value === '') return '';
  if (column.type === 'number') {
    var number = Number(String(value).replace(/,/g, ''));
    return isFinite(number) ? number : value;
  }
  if (column.type === 'boolean') {
    if (value === true || String(value).toLowerCase() === 'true') return true;
    if (value === false || String(value).toLowerCase() === 'false') return false;
  }
  return String(value);
}

function headersForJob_(job) {
  return job.columns.map(function(column) { return column.header; }).concat(
    (job.derived || []).map(function(column) { return column.header; })
  );
}

function textColumnIndexes_(job) {
  var indexes = [];
  (job.columns || []).forEach(function(column, index) {
    if (column.type !== 'number' && column.type !== 'boolean') indexes.push(index + 1);
  });
  return indexes;
}

function mapGaqlRow_(job, row) {
  var values = job.columns.map(function(column) {
    return coerceReportValue_(column, row[column.field]);
  });
  (job.derived || []).forEach(function(derived) { values.push(derived.compute(row)); });
  return values;
}

function runGaqlChunk_(job, ranges, campaignIds, sheet, runtime, writerOptions) {
  runtime = runtime || {
    report: function(query) { return reportRuntime_(query); },
    sleep: function(ms) { Utilities.sleep(ms); }
  };
  writerOptions = writerOptions || {};
  if (!writerOptions.sleep && runtime.sleep) writerOptions.sleep = runtime.sleep;
  var eligibleSourceIndex = null;
  if (job.eligibleSource) {
    if (!Object.prototype.hasOwnProperty.call(writerOptions, 'eligibleSourceRows')) {
      throw new Error('Missing eligible source rows for ' + job.id + '.');
    }
    eligibleSourceIndex = indexRowsByEntityKey_(
      writerOptions.eligibleSourceRows,
      job.eligibleSource.keyFields
    );
  }
  var activeKeys = null;
  if (job.activityLookup) {
    activeKeys = {};
    var activityQuery = buildActivityLookupQuery_(job, ranges.aggregate, campaignIds);
    var activityIterator = runtime.report(activityQuery).rows();
    while (activityIterator.hasNext()) {
      var activityRow = activityIterator.next();
      if (hasSurveyActivity_(activityRow)) {
        activeKeys[entityKey_(activityRow, job.activityLookup.keyFields)] = true;
      }
    }
  }
  var query = buildGaqlQuery_(job, ranges, campaignIds);
  var iterator = runtime.report(query).rows();
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), writerOptions);
  while (iterator.hasNext()) {
    var row = iterator.next();
    if (eligibleSourceIndex) {
      var sourceRow = eligibleSourceIndex[entityKey_(row, job.eligibleSource.keyFields)];
      if (!sourceRow) continue;
      var enriched = {};
      Object.keys(row).forEach(function(field) { enriched[field] = row[field]; });
      (job.eligibleSource.enrichFields || []).forEach(function(field) {
        enriched[field] = sourceRow[field] === undefined || sourceRow[field] === null ? '' : sourceRow[field];
      });
      row = enriched;
    }
    var statusFields = job.activityLookup ? job.activityLookup.statusFields : job.surveyStatusFields;
    var keyFields = job.activityLookup ? job.activityLookup.keyFields : null;
    if (!statusFields || shouldIncludeSurveyEntity_(row, statusFields, activeKeys, keyFields)) {
      writer.push(mapGaqlRow_(job, row));
    }
  }
  writer.flush();
  if (runtime.sleep && Number(CONFIG.THROTTLE_MS) > 0) runtime.sleep(Number(CONFIG.THROTTLE_MS));
  return writer.count();
}

function normalizeGeoCriterionId_(value) {
  if (value === null || value === undefined || value === '') return '';
  var text = String(value).trim();
  var match = /^(?:geoTargetConstants\/)?(\d+)$/.exec(text);
  if (!match) {
    throw new Error('Google returned an unexpected geographic criterion identifier shape.');
  }
  return match[1];
}

function sortNumericTextIds_(ids) {
  return ids.sort(function(left, right) {
    if (left.length !== right.length) return left.length - right.length;
    return left < right ? -1 : (left > right ? 1 : 0);
  });
}

function normalizedUniqueGeoIds_(ids) {
  var seen = {};
  (ids || []).forEach(function(value) {
    var id = normalizeGeoCriterionId_(value);
    if (id) seen[id] = true;
  });
  return sortNumericTextIds_(Object.keys(seen));
}

function buildGeoTargetLookupQueries_(ids, batchSize) {
  var size = batchSize === undefined ? GEO_TARGET_LOOKUP_BATCH_SIZE : Number(batchSize);
  if (!isFinite(size) || size < 1 || size > GEO_TARGET_LOOKUP_BATCH_SIZE ||
      Math.floor(size) !== size) {
    throw new Error('Geographic label lookup batch size must be an integer from 1 through ' +
      GEO_TARGET_LOOKUP_BATCH_SIZE + '.');
  }
  var uniqueIds = normalizedUniqueGeoIds_(ids);
  var fields = [
    'geo_target_constant.id',
    'geo_target_constant.resource_name',
    'geo_target_constant.name',
    'geo_target_constant.canonical_name',
    'geo_target_constant.country_code',
    'geo_target_constant.target_type',
    'geo_target_constant.status'
  ];
  var queries = [];
  for (var offset = 0; offset < uniqueIds.length; offset += size) {
    var batch = uniqueIds.slice(offset, offset + size);
    queries.push(
      'SELECT\n  ' + fields.join(',\n  ') +
      '\nFROM geo_target_constant\nWHERE geo_target_constant.id IN (' + batch.join(', ') + ')' +
      '\nORDER BY geo_target_constant.id'
    );
  }
  return queries;
}

function resolveGeoTargetMetadata_(ids, runtime, cache) {
  var uniqueIds = normalizedUniqueGeoIds_(ids);
  var freshIds = uniqueIds.filter(function(id) {
    return !Object.prototype.hasOwnProperty.call(cache, id);
  });
  var queries = buildGeoTargetLookupQueries_(freshIds, GEO_TARGET_LOOKUP_BATCH_SIZE);
  var failures = [];
  queries.forEach(function(query, queryIndex) {
    var batch = freshIds.slice(
      queryIndex * GEO_TARGET_LOOKUP_BATCH_SIZE,
      (queryIndex + 1) * GEO_TARGET_LOOKUP_BATCH_SIZE
    );
    try {
      var iterator = runtime.report(query).rows();
      while (iterator.hasNext()) {
        var row = iterator.next();
        var id = normalizeGeoCriterionId_(row['geo_target_constant.id']);
        if (id && batch.indexOf(id) >= 0) cache[id] = row;
      }
      batch.forEach(function(id) {
        if (!Object.prototype.hasOwnProperty.call(cache, id)) cache[id] = null;
      });
    } catch (error) {
      var failure = expectedPartialSourceLimitation_(
        'Geographic label lookup',
        error
      );
      batch.forEach(function(id) { cache[id] = null; });
      failures.push(failure);
    }
  });
  var missingIds = uniqueIds.filter(function(id) { return cache[id] === null; });
  return { missingIds: missingIds, failures: failures };
}

function geoMetadataValue_(metadata, field) {
  if (!metadata) return '';
  var value = metadata[field];
  return value === null || value === undefined ? '' : value;
}

function runCampaignGeoChunk_(job, ranges, campaignIds, sheet, runtime, writerOptions) {
  runtime = runtime || {
    report: function(query, options) { return reportRuntime_(query, options); },
    sleep: function(ms) { Utilities.sleep(ms); }
  };
  writerOptions = writerOptions || {};
  if (!writerOptions.sleep && runtime.sleep) writerOptions.sleep = runtime.sleep;
  var geoCache = writerOptions.geoTargetCache || {};
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), writerOptions);
  var iterator = runtime.report(
    buildGaqlQuery_(job, ranges, campaignIds),
    { resolveGeoNames: false }
  ).rows();
  var pending = [];
  var seenKeys = {};
  var unresolved = {};
  var lookupFailures = [];
  var keyFields = job.dictionary.keyFields;
  var pendingLimit = Math.max(1, Number(writerOptions.batchRows || CONFIG.BATCH_ROWS || 1000));

  function processPending_() {
    if (!pending.length) return;
    var ids = [];
    pending.forEach(function(row) {
      ids.push(row['geographic_view.country_criterion_id']);
      ids.push(row['segments.geo_target_most_specific_location']);
      ids.push(row['segments.geo_target_state']);
    });
    var resolution = resolveGeoTargetMetadata_(ids, runtime, geoCache);
    resolution.missingIds.forEach(function(id) { unresolved[id] = true; });
    lookupFailures = lookupFailures.concat(resolution.failures);

    pending.forEach(function(row) {
      var countryId = normalizeGeoCriterionId_(row['geographic_view.country_criterion_id']);
      var locationId = normalizeGeoCriterionId_(row['segments.geo_target_most_specific_location']);
      var stateId = normalizeGeoCriterionId_(row['segments.geo_target_state']);
      var country = countryId ? geoCache[countryId] : null;
      var location = locationId ? geoCache[locationId] : null;
      var state = stateId ? geoCache[stateId] : null;
      row['geographic_view.country_criterion_id'] = countryId;
      row['segments.geo_target_most_specific_location'] = locationId;
      row['segments.geo_target_state'] = stateId;
      row.geo_country_name = geoMetadataValue_(country, 'geo_target_constant.name');
      row.geo_country_code = geoMetadataValue_(country, 'geo_target_constant.country_code');
      row.geo_target_most_specific_location_name =
        geoMetadataValue_(location, 'geo_target_constant.name');
      row.geo_target_most_specific_canonical_name =
        geoMetadataValue_(location, 'geo_target_constant.canonical_name');
      row.geo_target_most_specific_target_type =
        geoMetadataValue_(location, 'geo_target_constant.target_type');
      row.geo_target_state_name = geoMetadataValue_(state, 'geo_target_constant.name');
      row.geo_target_state_canonical_name =
        geoMetadataValue_(state, 'geo_target_constant.canonical_name');

      var keyRow = {};
      keyRow['campaign.id'] = row['campaign.id'];
      keyRow['geographic_view.location_type'] = row['geographic_view.location_type'];
      keyRow['geographic_view.country_criterion_id'] = countryId;
      keyRow.geo_target_most_specific_location_criterion_id = locationId;
      keyRow.geo_target_state_criterion_id = stateId;
      var key = entityKey_(keyRow, keyFields);
      if (seenKeys[key]) {
        throw new Error('campaign_geo returned a duplicate canonical geographic key.');
      }
      seenKeys[key] = true;
      writer.push(mapGaqlRow_(job, row));
    });
    pending = [];
  }

  while (iterator.hasNext()) {
    pending.push(iterator.next());
    if (pending.length >= pendingLimit) processPending_();
  }
  processPending_();
  writer.flush();
  if (runtime.sleep && Number(CONFIG.THROTTLE_MS) > 0) {
    runtime.sleep(Number(CONFIG.THROTTLE_MS));
  }
  var missingCount = Object.keys(unresolved).length;
  var limitations = [];
  if (missingCount) {
    limitations.push(missingCount +
      ' geographic criterion ID(s) were preserved, but current readable metadata was unavailable');
  }
  if (lookupFailures.length) {
    limitations.push('One or more geographic label lookup requests failed');
  }
  return {
    rows: writer.count(),
    status: limitations.length ? 'LIMITED' : 'OK',
    limitation: limitations.join('; ')
  };
}

function safeCampaignIdList_(campaignIds, label) {
  if (!campaignIds || !campaignIds.length) throw new Error(label + ' campaign chunk is empty.');
  return campaignIds.map(function(id) {
    var value = String(id);
    if (!/^\d+$/.test(value)) throw new Error('Unsafe campaign ID: ' + value);
    return value;
  }).join(', ');
}

function parseUrlList_(value) {
  if (value === null || value === undefined || value === '') return [];
  if (Array.isArray(value)) {
    return value.map(function(item) { return String(item || '').trim(); }).filter(Boolean);
  }
  var text = String(value).trim();
  if (!text) return [];
  if (text.charAt(0) === '[') {
    try {
      var parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        return parsed.map(function(item) { return String(item || '').trim(); }).filter(Boolean);
      }
    } catch (ignoredJson) {}
  }
  text = text.replace(/^\[\s*/, '').replace(/\s*\]$/, '');
  var parts = text.split(/,\s*(?=(?:https?:)?\/\/)/i);
  return parts.map(function(part) {
    return part.trim().replace(/^['"]|['"]$/g, '');
  }).filter(Boolean);
}

function normalizeLandingPageUrl_(url) {
  var text = String(url || '').trim();
  if (!text) return '';
  var absolute = /^(?:https?:)?\/\/([^\/?#]+)([^?#]*)(?:\?([^#]*))?/i.exec(text);
  if (absolute) {
    var host = String(absolute[1] || '').toLowerCase();
    var path = String(absolute[2] || '').replace(/\/+$/, '');
    var tracking = {
      utm_source: true, utm_medium: true, utm_campaign: true, utm_content: true,
      utm_term: true, gclid: true, gbraid: true, wbraid: true, msclkid: true,
      fbclid: true, yclid: true
    };
    var kept = String(absolute[3] || '').split('&').filter(function(part) {
      if (!part) return false;
      var rawKey = part.split('=')[0];
      var key;
      try { key = decodeURIComponent(rawKey).toLowerCase(); }
      catch (ignoredDecode) { key = rawKey.toLowerCase(); }
      return !tracking[key];
    });
    return host ? 'https://' + host + path + (kept.length ? '?' + kept.join('&') : '') : '';
  }
  return text.split('#')[0].split('?')[0].replace(/\/+$/, '');
}

function landingPageDomain_(normalizedUrl) {
  var match = /^https?:\/\/([^\/]+)/i.exec(String(normalizedUrl || ''));
  return match ? String(match[1]).toLowerCase() : '';
}

function adToLandingPageActivityLookup_() {
  return {
    resource: 'ad_group_ad',
    keyFields: ['campaign.id', 'ad_group.id', 'ad_group_ad.ad.id'],
    statusFields: ['campaign.status', 'ad_group.status', 'ad_group_ad.status']
  };
}

function buildAdToLandingPageRows_(row, activeKeys) {
  var lookup = adToLandingPageActivityLookup_();
  if (!shouldIncludeSurveyEntity_(row, lookup.statusFields, activeKeys, lookup.keyFields)) return [];
  var output = [];
  var seen = {};
  function add_(raw, source) {
    var value = String(raw || '').trim();
    if (!value) return;
    var key = source + '|' + value;
    if (seen[key]) return;
    seen[key] = true;
    var normalized = normalizeLandingPageUrl_(value);
    output.push([
      row['customer.id'] || '', row['customer.descriptive_name'] || '',
      row['customer.currency_code'] || '', row['customer.time_zone'] || '',
      row['campaign.id'] || '', row['campaign.name'] || '',
      row['ad_group.id'] || '', row['ad_group.name'] || '',
      row['ad_group_ad.ad.id'] || '', value, normalized, landingPageDomain_(normalized),
      '', '', source, row['ad_group_ad.ad.tracking_url_template'] || ''
    ]);
  }
  parseUrlList_(row['ad_group_ad.ad.final_urls']).forEach(function(url) { add_(url, 'FINAL'); });
  parseUrlList_(row['ad_group_ad.ad.final_mobile_urls']).forEach(function(url) { add_(url, 'MOBILE'); });
  return output;
}

function buildAdToLandingPageQuery_(campaignIds) {
  return 'SELECT\n  customer.id,\n  customer.descriptive_name,\n  customer.currency_code,\n' +
    '  customer.time_zone,\n  campaign.id,\n  campaign.name,\n  campaign.status,\n' +
    '  ad_group.id,\n  ad_group.name,\n  ad_group.status,\n' +
    '  ad_group_ad.ad.id,\n  ad_group_ad.status,\n' +
    '  ad_group_ad.ad.final_urls,\n  ad_group_ad.ad.final_mobile_urls,\n' +
    '  ad_group_ad.ad.tracking_url_template\nFROM ad_group_ad\nWHERE campaign.id IN (' +
    safeCampaignIdList_(campaignIds, 'Ad-to-landing-page') +
    ')\nORDER BY campaign.id, ad_group.id, ad_group_ad.ad.id';
}

function qualityScoreKey_(row) {
  return [
    row['campaign.id'] || '', row['ad_group.id'] || '',
    row['ad_group_criterion.criterion_id'] || ''
  ].join('|');
}

function buildQualityScoreQueries_(job, range, campaignIds) {
  if (!range || !range.start || !range.end) throw new Error('Quality Score requires a date range.');
  var ids = safeCampaignIdList_(campaignIds, 'Quality Score');
  var staticFields = job.columns.filter(function(column) {
    return column.field.indexOf('metrics.') !== 0;
  }).map(function(column) { return column.field; });
  var metricFields = performanceColumns_().map(function(column) { return column.field; });
  return {
    staticQuery: 'SELECT\n  ' + staticFields.join(',\n  ') +
      '\nFROM ad_group_criterion\nWHERE ad_group_criterion.type = \'KEYWORD\'' +
      '\n  AND ad_group_criterion.negative = FALSE' +
      '\n  AND campaign.id IN (' + ids + ')' +
      '\nORDER BY campaign.id, ad_group.id, ad_group_criterion.criterion_id',
    metricsQuery: 'SELECT\n  campaign.id,\n  ad_group.id,\n  ad_group_criterion.criterion_id,\n  ' +
      metricFields.join(',\n  ') +
      "\nFROM keyword_view\nWHERE segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'" +
      '\n  AND metrics.impressions > 0' +
      '\n  AND campaign.id IN (' + ids + ')' +
      '\nORDER BY campaign.id, ad_group.id, ad_group_criterion.criterion_id'
  };
}

function buildQualityScoreRows_(job, staticRows, metricRows) {
  var staticByKey = {};
  var order = [];
  (staticRows || []).forEach(function(row) {
    var key = qualityScoreKey_(row);
    if (!staticByKey[key]) {
      staticByKey[key] = row;
      order.push(key);
    }
  });
  var totals = {};
  (metricRows || []).forEach(function(row) {
    var key = qualityScoreKey_(row);
    if (!staticByKey[key]) return;
    var total = totals[key] || {
      impressions: 0, clicks: 0, interactions: 0, costMicros: 0, conversions: 0,
      conversionsValue: 0, allConversions: 0, allConversionsValue: 0
    };
    total.impressions += metricNumber_(row, 'metrics.impressions');
    total.clicks += metricNumber_(row, 'metrics.clicks');
    total.interactions += metricNumber_(row, 'metrics.interactions');
    total.costMicros += metricNumber_(row, 'metrics.cost_micros');
    total.conversions += metricNumber_(row, 'metrics.conversions');
    total.conversionsValue += metricNumber_(row, 'metrics.conversions_value');
    total.allConversions += metricNumber_(row, 'metrics.all_conversions');
    total.allConversionsValue += metricNumber_(row, 'metrics.all_conversions_value');
    totals[key] = total;
  });
  var output = [];
  order.forEach(function(key) {
    var total = totals[key];
    if (!total || total.impressions <= 0) return;
    var joined = {};
    Object.keys(staticByKey[key]).forEach(function(field) { joined[field] = staticByKey[key][field]; });
    joined['metrics.impressions'] = total.impressions;
    joined['metrics.clicks'] = total.clicks;
    joined['metrics.interactions'] = total.interactions;
    joined['metrics.ctr'] = total.impressions ? total.clicks / total.impressions : 0;
    joined['metrics.average_cpc'] = total.clicks ? total.costMicros / total.clicks : '';
    joined['metrics.cost_micros'] = total.costMicros;
    joined['metrics.conversions'] = total.conversions;
    joined['metrics.conversions_value'] = total.conversionsValue;
    joined['metrics.all_conversions'] = total.allConversions;
    joined['metrics.all_conversions_value'] = total.allConversionsValue;
    output.push(mapGaqlRow_(job, joined));
  });
  return output;
}

function buildNegativeUnionRows_(tables) {
  tables = tables || {};
  var output = [];
  var seen = {};
  function add_(row) {
    var key = stableStringify_(row);
    if (!seen[key]) { seen[key] = true; output.push(row); }
  }

  (tables.neg_keywords_campaign || []).forEach(function(row) {
    add_([
      'DIRECT_CAMPAIGN', 'CAMPAIGN', '', row['campaign.id'] || '', row['campaign.name'] || '',
      '', '', row['campaign_criterion.criterion_id'] || '',
      row['campaign_criterion.keyword.text'] || '', row['campaign_criterion.keyword.match_type'] || '',
      row['campaign_criterion.status'] || '', ''
    ]);
  });
  (tables.neg_keywords_ad_group || []).forEach(function(row) {
    add_([
      'DIRECT_AD_GROUP', 'AD_GROUP', '', row['campaign.id'] || '', row['campaign.name'] || '',
      row['ad_group.id'] || '', row['ad_group.name'] || '',
      row['ad_group_criterion.criterion_id'] || '', row['ad_group_criterion.keyword.text'] || '',
      row['ad_group_criterion.keyword.match_type'] || '', row['ad_group_criterion.status'] || '', ''
    ]);
  });

  var campaignLinks = {};
  (tables.neg_keyword_shared_links || []).forEach(function(row) {
    var id = String(row['shared_set.id'] || '');
    campaignLinks[id] = campaignLinks[id] || [];
    campaignLinks[id].push(row);
  });
  var accountLinks = {};
  (tables.neg_keyword_account_links || []).forEach(function(row) {
    var resource = String(row['customer_negative_criterion.negative_keyword_list.shared_set'] || '');
    var match = /(?:sharedSets|shared_sets)\/(\d+)$/.exec(resource);
    if (match) accountLinks[match[1]] = true;
  });

  (tables.neg_keywords_shared || []).forEach(function(row) {
    var sharedId = String(row['shared_set.id'] || '');
    var links = campaignLinks[sharedId] || [];
    var common = {
      listName: row['shared_set.name'] || '',
      criterionId: row['shared_criterion.criterion_id'] || '',
      keyword: row['shared_criterion.keyword.text'] || '',
      matchType: row['shared_criterion.keyword.match_type'] || '',
      status: row['shared_set.status'] || '',
      sharedId: sharedId
    };
    links.forEach(function(link) {
      add_([
        'SHARED_LIST', 'CAMPAIGN_LIST', common.listName,
        link['campaign.id'] || '', link['campaign.name'] || '', '', '',
        common.criterionId, common.keyword, common.matchType, common.status, common.sharedId
      ]);
    });
    var isAccountList = row['shared_set.type'] === 'ACCOUNT_LEVEL_NEGATIVE_KEYWORDS' || accountLinks[sharedId];
    if (isAccountList) {
      add_([
        'ACCOUNT_LIST', 'ACCOUNT', common.listName, '', '', '', '',
        common.criterionId, common.keyword, common.matchType, common.status, common.sharedId
      ]);
    } else if (!links.length) {
      add_([
        'SHARED_LIST', 'SHARED_LIST_UNLINKED', common.listName, '', '', '', '',
        common.criterionId, common.keyword, common.matchType, common.status, common.sharedId
      ]);
    }
  });
  return output;
}

function derivedSourceLimitations_(sourceTabs, tabResults) {
  var limitations = [];
  (sourceTabs || []).forEach(function(tab) {
    var result = tabResults && tabResults[tab];
    if (result && result.priorPreserved) {
      throw new Error(
        'Refusing to mix current export rows with preserved source tab ' + tab +
        '. No mixed derived tab will be committed.'
      );
    }
    var status = result && result.status ? String(result.status) : 'UNAVAILABLE';
    if (status !== 'OK') limitations.push(tab + ' source status is ' + status);
  });
  return limitations;
}

function buildAudienceQueries_(range, campaignIds) {
  if (!range || !range.start || !range.end) throw new Error('Audience performance requires a date range.');
  if (!campaignIds || !campaignIds.length) throw new Error('Audience performance campaign chunk is empty.');
  var ids = campaignIds.map(function(id) {
    var value = String(id);
    if (!/^\d+$/.test(value)) throw new Error('Unsafe campaign ID: ' + value);
    return value;
  }).join(', ');
  var metrics = [
    'metrics.impressions', 'metrics.clicks', 'metrics.interactions',
    'metrics.ctr', 'metrics.average_cpc',
    'metrics.cost_micros', 'metrics.conversions', 'metrics.conversions_value',
    'metrics.all_conversions', 'metrics.all_conversions_value'
  ];
  var dateAndCampaign = "segments.date BETWEEN '" + range.start + "' AND '" + range.end +
    "'\n  AND campaign.id IN (" + ids + ')';
  return [
    {
      scope: 'CAMPAIGN',
      resourceField: 'campaign_criterion.user_list.user_list',
      criterionField: 'campaign_criterion.criterion_id',
      query: 'SELECT\n  campaign.id,\n  campaign.name,\n  campaign_criterion.criterion_id,\n' +
        '  campaign_criterion.user_list.user_list,\n  ' + metrics.join(',\n  ') +
        '\nFROM campaign_audience_view\nWHERE ' + dateAndCampaign +
        "\n  AND campaign_criterion.type = 'USER_LIST'" +
        '\nORDER BY campaign.id, campaign_criterion.criterion_id'
    },
    {
      scope: 'AD_GROUP',
      resourceField: 'ad_group_criterion.user_list.user_list',
      criterionField: 'ad_group_criterion.criterion_id',
      query: 'SELECT\n  campaign.id,\n  campaign.name,\n  ad_group.id,\n  ad_group.name,\n' +
        '  ad_group_criterion.criterion_id,\n  ad_group_criterion.user_list.user_list,\n  ' +
        metrics.join(',\n  ') + '\nFROM ad_group_audience_view\nWHERE ' + dateAndCampaign +
        "\n  AND ad_group_criterion.type = 'USER_LIST'" +
        '\nORDER BY campaign.id, ad_group.id, ad_group_criterion.criterion_id'
    }
  ];
}

function changeHistoryColumns_(includeSensitive) {
  var columns = [
    column_('change_event.change_date_time', 'date', 'change_date_time'),
    column_('change_event.client_type', 'text', 'client_type'),
    column_('change_event.change_resource_type', 'text', 'change_resource_type'),
    column_('change_event.resource_change_operation', 'text', 'resource_change_operation'),
    column_('campaign.name', 'text', 'campaign.name'),
    column_('ad_group.name', 'text', 'ad_group.name'),
    column_('change_event.change_resource_name', 'text', 'change_resource_name'),
    column_('change_event.changed_fields', 'text', 'changed_fields'),
    column_('change_event.resource_name', 'text', 'change_event_resource_name')
  ];
  if (includeSensitive) {
    columns.push(
      column_('change_event.user_email', 'text', 'user_email'),
      column_('change_event.old_resource', 'text', 'old_resource'),
      column_('change_event.new_resource', 'text', 'new_resource')
    );
  }
  return columns;
}

function buildChangeHistoryQuery_(range, cursor, includeSensitive, pageSize) {
  if (!range || !/^\d{4}-\d{2}-\d{2}$/.test(range.start) || !/^\d{4}-\d{2}-\d{2}$/.test(range.end)) {
    throw new Error('Change history requires valid frozen start and end dates.');
  }
  if (!/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(String(cursor || ''))) {
    throw new Error('Invalid Change Event pagination cursor: ' + cursor);
  }
  var limit = Number(pageSize || 10000);
  if (!isFinite(limit) || limit < 1 || limit > 10000 || Math.floor(limit) !== limit) {
    throw new Error('Change Event page size must be an integer from 1 through 10000.');
  }
  var fields = changeHistoryColumns_(includeSensitive).map(function(column) { return column.field; });
  var clientTypes = CONFIG.CHANGE_HISTORY_CLIENT_TYPES || [];
  var clientClause = '';
  if (clientTypes.length) {
    var safeClientTypes = clientTypes.map(function(clientType) {
      var value = String(clientType || '');
      if (!/^[A-Z][A-Z0-9_]*$/.test(value)) throw new Error('Unsafe Change Event client type: ' + value);
      return "'" + value + "'";
    });
    clientClause = '\n  AND change_event.client_type IN (' + safeClientTypes.join(', ') + ')';
  }
  return 'SELECT\n  ' + fields.join(',\n  ') +
    '\nFROM change_event' +
    "\nWHERE change_event.change_date_time >= '" + range.start + " 00:00:00'" +
    "\n  AND change_event.change_date_time <= '" + cursor + "'" +
    clientClause +
    '\nORDER BY change_event.change_date_time DESC' +
    '\nLIMIT ' + limit;
}

function changeEventKey_(row) {
  return String(row['change_event.resource_name'] || stableStringify_([
    row['change_event.change_date_time'], row['change_event.change_resource_name'],
    row['change_event.changed_fields'], row['change_event.resource_change_operation']
  ]));
}

function paginateChangeHistory_(range, includeSensitive, runtime, onRow, pageSize) {
  var limit = Number(pageSize || 10000);
  var cursor = range.end + ' 23:59:59';
  var startBoundary = range.start + ' 00:00:00';
  var seen = {};
  var emitted = 0;
  var pages = 0;
  while (true) {
    var query = buildChangeHistoryQuery_(range, cursor, includeSensitive, limit);
    var iterator = runtime.report(query).rows();
    var pageRows = 0;
    var newRows = 0;
    var oldest = '';
    while (iterator.hasNext()) {
      var row = iterator.next();
      pageRows++;
      var timestamp = String(row['change_event.change_date_time'] || '');
      if (!oldest || timestamp < oldest) oldest = timestamp;
      var key = changeEventKey_(row);
      if (!seen[key]) {
        seen[key] = true;
        newRows++;
        emitted++;
        onRow(row);
      }
    }
    pages++;
    if (pageRows < limit) return { rows: emitted, pages: pages, limited: false, limitation: '' };
    if (!oldest || oldest < startBoundary) {
      return { rows: emitted, pages: pages, limited: false, limitation: '' };
    }
    if (runtime.remainingSeconds &&
        Number(runtime.remainingSeconds()) <= Number(runtime.minRemainingSeconds || CONFIG.MIN_REMAINING_SECONDS)) {
      return {
        rows: emitted,
        pages: pages,
        limited: true,
        limitation: 'Change Event paging stopped before the execution time safety threshold.'
      };
    }
    if (newRows === 0 || oldest === cursor) {
      return {
        rows: emitted,
        pages: pages,
        limited: true,
        limitation: 'Inclusive Change Event cursor could not advance without risking omitted boundary rows.'
      };
    }
    cursor = oldest;
  }
}

function rangeLabel_(rangeKey) {
  if (rangeKey === 'aggregate') return 'Frozen 90 complete days ending yesterday';
  if (rangeKey === 'weekly') return 'Same frozen 90 complete days, grouped into week buckets';
  if (rangeKey === 'change') return 'Frozen 28 complete days ending yesterday';
  return 'Configuration snapshot at export time';
}

function sensitivityLabel_(tab) {
  if (tab === 'search_terms') return 'HIGH: user search queries and performance';
  if (tab === 'change_history') {
    return CONFIG.INCLUDE_SENSITIVE_CHANGE_DETAILS ?
      'HIGH: includes change-user email and old/new resource payloads' :
      'HIGH: operational changes; user email and payloads omitted';
  }
  if (tab === 'pmax_audience_signals') {
    return 'HIGH: Performance Max audience and search-theme signals';
  }
  if (tab === 'user_list_performance') {
    return 'HIGH: audience resource identifiers and performance';
  }
  return 'CONFIDENTIAL: Google Ads account configuration or performance';
}

function limitationLabel_(tab) {
  if (tab === INFO_SHEET_NAME) {
    return 'workbook_grid_cells and cell_safety_headroom are conservative safety figures captured during export and can include temporary checkpoint or removable blank-sheet overhead that is absent from the finished workbook.';
  }
  if (tab === 'search_terms') return 'Google may withhold lower-volume queries for privacy.';
  if (tab === 'landing_pages') {
    return 'landing_page_view metrics are URL-attributed and non-additive across rows, and may not reconcile to campaign totals. Do not sum this tab for campaign or account totals. Use campaign for authoritative totals; use landing_pages for URL-level analysis within the same exported date range, network, and device. CTR and average CPC are ratios and must not be summed.';
  }
  if (tab === 'ad_to_lp_map') {
    return 'Dynamic Search Ad (DSA) landing URLs are dynamically selected and are not represented in this configuration mapping. Use landing_pages for observed URL-level performance; a missing DSA mapping does not indicate an export failure.';
  }
  if (tab === 'campaign_weekly' || tab === 'ad_group_weekly' ||
      tab === 'pmax_asset_group_weekly') {
    return 'Weekly output uses the same exact 90-day reporting window as the aggregate tabs. The first and last boundary weeks can be partial. A missing entity-week means Google returned no metric row or the zero-activity row was omitted; fewer weekly rows do not indicate an export failure. Tabs are sequential query snapshots and may show small freshness differences.';
  }
  if (tab === 'pmax_audience_signals') return 'Signals are optimization hints, not targeting restrictions; availability varies. Rows are limited to the eligible asset groups in pmax_asset_groups.';
  if (tab === 'user_list_performance') return 'This is returned performance coverage, not current or complete audience inventory. Treat user_list_resource as an opaque stable Google identifier. Audience name, type, size, eligibility, membership settings, and complete zero-activity assignments are not exported; do not infer them from missing rows.';
  if (tab === 'pmax_assets') return 'Asset-association metrics are non-additive: do not sum asset rows or reconcile them directly to totals because one ad interaction can credit multiple assets. Use pmax_asset_groups or campaign for total performance.';
  if (tab === 'rsa_assets' || tab === 'demandgen_assets' || tab === 'asset_extensions') {
    return 'Asset-association metrics are non-additive: do not sum asset rows or reconcile them directly to campaign, ad-group, or ad totals because one interaction can credit multiple assets.';
  }
  if (tab === 'change_history') return 'Google exposes Change Event records only within the past 30 days, so this tab intentionally uses a safe 28-complete-day window instead of the workbook\'s 90-day performance window. It defaults to Google Ads web-client changes. Editor is not returned by Change Event despite its enum placeholder. The 10,000-row page rule applies; incomplete paging is labeled LIMITED.';
  if (tab === 'campaign_geo') {
    return 'Reported geography is inferred by Google and is not the same as configured targeting. Geographic metrics may not reconcile to campaign totals. Do not sum this tab for campaign or account totals; use campaign for authoritative totals. geo_target_most_specific_location_criterion_id and geo_target_state_criterion_id are the canonical row identity fields. Location and state names are display labels and may repeat; current readable metadata may be blank for retired or unavailable criteria. Use the declared ID key for joins, grouping, and deduplication.';
  }
  return 'Fields and resources can vary by campaign type, account eligibility, and Google Ads API support.';
}

function materialFilterLabel_(job) {
  var filters = (job.where || []).slice();
  if (job.chunked) {
    filters.push('Campaign universe includes currently non-inactive campaigns plus historical campaigns with nonzero 90-day activity');
  }
  if (job.surveyStatusFields || job.activityLookup || job.eligibility ||
      job.eligibleSource || job.kind === 'ad_to_lp_map') {
    filters.push('Inactive hierarchy entities with zero activity in the frozen 90-day survey are omitted');
  }
  return filters.join(' AND ') || 'None beyond resource compatibility';
}

function fieldDictionaryMetadata_(job, header) {
  var keyFields = {};
  ((job.dictionary && job.dictionary.keyFields) || []).forEach(function(field) {
    keyFields[field] = true;
  });
  var definition = columnDefinitionForHeader_(job, header);
  var isDerived = definition.isDerived === true || (job.derived || []).some(function(candidate) {
    return candidate.header === header;
  });
  var type = definition.type || (isDerived ? 'number' : 'text');
  var unit = '';
  if (type === 'id') {
    type = 'identifier';
    unit = 'text-preserved ID';
  } else if (type === 'date') {
    if (/date_time$/i.test(header) || header === 'change_date_time') {
      type = 'google_ads_datetime_text';
      unit = 'Google Ads yyyy-MM-dd HH:mm:ss[.fraction] text in customer.time_zone; no embedded UTC offset and not a native spreadsheet date serial';
    } else {
      type = 'iso_date_text';
      unit = 'YYYY-MM-DD ISO 8601 date stored as text; not a native spreadsheet date serial';
    }
  } else if (type === 'boolean') {
    unit = 'TRUE/FALSE';
  } else if (header === 'column_ordinal') {
    unit = '1-based output column position';
  } else if (header === 'campaign_criterion.proximity.radius') {
    unit = 'distance; see campaign_criterion.proximity.radius_units';
  } else if (/latitude_in_micro_degrees|longitude_in_micro_degrees/.test(header)) {
    unit = 'microdegrees; divide by 1,000,000 for decimal degrees';
  } else if (header === 'user_list.membership_life_span') {
    unit = 'days';
  } else if (header === 'ad_group_criterion.quality_info.quality_score') {
    unit = 'Google Quality Score from 1 through 10';
  } else if (header === 'segments.hour') {
    unit = 'hour of day from 0 through 23';
  } else if (isRawMicrosHeader_(header)) {
    type = 'number';
    unit = 'micro-units; 1,000,000 equals one account-currency unit';
  } else if (isPercentageHeader_(header)) {
    type = 'number';
    unit = 'decimal ratio; 0.10 displays as 10%';
  } else if (isCurrencyHeader_(header)) {
    type = 'number';
    unit = 'account currency';
  } else if (String(header).indexOf('per_cost') >= 0) {
    type = 'number';
    unit = 'decimal ratio';
  } else if (type === 'number') {
    unit = 'Google-reported number or count';
  } else {
    type = 'text';
    unit = isDerived ? 'Exporter-composed text or enum' : 'Google-reported text or enum';
  }
  if (definition.unit !== undefined) unit = String(definition.unit);
  var sourceFields = Array.isArray(definition.sourceFields) ? definition.sourceFields.slice() : [];
  if (!sourceFields.length && definition.field) sourceFields.push(definition.field);
  return {
    sourceFields: sourceFields.join(', '),
    dataType: type,
    unit: unit,
    isKey: Boolean(keyFields[header]),
    isDerived: isDerived,
    derivation: isDerived ? String(definition.derivation || 'Derived by the exporter') : '',
    blankWhen: definition.blankWhen === undefined ?
      'Not applicable or unavailable from Google Ads' : String(definition.blankWhen)
  };
}

function buildDataDictionaryRows_(manifest) {
  var headers = [
    'tab', 'purpose', 'row_grain', 'date_range', 'material_filters',
    'sensitive_data', 'google_side_limitations', 'keys', 'field_dictionary_reference'
  ];
  var rows = [headers];
  var entries = [{
    tab: START_HERE_SHEET_NAME,
    range: '',
    where: [],
    dictionary: {
      purpose: 'Values-only account summary, review flags, and workbook navigation',
      grain: 'mixed-layout run and account summary',
      keys: ''
    }
  }].concat(manifest.slice());
  entries.push({
    tab: INFO_SHEET_NAME,
    range: '',
    where: [],
    dictionary: {
      purpose: 'Run state, frozen ranges, row counts, limitations, and resume guidance',
      grain: 'run metadata and output tab',
      keys: 'run_id and tab'
    }
  });
  var seen = {};
  entries.forEach(function(job) {
    if (seen[job.tab]) return;
    seen[job.tab] = true;
    var dictionary = job.dictionary || {};
    rows.push([
      job.tab,
      dictionary.purpose || '',
      dictionary.grain || '',
      rangeLabel_(job.range),
      materialFilterLabel_(job),
      sensitivityLabel_(job.tab),
      limitationLabel_(job.tab),
      dictionary.keys || '',
      job.tab === INFO_SHEET_NAME || job.tab === START_HERE_SHEET_NAME ?
        '' : FIELD_DICTIONARY_SHEET_NAME
    ]);
  });
  return rows;
}

function buildFieldDictionaryRows_(manifest) {
  var rows = [[
    'tab', 'column_ordinal', 'field', 'source_fields', 'data_type',
    'unit', 'is_key', 'is_derived', 'derivation', 'blank_when'
  ]];
  (manifest || []).forEach(function(job) {
    if (!job || !Array.isArray(job.columns)) return;
    headersForJob_(job).forEach(function(header, index) {
      var metadata = fieldDictionaryMetadata_(job, header);
      rows.push([
        job.tab,
        index + 1,
        header,
        metadata.sourceFields,
        metadata.dataType,
        metadata.unit,
        metadata.isKey,
        metadata.isDerived,
        metadata.derivation,
        metadata.blankWhen
      ]);
    });
  });
  return rows;
}

function finiteNumberOrZero_(value) {
  var number = Number(value);
  return isFinite(number) ? number : 0;
}

function safeRatio_(numerator, denominator) {
  return denominator ? numerator / denominator : '';
}

function campaignSummaryMetrics_(rows) {
  return (rows || []).reduce(function(total, row) {
    total.impressions += finiteNumberOrZero_(row['metrics.impressions']);
    total.clicks += finiteNumberOrZero_(row['metrics.clicks']);
    total.interactions += finiteNumberOrZero_(row['metrics.interactions']);
    total.cost += finiteNumberOrZero_(row['metrics.cost_micros']) / 1000000;
    total.conversions += finiteNumberOrZero_(row['metrics.conversions']);
    total.conversionValue += finiteNumberOrZero_(row['metrics.conversions_value']);
    return total;
  }, {
    impressions: 0, clicks: 0, interactions: 0,
    cost: 0, conversions: 0, conversionValue: 0
  });
}

function finalizedSummaryMetrics_(totals) {
  return {
    cost: totals.cost,
    impressions: totals.impressions,
    clicks: totals.clicks,
    ctr: safeRatio_(totals.clicks, totals.impressions),
    conversions: totals.conversions,
    conversionRate: safeRatio_(totals.conversions, totals.interactions),
    costPerConversion: safeRatio_(totals.cost, totals.conversions),
    conversionValue: totals.conversionValue
  };
}

function startHereChannelRows_(campaignRows) {
  var grouped = {};
  (campaignRows || []).forEach(function(row) {
    var channel = String(row['campaign.advertising_channel_type'] || 'UNSPECIFIED');
    grouped[channel] = grouped[channel] || [];
    grouped[channel].push(row);
  });
  return Object.keys(grouped).sort().map(function(channel) {
    var summary = finalizedSummaryMetrics_(campaignSummaryMetrics_(grouped[channel]));
    summary.channel = channel;
    return summary;
  });
}

function stableCampaignReviewSort_(left, right) {
  if (right.cost !== left.cost) return right.cost - left.cost;
  var leftId = String(left.campaignId || '');
  var rightId = String(right.campaignId || '');
  return leftId < rightId ? -1 : (leftId > rightId ? 1 : 0);
}

function booleanTrue_(value) {
  return value === true || value === 1 || String(value || '').toUpperCase() === 'TRUE';
}

function startHereReviewRows_(state, campaignRows) {
  var current = state || {};
  var errors = [];
  var limitations = [];
  Object.keys(current.tabs || {}).sort().forEach(function(tab) {
    var result = current.tabs[tab] || {};
    var status = String(result.status || '');
    var detail = String(result.error || result.limitation || status);
    if (status === 'ERROR' || status === 'ERROR_PREVIOUS_PRESERVED') {
      errors.push({
        kind: 'ERROR_COVERAGE', severity: 'ERROR', tab: tab, campaignId: '',
        campaignName: '', cost: '', conversions: '', detail: detail
      });
    } else if (status === 'LIMITED') {
      limitations.push({
        kind: 'LIMITED_COVERAGE',
        severity: 'LIMITED', tab: tab, campaignId: '',
        campaignName: '', cost: '', conversions: '', detail: detail
      });
    }
  });

  var zeroConversion = [];
  var budgetRecommendations = [];
  (campaignRows || []).forEach(function(row) {
    var cost = finiteNumberOrZero_(row['metrics.cost_micros']) / 1000000;
    var conversions = finiteNumberOrZero_(row['metrics.conversions']);
    var campaignId = String(row['campaign.id'] || '');
    var campaignName = String(row['campaign.name'] || '');
    if (cost > 0 && conversions === 0) {
      zeroConversion.push({
        kind: 'ZERO_CONVERSION_SPEND', severity: 'REVIEW', tab: 'campaign',
        campaignId: campaignId, campaignName: campaignName, cost: cost,
        conversions: conversions,
        detail: 'Campaign has nonzero cost and zero conversions in the aggregate range.'
      });
    }
    var currentBudget = finiteNumberOrZero_(row['campaign_budget.amount_micros']) / 1000000;
    var recommendedBudget =
      finiteNumberOrZero_(row['campaign_budget.recommended_budget_amount_micros']) / 1000000;
    if (booleanTrue_(row['campaign_budget.has_recommended_budget']) &&
        recommendedBudget > currentBudget) {
      budgetRecommendations.push({
        kind: 'BUDGET_RECOMMENDATION', severity: 'REVIEW', tab: 'campaign',
        campaignId: campaignId, campaignName: campaignName, cost: cost,
        conversions: conversions, currentBudget: currentBudget,
        recommendedBudget: recommendedBudget,
        detail: 'Google reports a recommended campaign budget above the current budget.'
      });
    }
  });
  zeroConversion.sort(stableCampaignReviewSort_);
  budgetRecommendations.sort(stableCampaignReviewSort_);
  var omitted = [];
  if (zeroConversion.length > 25) {
    omitted.push({
      kind: 'OMITTED_COUNT', severity: 'INFO', tab: 'campaign', count: zeroConversion.length - 25,
      campaignId: '', campaignName: '', cost: '', conversions: '',
      detail: String(zeroConversion.length - 25) +
        ' additional campaigns with spend and zero conversions omitted; review campaign for the complete list.'
    });
  }
  if (budgetRecommendations.length > 25) {
    omitted.push({
      kind: 'OMITTED_COUNT', severity: 'INFO', tab: 'campaign',
      count: budgetRecommendations.length - 25, campaignId: '', campaignName: '',
      cost: '', conversions: '', detail: String(budgetRecommendations.length - 25) +
        ' additional budget recommendations omitted; review campaign for the complete list.'
    });
  }
  return errors.concat(
    limitations,
    zeroConversion.slice(0, 25),
    budgetRecommendations.slice(0, 25),
    omitted
  );
}

function startHereRecommendedUse_(group) {
  var guidance = {
    performance: 'Use to analyze account results and performance drivers.',
    structure: 'Use to audit current entity structure and active inventory.',
    creative: 'Use to review ads, assets, and creative associations.',
    targeting: 'Use to review configured location and proximity targeting.',
    negatives: 'Use to audit negative-keyword coverage and provenance.',
    audiences: 'Use to review returned audience performance and Performance Max signals; do not treat these tabs as complete audience inventory.',
    governance: 'Use to verify configuration, provenance, coverage, and change history.'
  };
  return guidance[group] || 'Use for workbook review.';
}

function startHereDirectoryRows_(manifest, state) {
  var current = state || {};
  var dictionary = buildDataDictionaryRows_(manifest || []);
  var headers = dictionary[0];
  var byTab = {};
  dictionary.slice(1).forEach(function(row) {
    var object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    byTab[object.tab] = object;
  });
  var output = [];
  [
    'performance', 'structure', 'creative', 'targeting',
    'negatives', 'audiences', 'governance'
  ].forEach(function(group) {
    START_HERE_TAB_GROUPS[group].forEach(function(tab) {
      var item = byTab[tab];
      if (!item) throw new Error('START_HERE directory is missing dictionary metadata for ' + tab + '.');
      var result = tab === INFO_SHEET_NAME ? {
        status: nativeWorkbookStatus_(current),
        rows: Math.max(0, buildExportInfoRows_(current, manifest || []).length - 8)
      } : ((current.tabs || {})[tab] || {});
      output.push({
        group: group,
        tab: tab,
        status: String(result.status || (current.status === 'RUNNING' ? 'PENDING' : 'NOT_REPORTED')),
        rows: typeof result.rows === 'number' && isFinite(result.rows) ? result.rows : '',
        purpose: item.purpose,
        grain: item.row_grain,
        dateRange: item.date_range,
        recommendedUse: startHereRecommendedUse_(group)
      });
    });
  });
  return output;
}

function startHereReviewLabel_(review) {
  var labels = {
    ERROR_COVERAGE: 'Export error',
    LIMITED_COVERAGE: 'Limited data coverage',
    ZERO_CONVERSION_SPEND: 'Spend with zero conversions',
    BUDGET_RECOMMENDATION: 'Google budget recommendation',
    OMITTED_COUNT: 'Additional review items'
  };
  return labels[String(review && review.kind || '')] || String(review && review.kind || 'Review item');
}

function startHereReviewDetail_(review) {
  return String(review && review.detail || '');
}

function startHereDirectoryStatusLabel_(item) {
  var status = String(item && item.status || '');
  if (status === 'OK' && Number(item && item.rows) === 0) {
    return 'OK — no matching records';
  }
  return status;
}

function startHereDataStatusLabel_(status) {
  var labels = {
    COMPLETE: 'Complete',
    COMPLETE_WITH_LIMITATIONS: 'Complete with limitations',
    COMPLETE_WITH_ERRORS: 'Finished with errors',
    FINALIZING: 'Finalizing',
    PAUSED: 'Paused; run main again to resume',
    RUNNING: 'Export in progress',
    RESET: 'Reset; run main to export'
  };
  var value = String(status || '');
  return labels[value] || value;
}

function hasPriorPreservedTabs_(state) {
  return Object.keys(state && state.tabs || {}).some(function(tab) {
    var result = state.tabs[tab];
    return Boolean(result && (result.priorPreserved || result.status === 'ERROR_PREVIOUS_PRESERVED'));
  });
}

function hasJobErrorTabs_(state) {
  return Object.keys(state && state.tabs || {}).some(function(tab) {
    var result = state.tabs[tab];
    return Boolean(result &&
      (result.status === 'ERROR' || result.status === 'ERROR_PREVIOUS_PRESERVED'));
  });
}

function nativeWorkbookStatus_(state) {
  var current = state || {};
  if (current.workbookError || hasPriorPreservedTabs_(current)) return 'NEEDS_REVIEW';
  if (current.status === 'COMPLETE') return 'READY';
  if (current.status === 'COMPLETE_WITH_LIMITATIONS') return 'READY_WITH_LIMITATIONS';
  if (current.status === 'RUNNING' || current.status === 'PAUSED' ||
      current.status === 'FINALIZING') return 'IN_PROGRESS';
  return 'NEEDS_REVIEW';
}

function canonicalSpreadsheetUrl_(spreadsheetId) {
  var id = String(spreadsheetId || '').trim();
  if (!id || !/^[A-Za-z0-9_-]+$/.test(id)) return '';
  return 'https://docs.google.com/spreadsheets/d/' + id + '/edit';
}

function buildStartHereModel_(state, campaignRows, manifest) {
  var current = state || {};
  return {
    account: {
      id: String(current.accountId || ''),
      name: String(current.accountName || ''),
      currencyCode: String(current.accountCurrencyCode || ''),
      range: current.ranges && current.ranges.aggregate ? current.ranges.aggregate : {},
      dataStatus: String(current.status || '')
    },
    kpis: finalizedSummaryMetrics_(campaignSummaryMetrics_(campaignRows)),
    channels: startHereChannelRows_(campaignRows),
    reviews: startHereReviewRows_(current, campaignRows),
    directory: startHereDirectoryRows_(manifest || [], current)
  };
}

function buildStartHereRows_(model) {
  var current = model || {};
  var account = current.account || {};
  var kpis = current.kpis || finalizedSummaryMetrics_(campaignSummaryMetrics_([]));
  var rows = [];
  function add_(values) {
    var row = (values || []).slice(0, 8);
    while (row.length < 8) row.push('');
    row.forEach(function(value) {
      if (typeof value === 'string' && value.charAt(0) === '=') {
        throw new Error('START_HERE values must not contain formulas.');
      }
    });
    rows.push(row);
  }
  add_(['Google Ads Analysis Workbook', VERSION, 'Values-only marketer summary']);
  add_(['Confidential account export. Share only with authorized recipients.']);
  add_([
    'Account name: ' + String(account.name || ''),
    'Customer ID: ' + String(account.id || ''),
    'Reporting range: ' + String((account.range || {}).start || '') + ' to ' +
      String((account.range || {}).end || ''),
    'Data status: ' + startHereDataStatusLabel_(account.dataStatus),
    'Deliverable: This Google Sheet',
    'XLSX downloads: sanitize with the bundled local tool before upload or sharing.'
  ]);
  add_([]);
  add_(['ACCOUNT KPIS']);
  add_([
    'Cost', kpis.cost, 'Impressions', kpis.impressions,
    'Clicks', kpis.clicks, 'CTR', kpis.ctr
  ]);
  add_([
    'Conversions', kpis.conversions, 'Conversion rate', kpis.conversionRate,
    'Cost / conversion', kpis.costPerConversion,
    'Conversion value', kpis.conversionValue
  ]);
  add_([]);
  add_(['CHANNEL SUMMARY']);
  add_([
    'Channel', 'Cost', 'Impressions', 'Clicks', 'CTR',
    'Conversions', 'Conversion rate', 'Cost / conversion'
  ]);
  (current.channels || []).forEach(function(channel) {
    add_([
      channel.channel, channel.cost, channel.impressions, channel.clicks, channel.ctr,
      channel.conversions, channel.conversionRate, channel.costPerConversion
    ]);
  });
  add_([]);
  add_(['REVIEW FIRST']);
  add_(['Severity', 'Fact', 'Tab', 'Campaign ID', 'Campaign', 'Cost', 'Conversions', 'Detail']);
  if (!(current.reviews || []).length) {
    add_(['INFO', 'No deterministic review flags were produced']);
  } else {
    current.reviews.forEach(function(review) {
      add_([
        review.severity || '', startHereReviewLabel_(review), review.tab || '', review.campaignId || '',
        review.campaignName || '', review.cost === undefined ? '' : review.cost,
        review.conversions === undefined ? '' : review.conversions,
        startHereReviewDetail_(review)
      ]);
    });
  }
  add_([]);
  add_(['WORKBOOK DIRECTORY']);
  add_(['Group', 'Tab', 'Status', 'Rows', 'Purpose', 'Row grain', 'Date range', 'Recommended use']);
  (current.directory || []).forEach(function(item) {
    add_([
      item.group, item.tab, startHereDirectoryStatusLabel_(item), item.rows,
      item.purpose, item.grain, item.dateRange, item.recommendedUse
    ]);
  });
  return rows;
}

function formatTimestampIso_(value) {
  var number = Number(value || 0);
  if (!isFinite(number) || number <= 0) return '';
  return new Date(number).toISOString();
}

function buildExportInfoRows_(state, manifest) {
  function range_(key) {
    var range = state.ranges && state.ranges[key];
    return range ? range.start + ' through ' + range.end : '';
  }
  function detail_(result) {
    if (!result) return '';
    var details = [];
    [result.limitation, result.error].forEach(function(value) {
      var text = String(value || '');
      if (text && details.indexOf(text) < 0) details.push(text);
    });
    return details.join(' | ');
  }
  var hasPriorPreserved = hasPriorPreservedTabs_(state);
  var hasJobErrors = hasJobErrorTabs_(state);
  var instruction;
  if (state.status === 'RESET') {
    instruction = 'Checkpoint reset completed. Existing final report tabs were preserved and may use an older output schema. Run main() to regenerate the export before analysis.';
  } else if (state.status === 'PAUSED' || state.status === 'RUNNING' ||
      state.status === 'FINALIZING') {
    instruction = 'Run main() again to resume this frozen export. If the run reports that the ' +
      'recovery checkpoint is missing, set ALLOW_RESET=true, run resetExportState() once, ' +
      'restore ALLOW_RESET=false, and start a replacement export with main().';
  } else if (hasJobErrors) {
    instruction = 'The export has one or more failed jobs and cannot resume those jobs in place. ' +
      (hasPriorPreserved ?
        'One or more tabs preserved prior data, so the workbook may contain mixed runs or schemas. ' : '') +
      'Review every ERROR row, then set ALLOW_RESET=true, run resetExportState() once, ' +
      'restore ALLOW_RESET=false, and start a replacement export with main().';
  } else if (state.workbookError) {
    instruction = /Could not retain the checkpoint/i.test(String(state.workbookError)) ?
      'The native Google Sheets workbook is not ready and has no resumable checkpoint. ' +
        'Set ALLOW_RESET=true, run resetExportState() once, restore ALLOW_RESET=false, ' +
        'and start a replacement export with main(): ' + state.workbookError :
      'The native Google Sheets workbook is not ready because finalization failed. ' +
        'Run main() again to retry: ' + state.workbookError;
  } else if (state.status === 'COMPLETE' || state.status === 'COMPLETE_WITH_LIMITATIONS') {
    if (hasPriorPreserved) {
      instruction = 'Export finished with preserved prior tabs; the workbook may contain more than ' +
        'one output schema and is not ready for analysis. Regenerate the failed tabs.';
    } else {
      instruction = state.status === 'COMPLETE_WITH_LIMITATIONS' ?
        'The native Google Sheets workbook is complete with limited coverage. Review every LIMITED row ' +
          'before analysis. No additional step is required for native use. If you manually download XLSX, ' +
          'run the bundled sanitize-downloaded-xlsx.js tool before upload or sharing. Run main() again ' +
          'when you want a fresh export.' :
        'The native Google Sheets workbook is complete and ready for analysis. No additional step is ' +
          'required for native use. If you manually download XLSX, run the bundled ' +
          'sanitize-downloaded-xlsx.js tool before upload or sharing. Run main() again when you want ' +
          'a fresh export.';
    }
  } else if (state.status === 'COMPLETE_WITH_ERRORS') {
    instruction = hasPriorPreserved ?
      'Export finished with errors and preserved prior tabs; the workbook may contain more than one output schema. Regenerate failed tabs before analysis.' :
      'Export finished with errors. Review ERROR rows before analysis.';
  } else {
    instruction = 'Run main() to create or resume the Google Ads account export.';
  }
  var gridCells = state.workbookGridCells === undefined ? '' : Number(state.workbookGridCells);
  var cellLimit = state.workbookCellSafetyLimit === undefined ? '' : Number(state.workbookCellSafetyLimit);
  var headroom = gridCells === '' || cellLimit === '' ? '' : Math.max(0, cellLimit - gridCells);
  var terminalSchemaStatus = state.status === 'COMPLETE' ||
    state.status === 'COMPLETE_WITH_LIMITATIONS' || state.status === 'FINALIZING';
  var realizedOutputSchemaVersion = terminalSchemaStatus && !hasPriorPreserved ?
    (state.outputSchemaVersion === undefined ? OUTPUT_SCHEMA_VERSION : state.outputSchemaVersion) : '';
  var realizedRuntimeContractVersion = terminalSchemaStatus && !hasPriorPreserved ?
    (state.runtimeContractVersion === undefined ?
      RUNTIME_CONTRACT_VERSION : state.runtimeContractVersion) : '';
  var snapshotWarning =
    'Tabs are sequential, non-atomic query snapshots. Google may revise historical attribution, ' +
    'late conversions, or invalid-traffic adjustments while a multi-minute export runs. Small ' +
    'same-range differences can therefore occur; use campaign as the authoritative aggregate baseline.';
  var rows = [
    [
      OWNER_KEY, VERSION, 'output_schema_version',
      realizedOutputSchemaVersion,
      'Confidential Google Ads account export', '', '', ''
    ],
    ['run_id', state.runId || '', 'overall_status', state.status || '', 'started_at', formatTimestampIso_(state.startedAtMs), 'workbook_status', nativeWorkbookStatus_(state)],
    ['account_id', state.accountId || '', 'account_name', state.accountName || '', 'updated_at', formatTimestampIso_(state.updatedAtMs), 'deliverable_type', 'NATIVE_GOOGLE_SHEET'],
    ['aggregate_range', range_('aggregate'), 'weekly_range', range_('weekly'), 'change_range', range_('change'), 'last_complete_day', state.ranges && state.ranges.aggregate ? state.ranges.aggregate.end || '' : ''],
    ['current_job', state.currentJobId || '', 'job_index', state.jobIndex || 0, 'manifest_jobs', manifest.length, 'workbook_url', canonicalSpreadsheetUrl_(state.spreadsheetId)],
    ['next_action', instruction, 'workbook_grid_cells', gridCells, 'cell_safety_headroom', headroom, 'reporting_window', 'LAST_90_COMPLETE_DAYS'],
    [
      'snapshot_semantics', snapshotWarning,
      'runtime_contract_version', realizedRuntimeContractVersion,
      'checkpoint_schema_version', STATE_SCHEMA_VERSION,
      'refresh_behavior', 'Run main() again after completion for a fresh export.'
    ],
    [
      'tab', 'status', 'rows', 'duration_seconds',
      'source_read_started_at', 'source_read_completed_at',
      'prior_data_preserved', 'limitation_or_error'
    ]
  ];
  manifest.forEach(function(job) {
    var result = state.tabs && state.tabs[job.tab];
    rows.push([
      job.tab,
      result ? result.status : 'PENDING',
      result ? result.rows : 0,
      result ? Number(result.durationMs || 0) / 1000 : 0,
      result ? formatTimestampIso_(result.sourceReadStartedAtMs) : '',
      result ? formatTimestampIso_(result.sourceReadCompletedAtMs) : '',
      result && result.priorPreserved ? 'YES' : 'NO',
      detail_(result)
    ]);
  });
  return rows;
}

function redactDiagnosticSample_(row) {
  var redacted = {};
  Object.keys(row || {}).forEach(function(field) {
    var lower = field.toLowerCase();
    var sensitive = /(^|\.)(id|name|resource_name|user_email|search_term)(\.|$)/.test(lower) ||
      /url|keyword\.text|audience|phone|address|criterion_id|segments\.geo_target_/.test(lower);
    redacted[field] = sensitive ? '[REDACTED]' : row[field];
  });
  return redacted;
}

// --------------------------------------------------------------------------
// Safe Sheets writer and transactional sheet replacement
// --------------------------------------------------------------------------

function withRetry_(operation, options) {
  options = options || {};
  var retries = Math.max(1, Number(options.retries || 4));
  var sleeper = options.sleep || function(ms) { Utilities.sleep(ms); };
  var lastError;
  for (var attempt = 0; attempt < retries; attempt++) {
    try { return operation(); }
    catch (error) {
      lastError = error;
      if (attempt + 1 >= retries) break;
      sleeper(Math.min(500 * Math.pow(2, attempt), 5000));
    }
  }
  throw lastError;
}

function ensureSheetCapacity_(sheet, requiredLastRow, requiredLastColumn, cellLimit) {
  var spreadsheet = sheet.getParent();
  var currentRows = sheet.getMaxRows();
  var currentColumns = sheet.getMaxColumns();
  var targetRows = Math.max(currentRows, requiredLastRow);
  var targetColumns = Math.max(currentColumns, requiredLastColumn);
  var currentCells = spreadsheet.getSheets().reduce(function(total, candidate) {
    return total + (candidate.getMaxRows() * candidate.getMaxColumns());
  }, 0);
  var projected = currentCells - (currentRows * currentColumns) + (targetRows * targetColumns);
  if (projected > Number(cellLimit)) {
    throw new Error('Workbook cell safety limit would be exceeded (' + projected +
      ' projected; limit ' + cellLimit + ').');
  }
  if (targetRows > currentRows) sheet.insertRowsAfter(currentRows, targetRows - currentRows);
  if (targetColumns > currentColumns) sheet.insertColumnsAfter(currentColumns, targetColumns - currentColumns);
}

function workbookGridCellCount_(spreadsheet) {
  return spreadsheet.getSheets().reduce(function(total, sheet) {
    return total + (Number(sheet.getMaxRows()) * Number(sheet.getMaxColumns()));
  }, 0);
}

function workbookGridCellCountExcluding_(spreadsheet, excludedNames) {
  var excluded = {};
  (excludedNames || []).forEach(function(name) { excluded[String(name)] = true; });
  return spreadsheet.getSheets().reduce(function(total, sheet) {
    if (excluded[String(sheet.getName())]) return total;
    return total + (Number(sheet.getMaxRows()) * Number(sheet.getMaxColumns()));
  }, 0);
}

function trimSheetGrid_(sheet) {
  var frozenRows = typeof sheet.getFrozenRows === 'function' ? Number(sheet.getFrozenRows() || 0) : 0;
  var frozenColumns = typeof sheet.getFrozenColumns === 'function' ? Number(sheet.getFrozenColumns() || 0) : 0;
  var usedRows = Math.max(1, Number(sheet.getLastRow() || 0), frozenRows + 1);
  var usedColumns = Math.max(1, Number(sheet.getLastColumn() || 0), frozenColumns + 1);
  var extraRows = Number(sheet.getMaxRows()) - usedRows;
  var extraColumns = Number(sheet.getMaxColumns()) - usedColumns;
  if (extraRows > 0) sheet.deleteRows(usedRows + 1, extraRows);
  if (extraColumns > 0) sheet.deleteColumns(usedColumns + 1, extraColumns);
  return sheet;
}

function createSafeRowBuffer_(sheet, headers, options) {
  options = options || {};
  var batchRows = Math.max(1, Number(options.batchRows || CONFIG.BATCH_ROWS || 1000));
  var cellLimit = Number(options.cellLimit || CONFIG.WORKBOOK_CELL_SAFETY_LIMIT || 9000000);
  var buffer = [];
  var written = 0;
  var nextRow = sheet.getLastRow() + 1;
  function flush_() {
    if (!buffer.length) return;
    var rows = buffer;
    ensureSheetCapacity_(sheet, nextRow + rows.length - 1, headers.length, cellLimit);
    withRetry_(function() {
      sheet.getRange(nextRow, 1, rows.length, headers.length).setValues(rows);
    }, options);
    nextRow += rows.length;
    buffer = [];
  }
  return {
    push: function(row) {
      buffer.push(encodeSheetRowForWrite_(row, headers));
      written++;
      if (buffer.length >= batchRows) flush_();
    },
    flush: flush_,
    count: function() { return written; }
  };
}

function rollbackPartialChunk_(sheet, startRow) {
  var first = Math.max(1, Number(startRow));
  var last = sheet.getLastRow();
  if (last < first) return 0;
  var count = last - first + 1;
  sheet.getRange(first, 1, count, Math.max(1, sheet.getLastColumn())).clearContent();
  return count;
}

function columnDefinitionForHeader_(job, header) {
  var columns = (job.columns || []).concat(job.derived || []);
  for (var index = 0; index < columns.length; index++) {
    if (columns[index].header === header) return columns[index];
  }
  return { header: header, type: 'text' };
}

function isRawMicrosHeader_(header) {
  var value = String(header || '');
  return /_micros$/.test(value) || value === 'metrics.average_cpc';
}

function isPercentageHeader_(header) {
  var value = String(header || '').toLowerCase();
  return value === 'metrics.ctr' || value === 'conversion_rate' ||
    value.indexOf('impression_share') >= 0 || value === 'campaign.optimization_score';
}

function isCurrencyHeader_(header) {
  var value = String(header || '').toLowerCase();
  return value === 'cost' || value === 'average_cpc' || value === 'cost_per_conversion' ||
    value === 'metrics.conversions_value' || value === 'metrics.all_conversions_value' ||
    /\.(amount|recommended_budget_amount|total_amount|target_cpa|cpc_bid|cpm_bid)$/.test(value) ||
    value === 'conversion_action.value_settings.default_value';
}

function currencyNumberFormat_(currencyCode) {
  var code = String(currencyCode || '').toUpperCase().replace(/[^A-Z]/g, '').substring(0, 3);
  return code ? '"' + code + '" #,##0.00' : '#,##0.00';
}

function numberFormatForHeader_(job, header, currencyCode) {
  var definition = columnDefinitionForHeader_(job, header);
  if (definition.type === 'id' || definition.type === 'text' || definition.type === 'date') return '@';
  if (isPercentageHeader_(header)) return '0.00%';
  if (isRawMicrosHeader_(header)) return '#,##0';
  if (isCurrencyHeader_(header)) return currencyNumberFormat_(currencyCode);
  if (String(header || '') === 'column_ordinal' ||
      /\.(impressions|clicks|interactions)$/.test(String(header || '')) ||
      /(^|\.)(hour|membership_life_span|size_for_search|size_for_display)$/.test(String(header || ''))) {
    return '#,##0';
  }
  if (definition.type === 'boolean') return 'General';
  return '#,##0.00';
}

function columnWidthForHeader_(header) {
  var value = String(header || '').toLowerCase();
  if (value === 'field') return 240;
  if (value === 'source_fields') return 280;
  if (value === 'unit') return 240;
  if (value === 'derivation') return 320;
  if (value === 'blank_when') return 280;
  if (/url|search_term|description|limitation|field_paths|payload|address|resource_name|changed_fields/.test(value)) return 320;
  if (/name|keyword\.text|text_asset|asset\.text|purpose|material_filters|google_side_limitations|keys/.test(value)) return 260;
  if (/status|type|scope|device|network|match_type|field_type|source/.test(value)) return 170;
  if (/date|time|week/.test(value)) return 140;
  if (/(^|\.)(id|criterion_id)$/.test(value)) return 150;
  if (/metrics|cost|conversion|budget|bid|score|share|rows|duration/.test(value)) return 150;
  return 180;
}

function headerRowHeight_(headers) {
  var estimatedLines = (headers || []).reduce(function(maximum, header) {
    var width = columnWidthForHeader_(header);
    var charactersPerLine = Math.max(10, Math.floor(width / 8));
    return Math.max(maximum, Math.ceil(String(header || '').length / charactersPerLine));
  }, 1);
  return Math.max(56, Math.min(96, estimatedLines * 18));
}

function styleHeaderRange_(range) {
  range.setBackground('#1F4E78')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setShowHyperlink(false)
    .setVerticalAlignment('middle')
    .setWrap(true);
}

function formatReportSheet_(sheet, job, currencyCode, emptyNote) {
  var headers = headersForJob_(job);
  var lastColumn = headers.length;
  var lastRow = Math.max(1, sheet.getLastRow());
  if (!lastColumn) return sheet;

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(Math.min(Number(CONFIG.FREEZE_CONTEXT_COLUMNS || 0), lastColumn));
  sheet.setRowHeight(1, headerRowHeight_(headers));
  styleHeaderRange_(sheet.getRange(1, 1, 1, lastColumn));
  if (lastRow === 1) {
    var noteRange = sheet.getRange(1, 1, 1, 1);
    if (typeof noteRange.setNote === 'function') {
      noteRange.setNote(emptyNote ||
        'Export completed successfully. Google Ads returned no matching records for this account and reporting range.');
    }
  }

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (lastRow > 1) sheet.getRange(1, 1, lastRow, lastColumn).createFilter();

  headers.forEach(function(header, index) {
    var column = index + 1;
    sheet.setColumnWidth(column, columnWidthForHeader_(header));
    if (lastRow > 1) {
      sheet.getRange(2, column, lastRow - 1, 1)
        .setNumberFormat(numberFormatForHeader_(job, header, currencyCode));
    }
    if (Boolean(CONFIG.HIDE_RAW_MICROS_COLUMNS) && isRawMicrosHeader_(header)) {
      sheet.hideColumns(column);
    }
    var wrapBody =
      (job.tab === 'campaign' && header === 'campaign.name') ||
      (job.tab === 'change_history' && [
        'campaign.name', 'ad_group.name', 'change_resource_name',
        'changed_fields', 'change_event_resource_name'
      ].indexOf(header) >= 0);
    if (lastRow > 1 && wrapBody) {
      sheet.getRange(2, column, lastRow - 1, 1).setWrap(true).setVerticalAlignment('top');
    }
  });

  if ((job.tab === DICTIONARY_SHEET_NAME || job.tab === FIELD_DICTIONARY_SHEET_NAME) &&
      lastRow > 1 && lastColumn > 1) {
    sheet.getRange(2, 2, lastRow - 1, lastColumn - 1).setWrap(true).setVerticalAlignment('top');
  }
  return sheet;
}

function formatExportInfoSheet_(sheet) {
  var lastRow = Math.max(1, sheet.getLastRow());
  var lastColumn = Math.max(1, sheet.getLastColumn());
  sheet.setFrozenRows(Math.min(1, lastRow));
  sheet.setFrozenColumns(0);
  sheet.setRowHeight(1, 34);
  styleHeaderRange_(sheet.getRange(1, 1, 1, lastColumn));
  if (lastRow >= 8) {
    sheet.setRowHeight(8, 34);
    styleHeaderRange_(sheet.getRange(8, 1, 1, lastColumn));
  }
  sheet.getRange(1, 1, lastRow, lastColumn).setWrap(true).setVerticalAlignment('middle');
  [180, 300, 170, 240, 205, 205, 170, 420].forEach(function(width, index) {
    if (index < lastColumn) sheet.setColumnWidth(index + 1, width);
  });
  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (lastRow > 8) sheet.getRange(8, 1, lastRow - 7, lastColumn).createFilter();
  if (lastRow > 8 && lastColumn >= 4) {
    sheet.getRange(9, 3, lastRow - 8, 1).setNumberFormat('#,##0');
    sheet.getRange(9, 4, lastRow - 8, 1).setNumberFormat('#,##0.0');
  }
  if (lastRow >= 2 && lastColumn >= 4) {
    var overall = statusPalette_(sheet.getRange(2, 4, 1, 1).getDisplayValues()[0][0]);
    if (overall) {
      sheet.getRange(2, 4, 1, 1)
        .setBackground(overall.background)
        .setFontColor(overall.font)
        .setFontWeight('bold');
    }
  }
  if (lastRow >= 2 && lastColumn >= 8) {
    var workbookStatus = statusPalette_(sheet.getRange(2, 8, 1, 1).getDisplayValues()[0][0]);
    if (workbookStatus) {
      sheet.getRange(2, 8, 1, 1)
        .setBackground(workbookStatus.background)
        .setFontColor(workbookStatus.font)
        .setFontWeight('bold');
    }
  }
  if (lastRow > 8 && lastColumn >= 2) {
    sheet.getRange(9, 2, lastRow - 8, 1).getDisplayValues().forEach(function(row, index) {
      var status = String(row[0] || '');
      if (status !== 'LIMITED' && status !== 'ERROR' && status !== 'ERROR_PREVIOUS_PRESERVED') return;
      var palette = statusPalette_(status);
      if (!palette) return;
      sheet.getRange(index + 9, 1, 1, lastColumn)
        .setBackground(palette.background)
        .setFontColor(palette.font);
    });
  }
  return sheet;
}

function statusPalette_(status) {
  var value = String(status || '').toUpperCase();
  if (['COMPLETE', 'OK', 'READY'].indexOf(value) >= 0) {
    return { background: '#D9EAD3', font: '#274E13' };
  }
  if ([
    'COMPLETE_WITH_LIMITATIONS', 'LIMITED', 'PAUSED', 'FINALIZING', 'IN_PROGRESS',
    'READY_WITH_LIMITATIONS'
  ].indexOf(value) >= 0) {
    return { background: '#FFF2CC', font: '#7F6000' };
  }
  if (['COMPLETE_WITH_ERRORS', 'ERROR', 'ERROR_PREVIOUS_PRESERVED', 'NEEDS_REVIEW'].indexOf(value) >= 0) {
    return { background: '#F4CCCC', font: '#990000' };
  }
  return null;
}

function formatStartHereSheet_(sheet, currencyCode) {
  var lastRow = Math.max(1, sheet.getLastRow());
  var lastColumn = Math.max(1, sheet.getLastColumn());
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var sectionRows = {};
  values.forEach(function(row, index) {
    var label = String(row[0] || '');
    if (['ACCOUNT KPIS', 'CHANNEL SUMMARY', 'REVIEW FIRST', 'WORKBOOK DIRECTORY'].indexOf(label) >= 0) {
      sectionRows[label] = index + 1;
    }
  });
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setRowHeight(1, 42);
  if (lastRow >= 3) {
    sheet.setRowHeight(3, 72);
    sheet.getRange(3, 1, 1, lastColumn).setWrap(true);
  }
  [190, 240, 200, 110, 280, 220, 220, 360].forEach(function(width, index) {
    if (index < lastColumn) sheet.setColumnWidth(index + 1, width);
  });
  sheet.getRange(1, 1, lastRow, lastColumn)
    .setBackground('#FFFFFF')
    .setFontColor('#1F1F1F')
    .setFontWeight('normal')
    .setShowHyperlink(false)
    .setVerticalAlignment('middle');
  styleHeaderRange_(sheet.getRange(1, 1, 1, lastColumn));
  Object.keys(sectionRows).forEach(function(label) {
    var rowNumber = sectionRows[label];
    sheet.setRowHeight(rowNumber, 34);
    styleHeaderRange_(sheet.getRange(rowNumber, 1, 1, lastColumn));
    if (rowNumber < lastRow) {
      sheet.getRange(rowNumber + 1, 1, 1, lastColumn)
        .setBackground('#D9EAF7')
        .setFontColor('#1F1F1F')
        .setFontWeight('bold')
        .setShowHyperlink(false)
        .setWrap(true);
    }
  });
  var moneyFormat = currencyNumberFormat_(currencyCode);
  if (sectionRows['ACCOUNT KPIS']) {
    var firstKpiRow = sectionRows['ACCOUNT KPIS'] + 1;
    [2, 6].forEach(function(column) {
      sheet.getRange(firstKpiRow, column, 1, 1).setNumberFormat(column === 2 ? moneyFormat : '#,##0');
    });
    sheet.getRange(firstKpiRow, 4, 1, 1).setNumberFormat('#,##0');
    sheet.getRange(firstKpiRow, 8, 1, 1).setNumberFormat('0.00%');
    if (firstKpiRow + 1 <= lastRow) {
      sheet.setRowHeight(firstKpiRow + 1, 54);
      sheet.getRange(firstKpiRow, 1, 2, lastColumn).setWrap(true);
      sheet.getRange(firstKpiRow + 1, 2, 1, 1).setNumberFormat('#,##0.00');
      sheet.getRange(firstKpiRow + 1, 4, 1, 1).setNumberFormat('0.00%');
      sheet.getRange(firstKpiRow + 1, 6, 1, 1).setNumberFormat(moneyFormat);
      sheet.getRange(firstKpiRow + 1, 8, 1, 1).setNumberFormat(moneyFormat);
    }
  }
  if (sectionRows['CHANNEL SUMMARY'] && sectionRows['REVIEW FIRST']) {
    var channelStart = sectionRows['CHANNEL SUMMARY'] + 2;
    var channelCount = Math.max(0, sectionRows['REVIEW FIRST'] - channelStart - 1);
    if (channelCount > 0) {
      sheet.getRange(channelStart, 2, channelCount, 1).setNumberFormat(moneyFormat);
      sheet.getRange(channelStart, 3, channelCount, 2).setNumberFormat('#,##0');
      sheet.getRange(channelStart, 5, channelCount, 1).setNumberFormat('0.00%');
      sheet.getRange(channelStart, 6, channelCount, 1).setNumberFormat('#,##0.00');
      sheet.getRange(channelStart, 7, channelCount, 1).setNumberFormat('0.00%');
      sheet.getRange(channelStart, 8, channelCount, 1).setNumberFormat(moneyFormat);
    }
  }
  if (sectionRows['REVIEW FIRST'] && sectionRows['WORKBOOK DIRECTORY']) {
    var reviewStart = sectionRows['REVIEW FIRST'] + 2;
    var reviewCount = Math.max(0, sectionRows['WORKBOOK DIRECTORY'] - reviewStart - 1);
    if (reviewCount > 0) {
      sheet.getRange(reviewStart, 4, reviewCount, 2).setNumberFormat('@');
      sheet.getRange(reviewStart, 6, reviewCount, 1).setNumberFormat(moneyFormat);
      sheet.getRange(reviewStart, 7, reviewCount, 1).setNumberFormat('#,##0.00');
      sheet.getRange(reviewStart, 1, reviewCount, lastColumn).setWrap(true);
    }
  }
  if (sectionRows['WORKBOOK DIRECTORY']) {
    var directoryStart = sectionRows['WORKBOOK DIRECTORY'] + 2;
    var directoryCount = lastRow - directoryStart + 1;
    if (directoryCount > 0) {
      sheet.getRange(directoryStart, 1, directoryCount, lastColumn)
        .setWrap(true)
        .setVerticalAlignment('top');
    }
  }
  if (typeof sheet.setTabColor === 'function') sheet.setTabColor('#70AD47');
  return sheet;
}

function preferredTabOrder_(manifest) {
  var order = [
    START_HERE_SHEET_NAME,
    INFO_SHEET_NAME,
    DICTIONARY_SHEET_NAME,
    FIELD_DICTIONARY_SHEET_NAME
  ];
  (manifest || []).forEach(function(job) {
    if (order.indexOf(job.tab) < 0) order.push(job.tab);
  });
  return order;
}

function sheetIsSafelyRemovableBlank_(sheet) {
  if (!sheet || typeof sheet.getLastRow !== 'function' ||
      typeof sheet.getLastColumn !== 'function') return false;
  if (Number(sheet.getLastRow()) !== 0 || Number(sheet.getLastColumn()) !== 0) {
    return false;
  }
  var collectionMethods = [
    'getCharts', 'getImages', 'getDrawings', 'getSlicers', 'getPivotTables',
    'getDataSourcePivotTables', 'getDataSourceTables', 'getDataSourceFormulas',
    'getNamedRanges', 'getDeveloperMetadata', 'getBandings',
    'getConditionalFormatRules'
  ];
  for (var index = 0; index < collectionMethods.length; index++) {
    var method = collectionMethods[index];
    if (typeof sheet[method] !== 'function') continue;
    var values = sheet[method]();
    if (values && Number(values.length) > 0) return false;
  }
  if (typeof sheet.getFilter === 'function' && sheet.getFilter()) return false;
  if (typeof sheet.getFormUrl === 'function' && sheet.getFormUrl()) return false;
  if (typeof sheet.getFrozenRows === 'function' && Number(sheet.getFrozenRows()) > 0) {
    return false;
  }
  if (typeof sheet.getFrozenColumns === 'function' && Number(sheet.getFrozenColumns()) > 0) {
    return false;
  }
  return true;
}

function finalizeWorkbookLayout_(spreadsheet, manifest, removableBlankSheetNames) {
  var desired = preferredTabOrder_(manifest);
  var declared = {};
  desired.forEach(function(name) { declared[name] = true; });
  var removable = {};
  (removableBlankSheetNames || []).forEach(function(name) { removable[String(name)] = true; });

  spreadsheet.getSheets().slice().forEach(function(sheet) {
    var name = sheet.getName();
    if (removable[name] && !declared[name] &&
        sheetIsSafelyRemovableBlank_(sheet) &&
        spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(sheet);
    }
  });

  var position = 1;
  desired.forEach(function(name) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) return;
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(position);
    position++;
  });
  var infoSheet = spreadsheet.getSheetByName(INFO_SHEET_NAME);
  var startSheet = spreadsheet.getSheetByName(START_HERE_SHEET_NAME);
  if (!infoSheet) throw new Error('Missing ' + INFO_SHEET_NAME + ' during workbook finalization.');
  if (!startSheet) throw new Error('Missing ' + START_HERE_SHEET_NAME + ' during workbook finalization.');
  var groupColors = {
    performance: '#4472C4',
    structure: '#5B9BD5',
    creative: '#8064A2',
    targeting: '#ED7D31',
    negatives: '#C00000',
    audiences: '#BF9000',
    governance: '#7F8C8D'
  };
  Object.keys(START_HERE_TAB_GROUPS).forEach(function(group) {
    (START_HERE_TAB_GROUPS[group] || []).forEach(function(name) {
      var groupedSheet = spreadsheet.getSheetByName(name);
      if (groupedSheet && typeof groupedSheet.setTabColor === 'function') {
        groupedSheet.setTabColor(groupColors[group]);
      }
    });
  });
  [
    [START_HERE_SHEET_NAME, '#70AD47'],
    [INFO_SHEET_NAME, '#1F4E78'],
    [DICTIONARY_SHEET_NAME, '#2F75B5'],
    [FIELD_DICTIONARY_SHEET_NAME, '#5B9BD5']
  ].forEach(function(entry) {
    var metadataSheet = spreadsheet.getSheetByName(entry[0]);
    if (metadataSheet && typeof metadataSheet.setTabColor === 'function') {
      metadataSheet.setTabColor(entry[1]);
    }
  });
  spreadsheet.setActiveSheet(startSheet);
  return spreadsheet.getSheets().map(function(sheet) { return sheet.getName(); });
}

function safeSheetSuffix_(name) {
  return String(name).replace(/[\\/:?*\[\]]/g, '_').substring(0, 70);
}

function stageSheetName_(finalName) { return STAGE_PREFIX + safeSheetSuffix_(finalName); }
function backupSheetName_(finalName) { return BACKUP_PREFIX + safeSheetSuffix_(finalName); }

function isExpectedOptionalSourceUnavailableError_(error) {
  var text = String(error && error.message ? error.message : error || '');
  return /\b(?:UNRECOGNIZED_FIELD|PROHIBITED_FIELD_IN_SELECT_CLAUSE|PROHIBITED_RESOURCE_TYPE_IN_SELECT_CLAUSE|UNSUPPORTED_FIELD|UNSUPPORTED_RESOURCE)\b/i.test(text);
}

function expectedPartialSourceLimitation_(label, error) {
  if (!isExpectedOptionalSourceUnavailableError_(error)) throw error;
  return String(label || 'Optional source') + ': ' +
    String(error && error.message ? error.message : error).substring(0, 180);
}

function commitStagedSheet_(spreadsheet, stageName, finalName) {
  var stage = spreadsheet.getSheetByName(stageName);
  if (!stage) throw new Error('Missing staging sheet: ' + stageName);
  var finalSheet = spreadsheet.getSheetByName(finalName);
  var backupName = backupSheetName_(finalName);
  if (spreadsheet.getSheetByName(backupName)) throw new Error('Unresolved exporter backup exists: ' + backupName);
  var finalRenamed = false;
  var stageRenamed = false;
  try {
    if (finalSheet) { finalSheet.setName(backupName); finalRenamed = true; }
    stage.setName(finalName);
    stageRenamed = true;
    if (finalRenamed) spreadsheet.deleteSheet(finalSheet);
    return stage;
  } catch (error) {
    try {
      if (stageRenamed && stage.getName() === finalName) stage.setName(stageName);
      if (finalRenamed && finalSheet.getName() === backupName) finalSheet.setName(finalName);
    } catch (rollbackError) {
      throw new Error('Staged replacement failed (' + error + ') and rollback also failed (' + rollbackError + ').');
    }
    throw error;
  }
}

// --------------------------------------------------------------------------
// Resumable manifest engine
// --------------------------------------------------------------------------

function runManifestEngine_(
  state,
  manifest,
  adapter,
  minRemainingSeconds,
  minCommitRemainingSeconds,
  infoRefreshIntervalSeconds
) {
  validateManifest_(manifest);
  var minimum = Number(minRemainingSeconds);
  if (!isFinite(minimum) || minimum < 0) throw new Error('Invalid minimum remaining seconds.');
  var commitMinimum = minCommitRemainingSeconds === undefined ?
    minimum : Number(minCommitRemainingSeconds);
  if (!isFinite(commitMinimum) || commitMinimum < 0) {
    throw new Error('Invalid minimum commit remaining seconds.');
  }
  if (commitMinimum < minimum) {
    throw new Error('Minimum commit remaining seconds cannot be below the ordinary reserve.');
  }
  var infoIntervalSeconds = infoRefreshIntervalSeconds === undefined ?
    0 : Number(infoRefreshIntervalSeconds);
  if (!isFinite(infoIntervalSeconds) || infoIntervalSeconds < 0) {
    throw new Error('Invalid info refresh interval seconds.');
  }
  var infoIntervalMs = infoIntervalSeconds * 1000;
  var lastInfoWriteMs = null;
  function now_() { return adapter.nowMs ? Number(adapter.nowMs()) : new Date().getTime(); }
  function save_() { state.updatedAtMs = now_(); adapter.saveState(state); }
  function writeInfo_(force) {
    if (!adapter.writeInfo) return;
    var currentMs = now_();
    if (force || lastInfoWriteMs === null || infoIntervalMs === 0 ||
        currentMs - lastInfoWriteMs >= infoIntervalMs) {
      adapter.writeInfo(state);
      lastInfoWriteMs = currentMs;
    }
  }
  function writeProgressSummary_() {
    if (typeof adapter.writeProgressSummary === 'function') {
      adapter.writeProgressSummary(state, manifest);
    }
  }
  function reportTerminalWorkbookFailure_(failure, prefix) {
    state.status = 'COMPLETE_WITH_ERRORS';
    state.workbookError = (
      String(prefix || '') +
      String(failure && failure.message ? failure.message : failure)
    ).substring(0, 500);
    try {
      save_();
    } catch (checkpointSaveError) {
      state.workbookError = (
        state.workbookError + ' | Could not retain the checkpoint: ' +
        String(checkpointSaveError && checkpointSaveError.message ?
          checkpointSaveError.message : checkpointSaveError)
      ).substring(0, 500);
    }
    try {
      writeInfo_(true);
    } catch (failureInfoError) {
      state.workbookError = (
        state.workbookError + ' | Could not refresh _export_info: ' +
        String(failureInfoError && failureInfoError.message ?
          failureInfoError.message : failureInfoError)
      ).substring(0, 500);
    }
    if (typeof adapter.writeFailureSummary === 'function') {
      try {
        adapter.writeFailureSummary(state, manifest);
      } catch (failureSummaryError) {
        state.workbookError = (
          state.workbookError + ' | Could not refresh the failure summary: ' +
          String(failureSummaryError && failureSummaryError.message ?
            failureSummaryError.message : failureSummaryError)
        ).substring(0, 500);
        try { writeInfo_(true); } catch (ignoredFailureInfoRetry) {}
      }
    }
    return state;
  }
  function pause_() {
    state.status = 'PAUSED';
    save_();
    writeInfo_(true);
    writeProgressSummary_();
    return state;
  }
  function hasTime_() { return Number(adapter.remainingSeconds()) > minimum; }
  function hasCommitTime_() { return Number(adapter.remainingSeconds()) > commitMinimum; }
  state.status = 'RUNNING';
  writeInfo_(true);
  writeProgressSummary_();

  while (state.jobIndex < manifest.length) {
    if (!hasTime_()) return pause_();
    var job = manifest[state.jobIndex];
    var jobStartedAt = now_();
    state.currentJobId = job.id;
    state.tabs[job.tab] = normalizeTabResult_(state.tabs[job.tab] || {
      status: 'RUNNING', rows: 0, durationMs: 0, error: '', limitation: '',
      partialLimited: false, priorPreserved: false,
      sourceReadStartedAtMs: 0, sourceReadCompletedAtMs: 0
    });
    var priorDurationMs = Number(state.tabs[job.tab].durationMs || 0);
    var accumulatedDuration_ = function() {
      return priorDurationMs + Math.max(0, now_() - jobStartedAt);
    };
    try {
      if (!state.stageSheetName) {
        state.stageSheetName = adapter.startJob(job, state);
        state.jobPhase = 'WRITING';
        save_();
      }
      if (state.chunkInProgress) {
        adapter.rollbackChunk(job, state);
        state.chunkInProgress = false;
        save_();
      }
      var chunkCount = job.chunked ? Number(adapter.getChunkCount(job, state)) : 1;
      if (!isFinite(chunkCount) || chunkCount < 0 || Math.floor(chunkCount) !== chunkCount) {
        throw new Error('Invalid chunk count for job ' + job.id + ': ' + chunkCount);
      }
      while (state.chunkIndex < chunkCount) {
        if (!hasTime_()) return pause_();
        if (!state.tabs[job.tab].sourceReadStartedAtMs) {
          state.tabs[job.tab].sourceReadStartedAtMs = now_();
        }
        state.chunkStartRow = Number(adapter.getChunkStartRow(job, state));
        state.chunkInProgress = true;
        state.jobPhase = 'WRITING';
        save_();
        var chunkResult = adapter.runChunk(job, state, state.chunkIndex);
        state.tabs[job.tab].sourceReadCompletedAtMs = now_();
        var rowsWritten;
        if (chunkResult && typeof chunkResult === 'object') {
          rowsWritten = Number(chunkResult.rows || 0);
          if (chunkResult.status === 'LIMITED' || chunkResult.limitation) {
            state.tabs[job.tab].partialLimited = true;
            var limitation = String(chunkResult.limitation || 'Partial compatibility limitation');
            if (state.tabs[job.tab].limitation.indexOf(limitation) < 0) {
              state.tabs[job.tab].limitation +=
                (state.tabs[job.tab].limitation ? ' | ' : '') + limitation;
            }
          }
        } else {
          rowsWritten = Number(chunkResult || 0);
        }
        if (!isFinite(rowsWritten) || rowsWritten < 0) throw new Error('Invalid row count returned by ' + job.id + '.');
        state.tabs[job.tab].rows += rowsWritten;
        state.chunkInProgress = false;
        state.chunkIndex++;
        state.tabs[job.tab].durationMs = accumulatedDuration_();
        save_();
        writeInfo_();
      }
      if (!hasCommitTime_()) return pause_();
      state.jobPhase = 'COMMITTING';
      save_();
      adapter.commitJob(job, state);
      state.tabs[job.tab].status = state.tabs[job.tab].partialLimited ? 'LIMITED' : 'OK';
      state.tabs[job.tab].durationMs = accumulatedDuration_();
      state.tabs[job.tab].error = '';
    } catch (error) {
      if (state.tabs[job.tab].sourceReadStartedAtMs && state.chunkInProgress) {
        state.tabs[job.tab].sourceReadCompletedAtMs = now_();
      }
      var prior = adapter.hasPriorFinal ? Boolean(adapter.hasPriorFinal(job, state)) : false;
      var errorText = String(error && error.message ? error.message : error).substring(0, 500);
      var expectedOptionalSourceError = job.required === false &&
        typeof adapter.isExpectedOptionalSourceError === 'function' &&
        Boolean(adapter.isExpectedOptionalSourceError(job, error));
      var canCommitEmptyLimitedTab = job.required === false &&
        expectedOptionalSourceError &&
        Boolean(state.stageSheetName) && state.jobPhase === 'WRITING' &&
        Boolean(state.tabs[job.tab].sourceReadStartedAtMs);
      state.tabs[job.tab].status = prior ? 'ERROR_PREVIOUS_PRESERVED' : 'ERROR';
      state.tabs[job.tab].priorPreserved = prior;
      state.tabs[job.tab].rows = 0;
      state.tabs[job.tab].durationMs = accumulatedDuration_();
      state.tabs[job.tab].error = errorText;
      if (canCommitEmptyLimitedTab) {
        try {
          if (typeof adapter.commitEmptyLimitedJob !== 'function') {
            throw new Error('Runtime cannot commit the required header-only limitation tab.');
          }
          adapter.commitEmptyLimitedJob(job, state);
          state.tabs[job.tab].status = 'LIMITED';
          state.tabs[job.tab].priorPreserved = false;
          state.tabs[job.tab].limitation =
            'Optional source unavailable; this current-run tab contains headers only.';
        } catch (limitedCommitError) {
          state.tabs[job.tab].status = prior ? 'ERROR_PREVIOUS_PRESERVED' : 'ERROR';
          state.tabs[job.tab].priorPreserved = prior;
          state.tabs[job.tab].error += ' | header-only tab commit failed: ' +
            String(limitedCommitError && limitedCommitError.message ?
              limitedCommitError.message : limitedCommitError).substring(0, 300);
          try { if (adapter.abortJob) adapter.abortJob(job, state); }
          catch (limitedAbortError) {
            state.tabs[job.tab].error += ' | cleanup failed: ' + limitedAbortError;
          }
        }
      } else {
        try { if (adapter.abortJob) adapter.abortJob(job, state); }
        catch (abortError) { state.tabs[job.tab].error += ' | cleanup failed: ' + abortError; }
      }
    }
    state.jobIndex++;
    state.currentJobId = '';
    state.chunkIndex = 0;
    state.chunkInProgress = false;
    state.chunkStartRow = 0;
    state.stageSheetName = '';
    state.jobPhase = '';
    save_();
    writeInfo_();
  }
  var terminalFailure = manifest.some(function(job) {
    var tab = state.tabs[job.tab];
    return tab &&
      (tab.status === 'ERROR' || tab.status === 'ERROR_PREVIOUS_PRESERVED');
  });
  var limitedCoverage = manifest.some(function(job) {
    var tab = state.tabs[job.tab];
    return tab && tab.status === 'LIMITED';
  });
  if (!hasCommitTime_()) return pause_();
  if (terminalFailure) {
    return reportTerminalWorkbookFailure_(
      new Error('One or more export jobs failed. Review the ERROR rows.'),
      'Export job failure: '
    );
  }
  var terminalStatus = limitedCoverage ? 'COMPLETE_WITH_LIMITATIONS' : 'COMPLETE';
  state.status = 'FINALIZING';
  state.workbookError = '';
  save_();
  writeInfo_(true);
  try {
    if (adapter.finalizeWorkbook) adapter.finalizeWorkbook(state, manifest);
  } catch (finalizeError) {
    return reportTerminalWorkbookFailure_(
      finalizeError,
      'Workbook finalization failed: '
    );
  }
  if (!hasCommitTime_()) return pause_();
  try {
    adapter.clearState();
  } catch (clearStateError) {
    return reportTerminalWorkbookFailure_(
      clearStateError,
      'Checkpoint cleanup failed: '
    );
  }
  state.status = terminalStatus;
  try {
    if (typeof adapter.publishWorkbook !== 'function') {
      throw new Error('Native workbook publication is unavailable.');
    }
    adapter.publishWorkbook(state, manifest);
  } catch (publishError) {
    return reportTerminalWorkbookFailure_(
      publishError,
      'Native workbook publication failed: '
    );
  }
  return state;
}

// --------------------------------------------------------------------------
// Native Google Sheets terminal-workbook validation
// --------------------------------------------------------------------------

function normalizedCustomerId_(value) {
  return String(value === null || value === undefined ? '' : value).replace(/\D/g, '');
}

function exportInfoMetadataValue_(rows, key, statusHeaderIndex) {
  var stop = statusHeaderIndex === undefined ? rows.length : statusHeaderIndex;
  for (var rowIndex = 0; rowIndex < stop; rowIndex++) {
    var row = rows[rowIndex] || [];
    for (var columnIndex = 0; columnIndex < row.length - 1; columnIndex++) {
      if (String(row[columnIndex] || '') === String(key)) return row[columnIndex + 1];
    }
  }
  return '';
}

function parsedExportRange_(value) {
  var match = /^(\d{4}-\d{2}-\d{2})\s+through\s+(\d{4}-\d{2}-\d{2})$/.exec(
    String(value || '')
  );
  return match ? { start: match[1], end: match[2] } : {};
}

function parsedTimestampMs_(value) {
  var timestamp = Date.parse(String(value || ''));
  return isFinite(timestamp) ? timestamp : 0;
}

function nativeWorkbookSnapshotRuntime_(spreadsheet, spreadsheetId, manifest, expectedAccountId) {
  var infoSheet = spreadsheet.getSheetByName(INFO_SHEET_NAME);
  if (!infoSheet || infoSheet.getLastRow() < 1 || infoSheet.getLastColumn() < 1) {
    throw new Error('Native workbook validation requires a completed _export_info tab.');
  }
  var sheets = spreadsheet.getSheets();
  var snapshot = {
    spreadsheetId: String(spreadsheetId || ''),
    expectedAccountId: expectedAccountId,
    sheetNames: sheets.map(function(sheet) { return String(sheet.getName()); }),
    visibleSheetNames: sheets.filter(function(sheet) {
      return !(typeof sheet.isSheetHidden === 'function' && sheet.isSheetHidden());
    }).map(function(sheet) { return String(sheet.getName()); }),
    hiddenSheetNames: sheets.filter(function(sheet) {
      return typeof sheet.isSheetHidden === 'function' && sheet.isSheetHidden();
    }).map(function(sheet) { return String(sheet.getName()); }),
    infoRows: infoSheet.getRange(
      1, 1, infoSheet.getLastRow(), infoSheet.getLastColumn()
    ).getValues(),
    workbookGridCells: workbookGridCellCount_(spreadsheet),
    sheets: {}
  };
  manifest.forEach(function(job) {
    var sheet = spreadsheet.getSheetByName(job.tab);
    if (!sheet) return;
    var actualWidth = Number(sheet.getLastColumn());
    var sheetSnapshot = {
      headers: actualWidth < 1 ? [] :
        sheet.getRange(1, 1, 1, actualWidth).getValues()[0],
      rowCount: Math.max(0, Number(sheet.getLastRow()) - 1)
    };
    snapshot.sheets[job.tab] = sheetSnapshot;
  });
  return snapshot;
}

function assertNativeWorkbookSnapshot_(snapshot, manifest, options) {
  options = options || {};
  if (!snapshot || !Array.isArray(snapshot.infoRows) || !snapshot.infoRows.length) {
    throw new Error('Native workbook validation requires readable _export_info metadata.');
  }
  validateManifest_(manifest);
  var infoRows = snapshot.infoRows;
  if (String(infoRows[0][0] || '') !== OWNER_KEY || String(infoRows[0][1] || '') !== VERSION) {
    throw new Error('Workbook owner or exporter version does not match this script.');
  }
  var statusHeaderIndex = -1;
  for (var index = 0; index < infoRows.length; index++) {
    if (String((infoRows[index] || [])[0] || '') === 'tab') {
      statusHeaderIndex = index;
      break;
    }
  }
  if (statusHeaderIndex < 0) throw new Error('_export_info has no manifest status table.');
  var statusHeader = infoRows[statusHeaderIndex].map(String);
  var requiredColumns = [
    'tab', 'status', 'rows', 'duration_seconds', 'source_read_started_at',
    'source_read_completed_at', 'prior_data_preserved', 'limitation_or_error'
  ];
  if (stableStringify_(statusHeader) !== stableStringify_(requiredColumns)) {
    throw new Error('_export_info status columns do not match the current contract.');
  }

  var runId = String(exportInfoMetadataValue_(infoRows, 'run_id', statusHeaderIndex) || '');
  if (!runId || runId.indexOf('run-') !== 0) {
    throw new Error('Workbook has no completed export run ID.');
  }
  var overallStatus = String(
    exportInfoMetadataValue_(infoRows, 'overall_status', statusHeaderIndex) || ''
  );
  var finalizing = options.allowFinalizing === true && overallStatus === 'FINALIZING';
  if (overallStatus !== 'COMPLETE' && overallStatus !== 'COMPLETE_WITH_LIMITATIONS' &&
      !finalizing) {
    throw new Error('Native workbook is not terminal and usable: ' + overallStatus + '.');
  }
  var outputSchemaVersion = Number(
    exportInfoMetadataValue_(infoRows, 'output_schema_version', statusHeaderIndex)
  );
  var runtimeContractVersion = Number(
    exportInfoMetadataValue_(infoRows, 'runtime_contract_version', statusHeaderIndex)
  );
  if (outputSchemaVersion !== OUTPUT_SCHEMA_VERSION) {
    throw new Error('Workbook output schema version does not match this release.');
  }
  if (runtimeContractVersion !== RUNTIME_CONTRACT_VERSION) {
    throw new Error('Workbook runtime contract version does not match this release.');
  }
  var workbookAccountId = normalizedCustomerId_(
    exportInfoMetadataValue_(infoRows, 'account_id', statusHeaderIndex)
  );
  var expectedAccountId = normalizedCustomerId_(snapshot.expectedAccountId);
  if (!workbookAccountId || (expectedAccountId && workbookAccountId !== expectedAccountId)) {
    throw new Error('Workbook account does not match the current Google Ads account.');
  }
  var spreadsheetId = String(snapshot.spreadsheetId || '');
  if (!spreadsheetId) throw new Error('Workbook spreadsheet identity is missing.');

  var workbookStatus = String(
    exportInfoMetadataValue_(infoRows, 'workbook_status', statusHeaderIndex) || ''
  );
  var expectedWorkbookStatus = finalizing ? 'IN_PROGRESS' :
    (overallStatus === 'COMPLETE' ? 'READY' : 'READY_WITH_LIMITATIONS');
  if (workbookStatus !== expectedWorkbookStatus) {
    throw new Error('Native workbook status does not match its terminal export status.');
  }
  if (String(exportInfoMetadataValue_(infoRows, 'deliverable_type', statusHeaderIndex)) !==
      'NATIVE_GOOGLE_SHEET') {
    throw new Error('Workbook deliverable type does not match the native Sheet contract.');
  }
  if (String(exportInfoMetadataValue_(infoRows, 'workbook_url', statusHeaderIndex)) !==
      canonicalSpreadsheetUrl_(spreadsheetId)) {
    throw new Error('Workbook URL does not match the configured spreadsheet.');
  }
  if (String(exportInfoMetadataValue_(infoRows, 'reporting_window', statusHeaderIndex)) !==
      'LAST_90_COMPLETE_DAYS') {
    throw new Error('Workbook reporting-window metadata is invalid.');
  }
  var aggregateRange = parsedExportRange_(
    exportInfoMetadataValue_(infoRows, 'aggregate_range', statusHeaderIndex)
  );
  if (!aggregateRange.start || addDaysYmd_(aggregateRange.start, 89) !== aggregateRange.end) {
    throw new Error('Workbook aggregate range is not exactly 90 complete days.');
  }
  var weeklyRange = parsedExportRange_(
    exportInfoMetadataValue_(infoRows, 'weekly_range', statusHeaderIndex)
  );
  if (!weeklyRange.start || weeklyRange.start !== aggregateRange.start ||
      weeklyRange.end !== aggregateRange.end) {
    throw new Error('Workbook weekly range does not match the exact 90-day aggregate range.');
  }
  var changeRange = parsedExportRange_(
    exportInfoMetadataValue_(infoRows, 'change_range', statusHeaderIndex)
  );
  if (!changeRange.start || addDaysYmd_(changeRange.start, 27) !== changeRange.end ||
      changeRange.end !== aggregateRange.end) {
    throw new Error('Workbook change-history range is not the disclosed 28 complete days.');
  }
  if (String(exportInfoMetadataValue_(infoRows, 'last_complete_day', statusHeaderIndex)) !==
      aggregateRange.end) {
    throw new Error('Workbook last-complete-day metadata does not match the aggregate range.');
  }

  var expectedVisibleSheetNames = preferredTabOrder_(manifest);
  var actualVisibleSheetNames = (snapshot.visibleSheetNames || []).map(String);
  if (expectedVisibleSheetNames.length !== 41 ||
      stableStringify_(actualVisibleSheetNames) !== stableStringify_(expectedVisibleSheetNames)) {
    throw new Error('Workbook visible tab set or order does not match the 41-sheet contract.');
  }
  var hiddenSheetNames = (snapshot.hiddenSheetNames || []).map(String);
  var permittedHiddenNames = options.allowCheckpoint === true ? [STATE_SHEET_NAME] : [];
  var unexpectedHidden = hiddenSheetNames.filter(function(name) {
    return permittedHiddenNames.indexOf(name) < 0;
  });
  if (unexpectedHidden.length) {
    throw new Error('Workbook contains an unexpected hidden tab: ' + unexpectedHidden.join(', ') + '.');
  }
  if (options.allowCheckpoint === true && hiddenSheetNames.indexOf(STATE_SHEET_NAME) < 0) {
    throw new Error('Workbook recovery checkpoint is missing before native finalization.');
  }
  if (hiddenSheetNames.indexOf(STATE_SHEET_NAME) >= 0 && options.allowCheckpoint !== true) {
    throw new Error('Workbook checkpoint remains after finalization.');
  }
  var temporarySheets = (snapshot.sheetNames || []).filter(function(name) {
    return name.indexOf(STAGE_PREFIX) === 0 || name.indexOf(BACKUP_PREFIX) === 0;
  });
  if (temporarySheets.length) {
    throw new Error('Workbook contains temporary staging or backup tabs: ' + temporarySheets.join(', ') + '.');
  }
  var allowedNames = expectedVisibleSheetNames.concat(permittedHiddenNames);
  var undeclaredSheets = (snapshot.sheetNames || []).filter(function(name) {
    return allowedNames.indexOf(name) < 0;
  });
  if (undeclaredSheets.length) {
    throw new Error('Workbook contains undeclared tabs: ' + undeclaredSheets.join(', ') + '.');
  }

  var statusRows = infoRows.slice(statusHeaderIndex + 1).filter(function(row) {
    return row.some(function(value) { return String(value || '') !== ''; });
  });
  if (statusRows.length !== manifest.length) {
    throw new Error('_export_info manifest row count does not match the current manifest.');
  }
  var tabs = {};
  var hasLimitedJob = false;
  manifest.forEach(function(job, jobIndex) {
    var row = statusRows[jobIndex];
    if (String(row[0] || '') !== job.tab) {
      throw new Error('_export_info manifest order changed at ' + job.tab + '.');
    }
    var jobStatus = String(row[1] || '');
    if (['OK', 'LIMITED'].indexOf(jobStatus) < 0) {
      throw new Error('Tab ' + job.tab + ' is not terminal and usable: ' + jobStatus + '.');
    }
    if (jobStatus === 'LIMITED') hasLimitedJob = true;
    if (String(row[6] || '').toUpperCase() === 'YES') {
      throw new Error('Tab ' + job.tab + ' preserved prior data and is not a current native export.');
    }
    var declaredRows = row[2];
    if (typeof declaredRows !== 'number' || !isFinite(declaredRows) ||
        declaredRows < 0 || Math.floor(declaredRows) !== declaredRows) {
      throw new Error('Tab ' + job.tab + ' has an invalid typed integer declared row count.');
    }
    var physical = snapshot.sheets && snapshot.sheets[job.tab];
    if (!physical) throw new Error('Workbook is missing output tab ' + job.tab + '.');
    var expectedHeaders = headersForJob_(job).map(String);
    var actualHeaders = (physical.headers || []).map(String);
    if (stableStringify_(actualHeaders) !== stableStringify_(expectedHeaders)) {
      throw new Error('Tab ' + job.tab + ' headers do not match the current manifest.');
    }
    if (Number(physical.rowCount) !== declaredRows) {
      throw new Error('Tab ' + job.tab + ' physical rows do not match _export_info row count.');
    }
    var rawLimitationDetail = String(row[7] || '');
    var limitation = jobStatus === 'LIMITED' ? rawLimitationDetail : '';
    tabs[job.tab] = {
      status: jobStatus,
      rows: declaredRows,
      durationMs: Number(row[3] || 0) * 1000,
      sourceReadStartedAtMs: parsedTimestampMs_(row[4]),
      sourceReadCompletedAtMs: parsedTimestampMs_(row[5]),
      priorPreserved: false,
      limitation: limitation,
      error: ''
    };
  });
  if (!finalizing && ((hasLimitedJob && overallStatus !== 'COMPLETE_WITH_LIMITATIONS') ||
      (!hasLimitedJob && overallStatus !== 'COMPLETE'))) {
    throw new Error('Workbook overall status does not match its manifest job statuses.');
  }
  if (Number(snapshot.workbookGridCells || 0) > Number(CONFIG.WORKBOOK_CELL_SAFETY_LIMIT)) {
    throw new Error('Workbook grid usage exceeds the configured safety limit.');
  }
  return {
    runId: runId,
    status: overallStatus,
    workbookStatus: workbookStatus,
    accountId: workbookAccountId,
    spreadsheetId: spreadsheetId,
    outputSchemaVersion: outputSchemaVersion,
    runtimeContractVersion: runtimeContractVersion,
    visibleSheetCount: actualVisibleSheetNames.length,
    workbookGridCells: Number(snapshot.workbookGridCells || 0),
    ranges: { aggregate: aggregateRange },
    tabs: tabs
  };
}

function validateNativeWorkbookRuntime_(spreadsheet, spreadsheetId, manifest, expectedAccountId, options) {
  return assertNativeWorkbookSnapshot_(
    nativeWorkbookSnapshotRuntime_(spreadsheet, spreadsheetId, manifest, expectedAccountId),
    manifest,
    options
  );
}

// --------------------------------------------------------------------------
// Google Ads Scripts runtime adapters
// --------------------------------------------------------------------------

function extractSpreadsheetId_(urlOrId) {
  var value = String(urlOrId || '').trim();
  if (!value || value.indexOf('INSERT-GOOGLE-SHEETS') === 0) {
    throw new Error('Set CONFIG.SPREADSHEET_URL to a dedicated Google Sheets workbook.');
  }
  if (value.indexOf('https://') !== 0) {
    if (!/^[A-Za-z0-9_-]+$/.test(value)) throw new Error('Invalid spreadsheet ID.');
    return value;
  }
  var match = /\/spreadsheets\/d\/([A-Za-z0-9_-]+)/.exec(value);
  if (!match) throw new Error('Could not extract a spreadsheet ID from CONFIG.SPREADSHEET_URL.');
  return match[1];
}

function validateRuntimeSafetyConfig_(config) {
  var candidate = config || {};
  [
    'INCLUDE_SENSITIVE_CHANGE_DETAILS',
    'DIAGNOSTICS_LOG_SAMPLE_ROWS',
    'HIDE_RAW_MICROS_COLUMNS',
    'ALLOW_RESET'
  ].forEach(function(key) {
    if (typeof candidate[key] !== 'boolean') {
      throw new Error('CONFIG.' + key + ' must be true or false without quotes.');
    }
  });
  if (typeof candidate.FREEZE_CONTEXT_COLUMNS !== 'number') {
    throw new Error('CONFIG.FREEZE_CONTEXT_COLUMNS must be a number.');
  }
  var freezeColumns = candidate.FREEZE_CONTEXT_COLUMNS;
  if (!isFinite(freezeColumns) || freezeColumns < 0 ||
      Math.floor(freezeColumns) !== freezeColumns) {
    throw new Error('CONFIG.FREEZE_CONTEXT_COLUMNS must be a nonnegative integer.');
  }
  if (typeof candidate.MIN_REMAINING_SECONDS !== 'number' ||
      typeof candidate.MIN_COMMIT_REMAINING_SECONDS !== 'number') {
    throw new Error(
      'CONFIG.MIN_REMAINING_SECONDS and CONFIG.MIN_COMMIT_REMAINING_SECONDS ' +
      'must be numbers.'
    );
  }
  var ordinaryReserve = candidate.MIN_REMAINING_SECONDS;
  var commitReserve = candidate.MIN_COMMIT_REMAINING_SECONDS;
  if (!isFinite(ordinaryReserve) || ordinaryReserve <= 0 ||
      ordinaryReserve >= GOOGLE_ADS_MAX_EXECUTION_SECONDS) {
    throw new Error('CONFIG.MIN_REMAINING_SECONDS must be below the Google Ads execution limit.');
  }
  if (!isFinite(commitReserve) || commitReserve < ordinaryReserve ||
      commitReserve >= GOOGLE_ADS_MAX_EXECUTION_SECONDS) {
    throw new Error(
      'CONFIG.MIN_COMMIT_REMAINING_SECONDS must be at least MIN_REMAINING_SECONDS ' +
      'and below the Google Ads execution limit.'
    );
  }
  return true;
}

function validateRuntimeConfig_() {
  validateRuntimeSafetyConfig_(CONFIG);
  extractSpreadsheetId_(CONFIG.SPREADSHEET_URL);
  if (!/^v\d+$/.test(String(CONFIG.API_VERSION || ''))) {
    throw new Error('CONFIG.API_VERSION must look like v25.');
  }
  [
    ['BATCH_ROWS', CONFIG.BATCH_ROWS],
    ['CAMPAIGN_CHUNK_SIZE', CONFIG.CAMPAIGN_CHUNK_SIZE],
    ['MIN_REMAINING_SECONDS', CONFIG.MIN_REMAINING_SECONDS],
    ['MIN_COMMIT_REMAINING_SECONDS', CONFIG.MIN_COMMIT_REMAINING_SECONDS],
    ['INFO_REFRESH_INTERVAL_SECONDS', CONFIG.INFO_REFRESH_INTERVAL_SECONDS],
    ['MAX_RESUME_AGE_HOURS', CONFIG.MAX_RESUME_AGE_HOURS],
    ['WORKBOOK_CELL_SAFETY_LIMIT', CONFIG.WORKBOOK_CELL_SAFETY_LIMIT]
  ].forEach(function(entry) {
    if (!isFinite(Number(entry[1])) || Number(entry[1]) <= 0) {
      throw new Error('CONFIG.' + entry[0] + ' must be a positive number.');
    }
  });
  if (Math.floor(Number(CONFIG.BATCH_ROWS)) !== Number(CONFIG.BATCH_ROWS) ||
      Math.floor(Number(CONFIG.CAMPAIGN_CHUNK_SIZE)) !== Number(CONFIG.CAMPAIGN_CHUNK_SIZE) ||
      Math.floor(Number(CONFIG.INFO_REFRESH_INTERVAL_SECONDS)) !==
        Number(CONFIG.INFO_REFRESH_INTERVAL_SECONDS)) {
    throw new Error(
      'CONFIG.BATCH_ROWS, CONFIG.CAMPAIGN_CHUNK_SIZE, and ' +
      'CONFIG.INFO_REFRESH_INTERVAL_SECONDS must be integers.'
    );
  }
  if (!Array.isArray(CONFIG.CHANGE_HISTORY_CLIENT_TYPES)) {
    throw new Error('CONFIG.CHANGE_HISTORY_CLIENT_TYPES must be an array (or an empty array for all client types).');
  }
  CONFIG.CHANGE_HISTORY_CLIENT_TYPES.forEach(function(clientType) {
    if (!/^[A-Z][A-Z0-9_]*$/.test(String(clientType || ''))) {
      throw new Error('Invalid CONFIG.CHANGE_HISTORY_CLIENT_TYPES value: ' + clientType);
    }
  });
}

function materialConfigSignature_(manifest) {
  return stableStringify_({
    apiVersion: String(CONFIG.API_VERSION),
    campaignChunkSize: Number(CONFIG.CAMPAIGN_CHUNK_SIZE),
    changeHistoryClientTypes: (CONFIG.CHANGE_HISTORY_CLIENT_TYPES || []).slice(),
    freezeContextColumns: Number(CONFIG.FREEZE_CONTEXT_COLUMNS),
    hideRawMicrosColumns: Boolean(CONFIG.HIDE_RAW_MICROS_COLUMNS),
    includeSensitiveChangeDetails: Boolean(CONFIG.INCLUDE_SENSITIVE_CHANGE_DETAILS),
    manifestExecutionSignature: manifestExecutionSignature_(
      manifest || getManifestDefinition_(),
      CONFIG.CAMPAIGN_CHUNK_SIZE
    ),
    outputSchemaVersion: OUTPUT_SCHEMA_VERSION,
    runtimeContractVersion: RUNTIME_CONTRACT_VERSION,
    workbookCellSafetyLimit: Number(CONFIG.WORKBOOK_CELL_SAFETY_LIMIT),
    version: VERSION
  });
}

function openSpreadsheetRuntime_(spreadsheetId) {
  return withRetry_(function() { return SpreadsheetApp.openById(spreadsheetId); }, { retries: 6 });
}

function reportRuntime_(query, reportOptions) {
  var options = { apiVersion: CONFIG.API_VERSION };
  if (reportOptions && Object.prototype.hasOwnProperty.call(reportOptions, 'resolveGeoNames')) {
    if (typeof reportOptions.resolveGeoNames !== 'boolean') {
      throw new Error('resolveGeoNames must be a boolean when supplied.');
    }
    options.resolveGeoNames = reportOptions.resolveGeoNames;
  }
  return AdsApp.report(query, options);
}

function getOrInsertSheetRuntime_(spreadsheet, name) {
  return withRetry_(function() {
    return spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  }, { retries: 6 });
}

function summarizeWorkbookRuntime_(spreadsheet) {
  return spreadsheet.getSheets().map(function(sheet) {
    var blank = sheetIsSafelyRemovableBlank_(sheet);
    var marker = blank ? '' : String(sheet.getRange(1, 1).getDisplayValue() || '');
    return { name: sheet.getName(), blank: blank, marker: marker };
  });
}

function writeMatrixRuntime_(sheet, rows) {
  if (!rows || !rows.length) return;
  var width = rows[0].length;
  var safeRows = rows.map(function(row) { return encodeSheetRowForWrite_(row, width); });
  sheet.clearContents();
  ensureSheetCapacity_(sheet, safeRows.length, width, CONFIG.WORKBOOK_CELL_SAFETY_LIMIT);
  withRetry_(function() {
    sheet.getRange(1, 1, safeRows.length, width).setValues(safeRows);
  }, { retries: 5 });
  try { sheet.setFrozenRows(1); } catch (ignored) {}
  SpreadsheetApp.flush();
}

function writeMatrixPreservingOwnerRuntime_(sheet, rows) {
  if (!rows || !rows.length) return;
  if (String(sheet.getRange(1, 1).getDisplayValue() || '') !== OWNER_KEY) {
    throw new Error('Owner-preserving metadata update requires an owned existing sheet.');
  }
  var width = rows[0].length;
  var safeRows = rows.map(function(row) { return encodeSheetRowForWrite_(row, width); });
  var previousLastRow = Math.max(1, Number(sheet.getLastRow()));
  var previousLastColumn = Math.max(1, Number(sheet.getLastColumn()));
  ensureSheetCapacity_(sheet, safeRows.length, width, CONFIG.WORKBOOK_CELL_SAFETY_LIMIT);
  withRetry_(function() {
    sheet.getRange(1, 1, safeRows.length, width).setValues(safeRows);
  }, { retries: 5 });
  if (previousLastRow > safeRows.length) {
    sheet.getRange(
      safeRows.length + 1,
      1,
      previousLastRow - safeRows.length,
      previousLastColumn
    ).clearContent();
  }
  if (previousLastColumn > width) {
    sheet.getRange(
      1,
      width + 1,
      Math.max(previousLastRow, safeRows.length),
      previousLastColumn - width
    ).clearContent();
  }
  try { sheet.setFrozenRows(1); } catch (ignored) {}
  SpreadsheetApp.flush();
}

function writeExportInfoRuntime_(spreadsheet, state, manifest, options) {
  var sheet = getOrInsertSheetRuntime_(spreadsheet, INFO_SHEET_NAME);
  var excludedNames = options && options.excludeCheckpointFromGrid ?
    [STATE_SHEET_NAME] : [];
  var existingOwnerMarker = String(
    sheet.getRange(1, 1).getDisplayValue() || ''
  ) === OWNER_KEY;
  var matrixWriter = (options && options.preserveOwnerMarker) || existingOwnerMarker ?
    writeMatrixPreservingOwnerRuntime_ : writeMatrixRuntime_;
  state.workbookGridCells = workbookGridCellCountExcluding_(spreadsheet, excludedNames);
  state.workbookCellSafetyLimit = Number(CONFIG.WORKBOOK_CELL_SAFETY_LIMIT);
  matrixWriter(sheet, buildExportInfoRows_(state, manifest));
  trimSheetGrid_(sheet);
  formatExportInfoSheet_(sheet);
  var finalGridCells = workbookGridCellCountExcluding_(spreadsheet, excludedNames);
  if (finalGridCells !== state.workbookGridCells) {
    state.workbookGridCells = finalGridCells;
    matrixWriter(sheet, buildExportInfoRows_(state, manifest));
    trimSheetGrid_(sheet);
    formatExportInfoSheet_(sheet);
  }
}

function sourceRowForCampaignRuntime_(sheet, campaignId) {
  if (!sheet || !campaignId || Number(sheet.getLastRow()) < 2 ||
      Number(sheet.getLastColumn()) < 1) return 1;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
  var campaignColumn = headers.indexOf('campaign.id');
  if (campaignColumn < 0) return 1;
  var values = sheet.getRange(2, campaignColumn + 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues();
  for (var index = 0; index < values.length; index++) {
    if (String(values[index][0] || '') === String(campaignId)) return index + 2;
  }
  return 1;
}

function applyStartHereNavigationLinks_(spreadsheet, sheet) {
  if (!spreadsheet || !sheet || typeof spreadsheet.getId !== 'function' ||
      typeof spreadsheet.getSheetByName !== 'function' ||
      typeof SpreadsheetApp === 'undefined' ||
      typeof SpreadsheetApp.newRichTextValue !== 'function') return 0;
  var lastRow = Number(sheet.getLastRow());
  var lastColumn = Math.min(8, Number(sheet.getLastColumn()));
  if (lastRow < 1 || lastColumn < 1) return 0;
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var reviewSection = -1;
  var directorySection = -1;
  values.forEach(function(row, index) {
    if (String(row[0] || '') === 'REVIEW FIRST') reviewSection = index;
    if (String(row[0] || '') === 'WORKBOOK DIRECTORY') directorySection = index;
  });
  var baseUrl = canonicalSpreadsheetUrl_(spreadsheet.getId());
  if (!baseUrl) return 0;
  var linked = 0;
  function linkCell_(rowIndex, columnIndex, tab, rangeA1) {
    var target = spreadsheet.getSheetByName(String(tab || ''));
    if (!target || typeof target.getSheetId !== 'function') return;
    var text = String((values[rowIndex] || [])[columnIndex] || '');
    if (!text) return;
    var url = baseUrl + '#gid=' + target.getSheetId() + '&range=' + String(rangeA1 || 'A1');
    var builder = SpreadsheetApp.newRichTextValue().setText(text);
    if (typeof builder.setLinkUrl !== 'function') return;
    var range = sheet.getRange(rowIndex + 1, columnIndex + 1, 1, 1);
    range.setRichTextValues([[builder.setLinkUrl(url).build()]]);
    if (typeof range.setShowHyperlink === 'function') range.setShowHyperlink(true);
    linked++;
  }
  if (reviewSection >= 0 && directorySection > reviewSection) {
    for (var reviewIndex = reviewSection + 2; reviewIndex < directorySection; reviewIndex++) {
      var reviewTab = String((values[reviewIndex] || [])[2] || '');
      if (!reviewTab) continue;
      linkCell_(reviewIndex, 2, reviewTab, 'A1');
      var campaignId = String((values[reviewIndex] || [])[3] || '');
      if (campaignId) {
        var targetSheet = spreadsheet.getSheetByName(reviewTab);
        var sourceRow = sourceRowForCampaignRuntime_(targetSheet, campaignId);
        linkCell_(reviewIndex, 4, reviewTab, 'A' + sourceRow);
      }
    }
  }
  if (directorySection >= 0) {
    for (var directoryIndex = directorySection + 2;
         directoryIndex < values.length; directoryIndex++) {
      var directoryTab = String((values[directoryIndex] || [])[1] || '');
      if (directoryTab) linkCell_(directoryIndex, 1, directoryTab, 'A1');
    }
  }
  return linked;
}

function writeStartHereRuntime_(spreadsheet, state, manifest, options) {
  var campaignRows = options && options.progressOnly ? [] :
    (readSheetObjectsRuntime_(spreadsheet, 'campaign') || []);
  var model = buildStartHereModel_(state, campaignRows, manifest);
  var rows = buildStartHereRows_(model);
  var sheet = getOrInsertSheetRuntime_(spreadsheet, START_HERE_SHEET_NAME);
  if (options && options.progressOnly) {
    ensureSheetCapacity_(sheet, 3, 4, CONFIG.WORKBOOK_CELL_SAFETY_LIMIT);
    sheet.getRange(3, 4, 1, 1).setValues([[
      encodeSheetCellForWrite_('Data status: ' + String(state.status || 'RUNNING'))
    ]]);
    SpreadsheetApp.flush();
  }
  writeMatrixRuntime_(sheet, rows);
  trimSheetGrid_(sheet);
  formatStartHereSheet_(sheet, state.accountCurrencyCode || '');
  applyStartHereNavigationLinks_(spreadsheet, sheet);
  if (typeof SpreadsheetApp !== 'undefined' && typeof SpreadsheetApp.flush === 'function') {
    SpreadsheetApp.flush();
  }
  return model;
}

function stateJsonForSheet_(state) {
  return JSON.stringify(compactStateForStorage_(state, 45000));
}

function saveStateSheetRuntime_(spreadsheet, state) {
  var sheet = spreadsheet.getSheetByName(STATE_SHEET_NAME);
  if (!sheet || String(sheet.getRange(1, 1).getDisplayValue()) !== OWNER_KEY) {
    throw new Error('Exporter state sheet is missing or not owned by this exporter.');
  }
  var json = stateJsonForSheet_(state);
  ensureSheetCapacity_(sheet, 4, 2, CONFIG.WORKBOOK_CELL_SAFETY_LIMIT);
  try { sheet.getRange(4, 2).setNumberFormat('@'); } catch (ignoredFormat) {}
  withRetry_(function() {
    sheet.getRange(4, 2, 1, 1).setValues([[encodeSheetCellForWrite_(json)]]);
  }, { retries: 5 });
  SpreadsheetApp.flush();
}

function loadStateSheetRuntime_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName(STATE_SHEET_NAME);
  if (!sheet) return null;
  if (String(sheet.getRange(1, 1).getDisplayValue()) !== OWNER_KEY ||
      String(sheet.getRange(4, 1).getDisplayValue()) !== 'state_json') {
    throw new Error('Saved export state sheet is unreadable. Set ALLOW_RESET=true and run resetExportState().');
  }
  var raw = String(sheet.getRange(4, 2).getDisplayValue() || '');
  if (!raw) throw new Error('Saved export state JSON is missing. Run resetExportState().');
  try {
    var state = JSON.parse(raw);
    if (String(sheet.getRange(2, 2).getDisplayValue()) !== String(state.runId || '')) {
      throw new Error('run ID mismatch');
    }
    return state;
  } catch (error) {
    throw new Error('Saved export state is unreadable. Set ALLOW_RESET=true and run resetExportState().');
  }
}

function fetchCampaignIdsRuntime_(range) {
  var queries = buildCampaignEligibilityQueries_(range);
  return collectEligibleCampaignIds_(
    collectReportRowsRuntime_(queries.current),
    collectReportRowsRuntime_(queries.activity)
  );
}

function ensureStateSheetHiddenRuntime_(sheet) {
  if (!sheet || typeof sheet.isSheetHidden !== 'function' ||
      typeof sheet.hideSheet !== 'function') {
    throw new Error('Exporter checkpoint visibility cannot be verified. Reset the export state.');
  }
  if (sheet.isSheetHidden()) return true;
  sheet.hideSheet();
  SpreadsheetApp.flush();
  if (!sheet.isSheetHidden()) {
    throw new Error('Exporter checkpoint could not be hidden. Run main() again to retry.');
  }
  return true;
}

function writeCampaignStateRuntime_(spreadsheet, state, campaignIds) {
  var sheet = getOrInsertSheetRuntime_(spreadsheet, STATE_SHEET_NAME);
  writeMatrixRuntime_(sheet, [
    [OWNER_KEY, VERSION],
    ['run_id', state.runId],
    ['campaign_id_count', campaignIds.length],
    ['state_json', stateJsonForSheet_(state)]
  ]);
  try {
    sheet.getRange(4, 1, 1, 2).setNumberFormat('@');
    sheet.getRange(5, 1, Math.max(1, sheet.getMaxRows() - 4), 1).setNumberFormat('@');
  } catch (ignoredFormat) {}
  var writer = createSafeRowBuffer_(sheet, ['campaign_id'], {
    batchRows: CONFIG.BATCH_ROWS,
    cellLimit: CONFIG.WORKBOOK_CELL_SAFETY_LIMIT,
    retries: 5
  });
  campaignIds.forEach(function(id) { writer.push([id]); });
  writer.flush();
  trimSheetGrid_(sheet);
  ensureStateSheetHiddenRuntime_(sheet);
  SpreadsheetApp.flush();
}

function readCampaignStateRuntime_(spreadsheet, state) {
  var sheet = spreadsheet.getSheetByName(STATE_SHEET_NAME);
  if (!sheet) throw new Error('Resume campaign checkpoint sheet is missing. Reset the export state.');
  if (String(sheet.getRange(1, 1).getDisplayValue()) !== OWNER_KEY ||
      String(sheet.getRange(2, 2).getDisplayValue()) !== state.runId) {
    throw new Error('Resume campaign checkpoint does not match the saved run. Reset the export state.');
  }
  ensureStateSheetHiddenRuntime_(sheet);
  var expected = sheet.getRange(3, 2).getValue();
  if (typeof expected !== 'number' || !isFinite(expected) || expected < 0 ||
      Math.floor(expected) !== expected) {
    throw new Error('Resume campaign checkpoint count is not a typed integer. Reset the export state.');
  }
  var ids = [];
  var lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, 1).getDisplayValues().forEach(function(row) {
      var id = String(row[0] || '').replace(/^'/, '');
      if (id) ids.push(id);
    });
  }
  var chunks = chunkCampaignIds_(ids, Math.max(1, ids.length || 1));
  var sorted = chunks.length ? chunks[0] : [];
  if (sorted.length !== expected) throw new Error('Resume campaign checkpoint row count changed. Reset the export state.');
  return sorted;
}

function prepareStageRuntime_(spreadsheet, job) {
  var name = stageSheetName_(job.tab);
  var sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  writeMatrixRuntime_(sheet, [headersForJob_(job)]);
  trimSheetGrid_(sheet);
  textColumnIndexes_(job).forEach(function(columnIndex) {
    try { sheet.getRange(1, columnIndex, sheet.getMaxRows(), 1).setNumberFormat('@'); }
    catch (ignoredFormat) {}
  });
  return name;
}

function commitRuntimeJob_(spreadsheet, job, state, emptyNote) {
  var stageName = state.stageSheetName || stageSheetName_(job.tab);
  var backupName = backupSheetName_(job.tab);
  var stage = spreadsheet.getSheetByName(stageName);
  var finalSheet = spreadsheet.getSheetByName(job.tab);
  var backup = spreadsheet.getSheetByName(backupName);

  if (backup && stage && !finalSheet) {
    backup.setName(job.tab);
    finalSheet = backup;
    backup = null;
  } else if (backup && !stage && finalSheet) {
    spreadsheet.deleteSheet(backup);
    return finalSheet;
  } else if (backup && !stage && !finalSheet) {
    backup.setName(job.tab);
    throw new Error('Prior tab was restored, but the completed staging tab is missing.');
  } else if (backup) {
    throw new Error('Ambiguous staged replacement state for tab ' + job.tab + '.');
  }

  if (!stage) {
    if (finalSheet && state.jobPhase === 'COMMITTING') return finalSheet;
    throw new Error('Missing staging sheet for ' + job.tab + '.');
  }
  trimSheetGrid_(stage);
  formatReportSheet_(stage, job, state.accountCurrencyCode || '', emptyNote);
  return commitStagedSheet_(spreadsheet, stageName, job.tab);
}

function commitEmptyLimitedRuntimeJob_(spreadsheet, job, state) {
  var stageName = state.stageSheetName || stageSheetName_(job.tab);
  var stage = spreadsheet.getSheetByName(stageName);
  if (!stage) {
    throw new Error('Missing staging sheet for header-only limitation tab ' + job.tab + '.');
  }
  writeMatrixRuntime_(stage, [headersForJob_(job)]);
  trimSheetGrid_(stage);
  state.jobPhase = 'COMMITTING';
  return commitRuntimeJob_(
    spreadsheet,
    job,
    state,
    'Export completed with limited coverage. Review this tab\'s LIMITED row in _export_info.'
  );
}

function abortRuntimeJob_(spreadsheet, job, state) {
  var stageName = state.stageSheetName || stageSheetName_(job.tab);
  var backupName = backupSheetName_(job.tab);
  var stage = spreadsheet.getSheetByName(stageName);
  var backup = spreadsheet.getSheetByName(backupName);
  var finalSheet = spreadsheet.getSheetByName(job.tab);
  if (backup) {
    if (stage) { spreadsheet.deleteSheet(stage); stage = null; }
    if (finalSheet) {
      finalSheet.setName(stageName);
      backup.setName(job.tab);
      spreadsheet.deleteSheet(finalSheet);
    } else {
      backup.setName(job.tab);
    }
  } else if (stage) {
    spreadsheet.deleteSheet(stage);
  }
}

function campaignChunksForJobRuntime_(context, job) {
  context.campaignChunksByJob = context.campaignChunksByJob || {};
  if (!Object.prototype.hasOwnProperty.call(context.campaignChunksByJob, job.id)) {
    context.campaignChunksByJob[job.id] = chunkCampaignIdsForJob_(
      job,
      context.campaignIds,
      CONFIG.CAMPAIGN_CHUNK_SIZE
    );
  }
  return context.campaignChunksByJob[job.id];
}

function createRuntimeAdapter_(context) {
  return {
    remainingSeconds: function() {
      return AdsApp.getExecutionInfo().getRemainingTime();
    },
    saveState: function(state) { saveStateSheetRuntime_(context.spreadsheet, state); },
    writeInfo: function(state) { writeExportInfoRuntime_(context.spreadsheet, state, context.manifest); },
    writeProgressSummary: function(state, manifest) {
      writeStartHereRuntime_(
        context.spreadsheet,
        state,
        manifest,
        { progressOnly: true }
      );
    },
    startJob: function(job) { return prepareStageRuntime_(context.spreadsheet, job); },
    getChunkCount: function(job) {
      return job.chunked ? campaignChunksForJobRuntime_(context, job).length : 1;
    },
    getChunkStartRow: function(job, state) {
      var sheet = context.spreadsheet.getSheetByName(state.stageSheetName);
      if (!sheet) throw new Error('Missing staging sheet for ' + job.tab + '.');
      return sheet.getLastRow() + 1;
    },
    rollbackChunk: function(job, state) {
      var sheet = context.spreadsheet.getSheetByName(state.stageSheetName);
      if (!sheet) throw new Error('Cannot roll back missing staging sheet for ' + job.tab + '.');
      rollbackPartialChunk_(sheet, state.chunkStartRow);
      SpreadsheetApp.flush();
    },
    runChunk: function(job, state, chunkIndex) {
      return runJobChunkRuntime_(context, job, state, chunkIndex);
    },
    commitJob: function(job, state) { return commitRuntimeJob_(context.spreadsheet, job, state); },
    commitEmptyLimitedJob: function(job, state) {
      return commitEmptyLimitedRuntimeJob_(context.spreadsheet, job, state);
    },
    isExpectedOptionalSourceError: function(_job, error) {
      return isExpectedOptionalSourceUnavailableError_(error);
    },
    abortJob: function(job, state) { return abortRuntimeJob_(context.spreadsheet, job, state); },
    hasPriorFinal: function(job) {
      return Boolean(context.spreadsheet.getSheetByName(job.tab) ||
        context.spreadsheet.getSheetByName(backupSheetName_(job.tab)));
    },
    finalizeWorkbook: function(state, manifest) {
      writeStartHereRuntime_(context.spreadsheet, state, manifest);
      finalizeWorkbookLayout_(
        context.spreadsheet,
        manifest,
        state.initialBlankSheetNames || []
      );
      writeExportInfoRuntime_(
        context.spreadsheet,
        state,
        manifest,
        { excludeCheckpointFromGrid: true }
      );
      var validation = validateNativeWorkbookRuntime_(
        context.spreadsheet,
        context.spreadsheetId,
        manifest,
        state.accountId,
        { allowCheckpoint: true, allowFinalizing: true }
      );
      SpreadsheetApp.flush();
      return validation;
    },
    publishWorkbook: function(state, manifest) {
      validateNativeWorkbookRuntime_(
        context.spreadsheet,
        context.spreadsheetId,
        manifest,
        state.accountId,
        { allowFinalizing: true }
      );
      writeStartHereRuntime_(context.spreadsheet, state, manifest);
      writeExportInfoRuntime_(
        context.spreadsheet,
        state,
        manifest,
        { preserveOwnerMarker: true }
      );
      SpreadsheetApp.flush();
      return validateNativeWorkbookRuntime_(
        context.spreadsheet,
        context.spreadsheetId,
        manifest,
        state.accountId
      );
    },
    writeFailureSummary: function(state, manifest) {
      writeStartHereRuntime_(context.spreadsheet, state, manifest);
      writeExportInfoRuntime_(context.spreadsheet, state, manifest);
      SpreadsheetApp.flush();
    },
    clearState: function() {
      var stateSheet = context.spreadsheet.getSheetByName(STATE_SHEET_NAME);
      if (stateSheet) context.spreadsheet.deleteSheet(stateSheet);
      SpreadsheetApp.flush();
    },
    nowMs: function() { return new Date().getTime(); }
  };
}

function assertAdvertiserAccountRuntime_() {
  var iterator = reportRuntime_(
    'SELECT\n  customer.manager\nFROM customer'
  ).rows();
  if (!iterator.hasNext()) {
    throw new Error('Could not determine whether the current Google Ads account is a manager account.');
  }
  var value = iterator.next()['customer.manager'];
  if (value === true || String(value).toLowerCase() === 'true') {
    throw new Error('This version runs only in an individual advertiser account, not a manager account.');
  }
  return true;
}

function exportInfoRowsRuntime_(spreadsheet) {
  var infoSheet = spreadsheet.getSheetByName(INFO_SHEET_NAME);
  if (!infoSheet || infoSheet.getLastRow() < 2 || infoSheet.getLastColumn() < 2) return null;
  return infoSheet.getRange(
    1, 1, infoSheet.getLastRow(), infoSheet.getLastColumn()
  ).getValues();
}

function assertOwnedWorkbookRestartIdentityRuntime_(spreadsheet, manifest, expectedAccountId) {
  var rows = exportInfoRowsRuntime_(spreadsheet);
  if (!rows || String((rows[0] || [])[0] || '') !== OWNER_KEY ||
      String((rows[0] || [])[1] || '') !== VERSION) {
    throw new Error('Owned workbook restart metadata is missing or uses another exporter version.');
  }
  var statusHeaderIndex = rows.length;
  for (var index = 0; index < rows.length; index++) {
    if (String((rows[index] || [])[0] || '') === 'tab') {
      statusHeaderIndex = index;
      break;
    }
  }
  var status = String(
    exportInfoMetadataValue_(rows, 'overall_status', statusHeaderIndex) || ''
  );
  if (['COMPLETE', 'COMPLETE_WITH_LIMITATIONS', 'RESET'].indexOf(status) < 0) {
    throw new Error(
      'Owned workbook has nonterminal metadata but its recovery checkpoint is missing. ' +
      'Run resetExportState() before starting a replacement export.'
    );
  }
  var workbookAccountId = normalizedCustomerId_(
    exportInfoMetadataValue_(rows, 'account_id', statusHeaderIndex)
  );
  var currentAccountId = normalizedCustomerId_(expectedAccountId);
  if ((!workbookAccountId && status !== 'RESET') ||
      (workbookAccountId && currentAccountId && workbookAccountId !== currentAccountId)) {
    throw new Error(
      'Owned workbook account does not match the current Google Ads advertiser account.'
    );
  }
  var declared = {};
  preferredTabOrder_(manifest).forEach(function(name) { declared[name] = true; });
  spreadsheet.getSheets().forEach(function(sheet) {
    var name = String(sheet.getName());
    var blank = sheetIsSafelyRemovableBlank_(sheet);
    if (!declared[name] && !blank) {
      throw new Error('Owned workbook contains an undeclared tab before restart: ' + name + '.');
    }
    if (typeof sheet.isSheetHidden === 'function' && sheet.isSheetHidden()) {
      throw new Error('Owned workbook contains a hidden tab before restart: ' + name + '.');
    }
  });
  return true;
}

function runExport_() {
  validateRuntimeConfig_();
  var account = AdsApp.currentAccount();
  var spreadsheetId = extractSpreadsheetId_(CONFIG.SPREADSHEET_URL);
  var spreadsheet = openSpreadsheetRuntime_(spreadsheetId);
  var workbookSummary = summarizeWorkbookRuntime_(spreadsheet);
  var ownership = assertWorkbookOwnership_(workbookSummary);
  var manifest = getManifestDefinition_();
  validateManifest_(manifest);
  var nowMs = new Date().getTime();
  var identity = {
    version: VERSION,
    outputSchemaVersion: OUTPUT_SCHEMA_VERSION,
    runtimeContractVersion: RUNTIME_CONTRACT_VERSION,
    accountId: account.getCustomerId(),
    spreadsheetId: spreadsheetId,
    configSignature: materialConfigSignature_(manifest)
  };
  var state = loadStateSheetRuntime_(spreadsheet);
  if (!state && ownership === 'owned') {
    assertOwnedWorkbookRestartIdentityRuntime_(
      spreadsheet,
      manifest,
      account.getCustomerId()
    );
  }
  if (state) {
    assertStateCompatible_(state, identity, nowMs, CONFIG.MAX_RESUME_AGE_HOURS);
    var currentManifestIds = manifest.map(function(job) { return job.id; });
    if (stableStringify_(state.manifest) !== stableStringify_(currentManifestIds)) {
      throw new Error('Saved export state has a different job manifest. Reset before continuing.');
    }
  }

  assertAdvertiserAccountRuntime_();
  var campaignIds;
  if (state) {
    campaignIds = readCampaignStateRuntime_(spreadsheet, state);
  } else {
    var orphanNames = spreadsheet.getSheets().map(function(sheet) {
      return sheet.getName();
    }).filter(function(name) {
      return name === STATE_SHEET_NAME || name.indexOf(STAGE_PREFIX) === 0 ||
        name.indexOf(BACKUP_PREFIX) === 0;
    });
    if (orphanNames.length) {
      throw new Error(
        'Exporter temporary sheets exist without a matching checkpoint. Run resetExportState().'
      );
    }
    var todayYmd = Utilities.formatDate(new Date(), account.getTimeZone(), 'yyyy-MM-dd');
    var ranges = buildFrozenRanges_(todayYmd);
    campaignIds = fetchCampaignIdsRuntime_(ranges.aggregate);
    state = createRunState_(
      identity,
      nowMs,
      ranges,
      manifest.map(function(job) { return job.id; })
    );
    state.accountName = account.getName();
    state.accountCurrencyCode = account.getCurrencyCode ? account.getCurrencyCode() : '';
    state.accountTimeZone = account.getTimeZone();
    state.initialBlankSheetNames = workbookSummary.filter(function(summary) {
      return summary.blank;
    }).map(function(summary) { return summary.name; });
    if (ownership === 'owned') {
      writeExportInfoRuntime_(spreadsheet, state, manifest);
      SpreadsheetApp.flush();
      writeStartHereRuntime_(spreadsheet, state, manifest, { progressOnly: true });
      SpreadsheetApp.flush();
    }
    try {
      writeCampaignStateRuntime_(spreadsheet, state, campaignIds);
    } catch (checkpointCreationError) {
      state.status = 'COMPLETE_WITH_ERRORS';
      state.updatedAtMs = new Date().getTime();
      state.workbookError = (
        'Could not retain the checkpoint: ' +
        String(checkpointCreationError && checkpointCreationError.message ?
          checkpointCreationError.message : checkpointCreationError)
      ).substring(0, 500);
      try {
        writeExportInfoRuntime_(spreadsheet, state, manifest);
        SpreadsheetApp.flush();
      } catch (checkpointInfoError) {
        if (typeof Logger !== 'undefined') {
          Logger.log('Could not publish checkpoint-failure metadata: ' + String(
            checkpointInfoError && checkpointInfoError.message ?
              checkpointInfoError.message : checkpointInfoError
          ));
        }
      }
      try {
        writeStartHereRuntime_(spreadsheet, state, manifest, { progressOnly: true });
        SpreadsheetApp.flush();
      } catch (checkpointSummaryError) {
        if (typeof Logger !== 'undefined') {
          Logger.log('Could not publish checkpoint-failure summary: ' + String(
            checkpointSummaryError && checkpointSummaryError.message ?
              checkpointSummaryError.message : checkpointSummaryError
          ));
        }
      }
      throw checkpointCreationError;
    }
  }

  var context = {
    spreadsheet: spreadsheet,
    spreadsheetId: spreadsheetId,
    manifest: manifest,
    campaignIds: campaignIds,
    campaignChunksByJob: {}
  };
  var result = runManifestEngine_(
    state,
    manifest,
    createRuntimeAdapter_(context),
    CONFIG.MIN_REMAINING_SECONDS,
    CONFIG.MIN_COMMIT_REMAINING_SECONDS,
    CONFIG.INFO_REFRESH_INTERVAL_SECONDS
  );
  if (typeof Logger !== 'undefined') {
    Logger.log('Native Google Sheet status: ' + String(result && result.status || 'UNKNOWN'));
    if (result && (result.status === 'COMPLETE' ||
        result.status === 'COMPLETE_WITH_LIMITATIONS')) {
      Logger.log('Native Google Sheet ready: ' + canonicalSpreadsheetUrl_(spreadsheetId));
    }
  }
  return result;
}

function runtimeWriterOptions_() {
  return {
    batchRows: CONFIG.BATCH_ROWS,
    cellLimit: CONFIG.WORKBOOK_CELL_SAFETY_LIMIT,
    retries: 5,
    sleep: function(ms) { Utilities.sleep(ms); }
  };
}

function readSheetObjectsRuntime_(spreadsheet, name) {
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) return [];
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  var headers = values[0];
  return values.slice(1).filter(function(row) {
    return row.some(function(value) { return value !== ''; });
  }).map(function(row) {
    var object = {};
    headers.forEach(function(header, index) { object[header] = row[index]; });
    return object;
  });
}

function eligibleSourceRuntime_(context, job, state) {
  var definition = job && job.eligibleSource;
  if (!definition || !definition.tab) return { rows: null, limitations: [] };
  context.eligibleSourceCache = context.eligibleSourceCache || {};
  if (context.eligibleSourceCache[definition.tab]) {
    return context.eligibleSourceCache[definition.tab];
  }
  var limitations = derivedSourceLimitations_([definition.tab], state.tabs || {});
  var rows = readSheetObjectsRuntime_(context.spreadsheet, definition.tab);
  if (rows === null) {
    throw new Error('Eligible source tab ' + definition.tab + ' is unavailable for ' + job.tab + '.');
  }
  var result = { rows: rows, limitations: limitations };
  context.eligibleSourceCache[definition.tab] = result;
  return result;
}

function collectReportRowsRuntime_(query) {
  var iterator = reportRuntime_(query).rows();
  var rows = [];
  while (iterator.hasNext()) rows.push(iterator.next());
  return rows;
}

function runAdToLandingPageRuntime_(job, state, campaignIds, sheet) {
  var lookup = adToLandingPageActivityLookup_();
  var lookupJob = { id: job.id, activityLookup: lookup };
  var activeKeys = {};
  var activityIterator = reportRuntime_(
    buildActivityLookupQuery_(lookupJob, state.ranges.aggregate, campaignIds)
  ).rows();
  while (activityIterator.hasNext()) {
    var activityRow = activityIterator.next();
    if (hasSurveyActivity_(activityRow)) {
      activeKeys[entityKey_(activityRow, lookup.keyFields)] = true;
    }
  }
  var iterator = reportRuntime_(buildAdToLandingPageQuery_(campaignIds)).rows();
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  while (iterator.hasNext()) {
    buildAdToLandingPageRows_(iterator.next(), activeKeys).forEach(function(row) { writer.push(row); });
  }
  writer.flush();
  if (Number(CONFIG.THROTTLE_MS) > 0) Utilities.sleep(Number(CONFIG.THROTTLE_MS));
  return writer.count();
}

function runQualityScoreRuntime_(job, state, campaignIds, sheet) {
  var queries = buildQualityScoreQueries_(job, state.ranges.aggregate, campaignIds);
  var staticRows = collectReportRowsRuntime_(queries.staticQuery);
  var metricRows = collectReportRowsRuntime_(queries.metricsQuery);
  var output = buildQualityScoreRows_(job, staticRows, metricRows);
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  output.forEach(function(row) { writer.push(row); });
  writer.flush();
  if (Number(CONFIG.THROTTLE_MS) > 0) Utilities.sleep(Number(CONFIG.THROTTLE_MS));
  return writer.count();
}

function runEntityPerformanceInventoryRuntime_(job, state, campaignIds, sheet) {
  var queries = buildEntityPerformanceQueries_(job, state.ranges.aggregate, campaignIds);
  var currentRows = collectReportRowsRuntime_(queries.current);
  var activityRows = collectReportRowsRuntime_(queries.activity);
  var output = buildEntityPerformanceRows_(job, currentRows, activityRows);
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  output.forEach(function(row) { writer.push(row); });
  writer.flush();
  if (Number(CONFIG.THROTTLE_MS) > 0) Utilities.sleep(Number(CONFIG.THROTTLE_MS));
  return writer.count();
}

function runNegativeUnionRuntime_(context, job, state, sheet) {
  var names = [
    'neg_keywords_campaign', 'neg_keywords_ad_group', 'neg_keywords_shared',
    'neg_keyword_shared_links', 'neg_keyword_account_links'
  ];
  var tables = {};
  var limitations = derivedSourceLimitations_(names, state.tabs || {});
  names.forEach(function(name) {
    tables[name] = readSheetObjectsRuntime_(context.spreadsheet, name);
    if (tables[name] === null) limitations.push(name + ' unavailable');
  });
  if (tables.neg_keywords_campaign === null || tables.neg_keywords_ad_group === null) {
    throw new Error('Direct negative keyword source tabs are unavailable; unified output was not committed.');
  }
  names.forEach(function(name) { if (tables[name] === null) tables[name] = []; });
  var rows = buildNegativeUnionRows_(tables);
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  rows.forEach(function(row) { writer.push(row); });
  writer.flush();
  return limitations.length ?
    { rows: writer.count(), status: 'LIMITED', limitation: limitations.join('; ') } :
    writer.count();
}

function runAudienceRuntime_(job, state, campaignIds, sheet) {
  var queries = buildAudienceQueries_(state.ranges.aggregate, campaignIds);
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  var failures = [];
  var successfulScopes = 0;
  queries.forEach(function(definition) {
    try {
      var iterator = reportRuntime_(definition.query).rows();
      while (iterator.hasNext()) {
        var row = iterator.next();
        var output = [
          definition.scope,
          row[definition.resourceField] || '',
          row[definition.criterionField] || '',
          row['campaign.id'] || '',
          row['campaign.name'] || '',
          definition.scope === 'AD_GROUP' ? (row['ad_group.id'] || '') : '',
          definition.scope === 'AD_GROUP' ? (row['ad_group.name'] || '') : ''
        ];
        performanceColumns_().forEach(function(column) {
          output.push(coerceReportValue_(column, row[column.field]));
        });
        (job.derived || []).forEach(function(derived) { output.push(derived.compute(row)); });
        writer.push(output);
      }
      successfulScopes++;
    } catch (error) {
      failures.push(expectedPartialSourceLimitation_(definition.scope, error));
    }
  });
  writer.flush();
  if (!successfulScopes) throw new Error('Both audience performance scopes failed: ' + failures.join(' | '));
  return failures.length ?
    { rows: writer.count(), status: 'LIMITED', limitation: failures.join(' | ') } :
    writer.count();
}

function buildAssetExtensionQueries_(range) {
  if (!range || !range.start || !range.end) throw new Error('Asset extensions require a date range.');
  var assetFields = [
    'asset.id', 'asset.name', 'asset.type', 'asset.text_asset.text',
    'asset.callout_asset.callout_text', 'asset.sitelink_asset.link_text',
    'asset.sitelink_asset.description1', 'asset.sitelink_asset.description2',
    'asset.structured_snippet_asset.header', 'asset.structured_snippet_asset.values',
    'asset.promotion_asset.promotion_target',
    'asset.call_asset.phone_number'
  ];
  var customerFields = [
    'customer.id', 'customer.descriptive_name', 'customer.currency_code', 'customer.time_zone'
  ];
  var metricFields = performanceColumns_().map(function(column) { return column.field; });
  var dateClause = "segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'";
  return [
    {
      scope: 'CUSTOMER', prefix: 'customer_asset',
      query: 'SELECT\n  ' + customerFields.concat([
        'customer_asset.field_type', 'customer_asset.status', 'customer_asset.source'
      ], assetFields, metricFields).join(',\n  ') +
        '\nFROM customer_asset\nWHERE ' + dateClause
    },
    {
      scope: 'CAMPAIGN', prefix: 'campaign_asset',
      query: 'SELECT\n  ' + customerFields.concat([
        'campaign.id', 'campaign.name', 'campaign_asset.field_type',
        'campaign_asset.status', 'campaign_asset.source'
      ], assetFields, metricFields).join(',\n  ') +
        '\nFROM campaign_asset\nWHERE ' + dateClause
    },
    {
      scope: 'AD_GROUP', prefix: 'ad_group_asset',
      query: 'SELECT\n  ' + customerFields.concat([
        'campaign.id', 'campaign.name', 'ad_group.id', 'ad_group.name',
        'ad_group_asset.field_type', 'ad_group_asset.status', 'ad_group_asset.source'
      ], assetFields, metricFields).join(',\n  ') +
        '\nFROM ad_group_asset\nWHERE ' + dateClause
    }
  ];
}

function assetDisplayText_(row) {
  var candidates = [
    row['asset.text_asset.text'], row['asset.callout_asset.callout_text'],
    row['asset.sitelink_asset.link_text'], row['asset.structured_snippet_asset.header'],
    row['asset.structured_snippet_asset.values'], row['asset.promotion_asset.promotion_target'],
    row['asset.call_asset.phone_number']
  ];
  for (var index = 0; index < candidates.length; index++) {
    if (candidates[index] !== '' && candidates[index] !== null && candidates[index] !== undefined) {
      return String(candidates[index]);
    }
  }
  return '';
}

function buildAssetExtensionRow_(job, definition, row) {
  if (!hasSurveyActivity_(row)) return null;
  var output = [
    definition.scope,
    row['customer.id'] || '', row['customer.descriptive_name'] || '',
    row['customer.currency_code'] || '', row['customer.time_zone'] || '',
    row['campaign.id'] || '', row['campaign.name'] || '',
    row['ad_group.id'] || '', row['ad_group.name'] || '',
    row[definition.prefix + '.field_type'] || '',
    row[definition.prefix + '.status'] || '',
    row[definition.prefix + '.source'] || '',
    row['asset.id'] || '', row['asset.name'] || '', row['asset.type'] || '',
    assetDisplayText_(row),
    row['asset.sitelink_asset.description1'] || '',
    row['asset.sitelink_asset.description2'] || '',
    row['asset.structured_snippet_asset.header'] || '',
    row['asset.structured_snippet_asset.values'] || '',
    row['asset.promotion_asset.promotion_target'] || '',
    row['asset.call_asset.phone_number'] || ''
  ];
  performanceColumns_().forEach(function(column) {
    output.push(coerceReportValue_(column, row[column.field]));
  });
  (job.derived || []).forEach(function(derived) { output.push(derived.compute(row)); });
  return output;
}

function runAssetExtensionsRuntime_(job, state, sheet) {
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  var failures = [];
  var successfulScopes = 0;
  buildAssetExtensionQueries_(state.ranges.aggregate).forEach(function(definition) {
    try {
      var iterator = reportRuntime_(definition.query).rows();
      while (iterator.hasNext()) {
        var output = buildAssetExtensionRow_(job, definition, iterator.next());
        if (output) writer.push(output);
      }
      successfulScopes++;
    } catch (error) {
      failures.push(expectedPartialSourceLimitation_(definition.scope, error));
    }
  });
  writer.flush();
  if (!successfulScopes) throw new Error('All asset association scopes failed: ' + failures.join(' | '));
  return failures.length ?
    { rows: writer.count(), status: 'LIMITED', limitation: failures.join(' | ') } :
    writer.count();
}

function runChangeHistoryRuntime_(job, state, sheet) {
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  var result = paginateChangeHistory_(
    state.ranges.change,
    CONFIG.INCLUDE_SENSITIVE_CHANGE_DETAILS,
    {
      report: function(query) { return reportRuntime_(query); },
      remainingSeconds: function() { return AdsApp.getExecutionInfo().getRemainingTime(); },
      minRemainingSeconds: CONFIG.MIN_REMAINING_SECONDS
    },
    function(row) {
      writer.push(job.columns.map(function(column) {
        return coerceReportValue_(column, row[column.field]);
      }));
    },
    10000
  );
  writer.flush();
  return result.limited ?
    { rows: writer.count(), status: 'LIMITED', limitation: result.limitation } :
    writer.count();
}

function runDictionaryRuntime_(context, job, sheet) {
  var rows = buildDataDictionaryRows_(context.manifest).slice(1);
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  rows.forEach(function(row) { writer.push(row); });
  writer.flush();
  return writer.count();
}

function runFieldDictionaryRuntime_(context, job, sheet) {
  var rows = buildFieldDictionaryRows_(context.manifest).slice(1);
  var writer = createSafeRowBuffer_(sheet, headersForJob_(job), runtimeWriterOptions_());
  rows.forEach(function(row) { writer.push(row); });
  writer.flush();
  return writer.count();
}

function supportedJobKinds_() {
  return [
    'gaql', 'campaign_geo', 'entity_performance_inventory', 'ad_to_lp_map', 'quality_score',
    'negative_union', 'audience_performance', 'asset_extensions',
    'change_history', 'data_dictionary', 'field_dictionary'
  ];
}

function runJobChunkRuntime_(context, job, state, chunkIndex) {
  if (supportedJobKinds_().indexOf(job.kind) < 0) throw new Error('Unsupported job kind: ' + job.kind);
  var sheet = context.spreadsheet.getSheetByName(state.stageSheetName);
  if (!sheet) throw new Error('Missing staging sheet for ' + job.tab + '.');
  var campaignIds = job.chunked ? campaignChunksForJobRuntime_(context, job)[chunkIndex] : null;
  if (job.chunked && !campaignIds) throw new Error('Missing campaign chunk ' + chunkIndex + ' for ' + job.id + '.');

  if (job.kind === 'gaql') {
    var writerOptions = runtimeWriterOptions_();
    var source = eligibleSourceRuntime_(context, job, state);
    if (job.eligibleSource) writerOptions.eligibleSourceRows = source.rows;
    var gaqlRows = runGaqlChunk_(
      job,
      state.ranges,
      campaignIds,
      sheet,
      {
        report: function(query) { return reportRuntime_(query); },
        sleep: function(ms) { Utilities.sleep(ms); }
      },
      writerOptions
    );
    return source.limitations.length ?
      { rows: gaqlRows, status: 'LIMITED', limitation: source.limitations.join('; ') } :
      gaqlRows;
  }
  if (job.kind === 'campaign_geo') {
    context.geoTargetCache = context.geoTargetCache || {};
    var geoWriterOptions = runtimeWriterOptions_();
    geoWriterOptions.geoTargetCache = context.geoTargetCache;
    return runCampaignGeoChunk_(
      job,
      state.ranges,
      campaignIds,
      sheet,
      {
        report: function(query, options) { return reportRuntime_(query, options); },
        sleep: function(ms) { Utilities.sleep(ms); }
      },
      geoWriterOptions
    );
  }
  if (job.kind === 'entity_performance_inventory') {
    return runEntityPerformanceInventoryRuntime_(job, state, campaignIds, sheet);
  }
  if (job.kind === 'ad_to_lp_map') return runAdToLandingPageRuntime_(job, state, campaignIds, sheet);
  if (job.kind === 'quality_score') return runQualityScoreRuntime_(job, state, campaignIds, sheet);
  if (job.kind === 'negative_union') return runNegativeUnionRuntime_(context, job, state, sheet);
  if (job.kind === 'audience_performance') return runAudienceRuntime_(job, state, campaignIds, sheet);
  if (job.kind === 'asset_extensions') return runAssetExtensionsRuntime_(job, state, sheet);
  if (job.kind === 'change_history') return runChangeHistoryRuntime_(job, state, sheet);
  if (job.kind === 'data_dictionary') return runDictionaryRuntime_(context, job, sheet);
  if (job.kind === 'field_dictionary') return runFieldDictionaryRuntime_(context, job, sheet);
  throw new Error('No runtime implementation for ' + job.kind + '.');
}

function redactLogText_(value) {
  return String(value || '')
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, '[REDACTED_EMAIL]')
    .replace(/customers\/\d+/gi, 'customers/[REDACTED]')
    .replace(/\b\d{10,}\b/g, '[REDACTED_ID]');
}

function diagnosticProbes_(range) {
  return [
    {
      name: 'campaign',
      fields: [
        'campaign.id', 'campaign.start_date_time',
        'campaign.end_date_time', 'metrics.interactions'
      ],
      query: 'SELECT\n  campaign.id,\n  campaign.start_date_time,\n  campaign.end_date_time,\n' +
        '  campaign_budget.amount_micros,\n  campaign_budget.recommended_budget_amount_micros,\n' +
        '  campaign.maximize_conversions.target_cpa_micros,\n' +
        '  campaign.target_cpa.target_cpa_micros,\n  metrics.interactions\nFROM campaign\n' +
        "WHERE segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'campaign_geo_raw_ids',
      fields: [
        'geographic_view.country_criterion_id',
        'segments.geo_target_most_specific_location',
        'segments.geo_target_state'
      ],
      reportOptions: { resolveGeoNames: false },
      query: 'SELECT\n  geographic_view.country_criterion_id,\n' +
        '  segments.geo_target_most_specific_location,\n  segments.geo_target_state,\n' +
        '  metrics.interactions\n' +
        'FROM geographic_view\nWHERE segments.date BETWEEN \'' + range.start +
        '\' AND \'' + range.end + '\'\nLIMIT 1'
    },
    {
      name: 'geo_target_constant_lookup',
      fields: [
        'geo_target_constant.id', 'geo_target_constant.name',
        'geo_target_constant.canonical_name'
      ],
      query: 'SELECT\n  geo_target_constant.id,\n  geo_target_constant.name,\n' +
        '  geo_target_constant.canonical_name\nFROM geo_target_constant\nLIMIT 1'
    },
    {
      name: 'rsa_assets',
      fields: ['campaign.id', 'ad_group_ad.ad.id', 'asset.id', 'metrics.interactions'],
      query: 'SELECT\n  campaign.id,\n  ad_group_ad.ad.id,\n  asset.id,\n  metrics.interactions\n' +
        "FROM ad_group_ad_asset_view\nWHERE campaign.advertising_channel_type = 'SEARCH'\n" +
        "  AND ad_group_ad.ad.type = 'RESPONSIVE_SEARCH_AD'\n" +
        "  AND segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'shared_negatives', fields: ['shared_set.id', 'shared_criterion.criterion_id'],
      query: 'SELECT\n  shared_set.id,\n  shared_criterion.criterion_id\nFROM shared_criterion\n' +
        "WHERE shared_set.type IN ('NEGATIVE_KEYWORDS', 'ACCOUNT_LEVEL_NEGATIVE_KEYWORDS')\nLIMIT 1"
    },
    {
      name: 'account_negative_links',
      fields: ['customer_negative_criterion.negative_keyword_list.shared_set'],
      query: 'SELECT\n  customer_negative_criterion.negative_keyword_list.shared_set\n' +
        "FROM customer_negative_criterion\nWHERE customer_negative_criterion.type = 'NEGATIVE_KEYWORD_LIST'\nLIMIT 1"
    },
    {
      name: 'proximity', fields: ['campaign_criterion.criterion_id'],
      query: 'SELECT\n  campaign_criterion.criterion_id,\n  campaign_criterion.proximity.radius\n' +
        "FROM campaign_criterion\nWHERE campaign_criterion.type = 'PROXIMITY'\nLIMIT 1"
    },
    {
      name: 'campaign_audience',
      fields: ['campaign_criterion.user_list.user_list', 'metrics.interactions'],
      query: 'SELECT\n  campaign_criterion.user_list.user_list,\n  metrics.interactions\n' +
        'FROM campaign_audience_view\n' +
        "WHERE segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'ad_group_audience',
      fields: ['ad_group_criterion.user_list.user_list', 'metrics.interactions'],
      query: 'SELECT\n  ad_group_criterion.user_list.user_list,\n  metrics.interactions\n' +
        'FROM ad_group_audience_view\n' +
        "WHERE segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'pmax_asset_group_inventory',
      fields: ['campaign.id', 'asset_group.id', 'asset_group.status'],
      query: 'SELECT\n  campaign.id,\n  asset_group.id,\n  asset_group.status,\n' +
        '  asset_group.ad_strength,\n  asset_group.final_urls\nFROM asset_group\n' +
        "WHERE campaign.advertising_channel_type = 'PERFORMANCE_MAX'\nLIMIT 1"
    },
    {
      name: 'pmax_asset_group_metrics',
      fields: ['campaign.id', 'asset_group.id', 'metrics.impressions'],
      query: 'SELECT\n  campaign.id,\n  asset_group.id,\n  asset_group.status,\n' +
        '  metrics.impressions,\n  metrics.clicks,\n  metrics.interactions,\n' +
        '  metrics.cost_micros\nFROM asset_group\n' +
        "WHERE campaign.advertising_channel_type = 'PERFORMANCE_MAX'\n  AND segments.date BETWEEN '" +
        range.start + "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'pmax_signals', fields: ['asset_group_signal.resource_name'],
      query: 'SELECT\n  asset_group_signal.resource_name,\n  asset_group_signal.audience.audience,\n' +
        '  asset_group_signal.search_theme.text\nFROM asset_group_signal\nLIMIT 1'
    },
    {
      name: 'pmax_asset_status',
      fields: ['asset_group_asset.primary_status', 'asset.id', 'metrics.interactions'],
      query: 'SELECT\n  asset_group_asset.primary_status,\n' +
        '  asset_group_asset.primary_status_reasons,\n  asset_group_asset.source,\n' +
        '  asset_group_asset.policy_summary.approval_status,\n  asset.id,\n' +
        '  metrics.interactions\nFROM asset_group_asset\n' +
        "WHERE segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'quality_score_static', fields: ['ad_group.id', 'ad_group_criterion.criterion_id'],
      query: 'SELECT\n  ad_group.id,\n  ad_group_criterion.criterion_id,\n' +
        '  ad_group_criterion.quality_info.quality_score\nFROM ad_group_criterion\n' +
        "WHERE ad_group_criterion.type = 'KEYWORD'\n  AND ad_group_criterion.negative = FALSE\nLIMIT 1"
    },
    {
      name: 'quality_score_metrics',
      fields: [
        'campaign.id', 'ad_group.id',
        'ad_group_criterion.criterion_id', 'metrics.interactions'
      ],
      query: 'SELECT\n  campaign.id,\n  ad_group.id,\n' +
        '  ad_group_criterion.criterion_id,\n  metrics.interactions\nFROM keyword_view\n' +
        "WHERE segments.date BETWEEN '" + range.start + "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'asset_extension_metrics', fields: ['campaign_asset.field_type', 'asset.id'],
      query: 'SELECT\n  campaign_asset.field_type,\n  campaign_asset.status,\n  campaign_asset.source,\n' +
        '  asset.id,\n  asset.sitelink_asset.description1,\n  asset.promotion_asset.promotion_target,\n' +
        '  metrics.impressions,\n  metrics.interactions\n' +
        "FROM campaign_asset\nWHERE segments.date BETWEEN '" + range.start +
        "' AND '" + range.end + "'\nLIMIT 1"
    },
    {
      name: 'change_history_safe', fields: changeHistoryColumns_(false).map(function(column) { return column.field; }),
      query: buildChangeHistoryQuery_({ start: range.start, end: range.end }, range.end + ' 23:59:59', false, 1)
    }
  ];
}

function runDiagnostics_() {
  validateRuntimeConfig_();
  var spreadsheetId = extractSpreadsheetId_(CONFIG.SPREADSHEET_URL);
  var results = [];
  Logger.log('Google Ads Analysis Workbook ' + VERSION + ' compatibility diagnostics');
  Logger.log('No spreadsheet writes and no raw account rows by default.');
  var spreadsheet = openSpreadsheetRuntime_(spreadsheetId);
  var ownership = assertWorkbookOwnership_(summarizeWorkbookRuntime_(spreadsheet));
  results.push({
    probe: 'native_sheet_target', status: 'SUPPORTED', rowsRead: 0, error: ''
  });
  Logger.log('native_sheet_target: SUPPORTED; ownership=' + ownership);
  assertAdvertiserAccountRuntime_();
  var account = AdsApp.currentAccount();
  var today = Utilities.formatDate(new Date(), account.getTimeZone(), 'yyyy-MM-dd');
  var end = addDaysYmd_(today, -1);
  var range = { start: addDaysYmd_(end, -6), end: end };
  diagnosticProbes_(range).forEach(function(probe) {
    try {
      var iterator = reportRuntime_(probe.query, probe.reportOptions).rows();
      var present = iterator.hasNext();
      var row = present ? iterator.next() : null;
      var result = { probe: probe.name, status: 'SUPPORTED', rowsRead: present ? 1 : 0, error: '' };
      results.push(result);
      Logger.log(probe.name + ': SUPPORTED; rows_read=' + result.rowsRead);
      if (present && CONFIG.DIAGNOSTICS_LOG_SAMPLE_ROWS) {
        var selected = {};
        probe.fields.forEach(function(field) { selected[field] = row[field]; });
        Logger.log(probe.name + ' redacted_sample=' + JSON.stringify(redactDiagnosticSample_(selected)));
      }
    } catch (error) {
      var message = redactLogText_(error && error.message ? error.message : error);
      results.push({ probe: probe.name, status: 'UNSUPPORTED_OR_ERROR', rowsRead: 0, error: message });
      Logger.log(probe.name + ': UNSUPPORTED_OR_ERROR; ' + message);
    }
  });
  Logger.log('Diagnostics complete. Unsupported optional resources will be labeled LIMITED during export.');
  return results;
}

function resetExportState_() {
  validateRuntimeConfig_();
  var spreadsheetId = extractSpreadsheetId_(CONFIG.SPREADSHEET_URL);
  var spreadsheet = openSpreadsheetRuntime_(spreadsheetId);
  assertWorkbookOwnership_(summarizeWorkbookRuntime_(spreadsheet));
  var names = spreadsheet.getSheets().map(function(sheet) { return sheet.getName(); });
  var targets = planReset_(CONFIG.ALLOW_RESET, names);

  targets.filter(function(name) { return name.indexOf(BACKUP_PREFIX) === 0; }).forEach(function(name) {
    var backup = spreadsheet.getSheetByName(name);
    if (!backup) return;
    var finalName = name.substring(BACKUP_PREFIX.length);
    var finalSheet = spreadsheet.getSheetByName(finalName);
    if (finalSheet) spreadsheet.deleteSheet(backup);
    else backup.setName(finalName);
  });
  targets.filter(function(name) { return name.indexOf(STAGE_PREFIX) === 0; }).forEach(function(name) {
    var stage = spreadsheet.getSheetByName(name);
    if (stage) spreadsheet.deleteSheet(stage);
  });
  var stateSheet = spreadsheet.getSheetByName(STATE_SHEET_NAME);
  if (stateSheet) spreadsheet.deleteSheet(stateSheet);

  var manifest = getManifestDefinition_();
  var resetState = {
    runId: 'reset-' + new Date().getTime(), status: 'RESET', startedAtMs: new Date().getTime(),
    updatedAtMs: new Date().getTime(), accountId: '', accountName: '', currentJobId: '',
    spreadsheetId: spreadsheetId,
    outputSchemaVersion: OUTPUT_SCHEMA_VERSION,
    runtimeContractVersion: RUNTIME_CONTRACT_VERSION,
    jobIndex: 0, ranges: {}, tabs: {}
  };
  writeExportInfoRuntime_(spreadsheet, resetState, manifest);
  Logger.log('Exporter checkpoint and temporary sheets reset. Existing final report tabs were preserved.');
  Logger.log('Set CONFIG.ALLOW_RESET back to false before the next export.');
  return targets.length;
}

function withRuntimeLock_(operation) {
  var lock = typeof LockService !== 'undefined' ? LockService.getScriptLock() : null;
  if (lock && !lock.tryLock(1000)) {
    throw new Error('Another exporter invocation is already running. Try again after it finishes.');
  }
  try { return operation(); }
  finally { if (lock) lock.releaseLock(); }
}

function main() {
  if (AdsApp.getExecutionInfo().isPreview()) return runDiagnostics_();
  return withRuntimeLock_(runExport_);
}

function runDiagnostics() {
  return runDiagnostics_();
}

function resetExportState() {
  if (AdsApp.getExecutionInfo().isPreview()) {
    throw new Error('resetExportState() is disabled in Preview. Click Run to perform the guarded reset.');
  }
  return withRuntimeLock_(resetExportState_);
}
