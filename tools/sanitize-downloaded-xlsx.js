#!/usr/bin/env node
/* SPDX-License-Identifier: MIT */
'use strict';

const fs = require('node:fs');
const crypto = require('node:crypto');
const path = require('node:path');
const { TextDecoder } = require('node:util');
const zlib = require('node:zlib');

const PACKAGE_RELATIONSHIP_NS =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const DOCUMENT_RELATIONSHIP_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SPREADSHEET_NS =
  'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CONTENT_TYPES_NS =
  'http://schemas.openxmlformats.org/package/2006/content-types';
const HYPERLINK_RELATIONSHIP =
  `${DOCUMENT_RELATIONSHIP_NS}/hyperlink`;
const WORKSHEET_RELATIONSHIP =
  `${DOCUMENT_RELATIONSHIP_NS}/worksheet`;
const OFFICE_DOCUMENT_RELATIONSHIP = `${DOCUMENT_RELATIONSHIP_NS}/officeDocument`;
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CANONICAL_GOOGLE_SHEET =
  /^https:\/\/docs\.google\.com\/spreadsheets\/d\/[A-Za-z0-9_-]+\/edit$/;
const CELL_REFERENCE = /^\$?[A-Z]{1,3}\$?[1-9][0-9]*$/;
const CELL_OR_RANGE_REFERENCE =
  /^\$?[A-Z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$/;
const NAVIGATION_LOCATION = /^gid=(?:0|[1-9][0-9]*)&range=([A-Z]{1,3}[1-9][0-9]*)$/;
const START_HERE_DIRECTORY_ORDER = [
  'campaign', 'campaign_weekly', 'imp_share', 'keywords', 'search_terms', 'ads',
  'landing_pages', 'campaign_device_network', 'ad_schedule', 'campaign_geo',
  'ad_group_weekly', 'pmax_asset_group_weekly', 'ad_group', 'quality_score_keywords',
  'campaign_inventory', 'ad_group_inventory', 'keyword_inventory', 'ad_inventory',
  'ad_to_lp_map', 'rsa_assets', 'demandgen_assets', 'pmax_asset_groups', 'pmax_assets',
  'asset_extensions', 'geo_targets', 'geo_proximity_targets', 'neg_keywords_campaign',
  'neg_keywords_ad_group', 'neg_keywords_shared', 'neg_keyword_shared_links',
  'neg_keyword_account_links', 'negative_keywords_all', 'pmax_audience_signals',
  'user_list_performance', 'conversion_actions', 'conversion_action_config',
  'change_history', '_export_info', '_data_dictionary', '_field_dictionary',
];

const ZIP_LOCAL_FILE_HEADER = 0x04034b50;
const ZIP_CENTRAL_DIRECTORY_HEADER = 0x02014b50;
const ZIP_END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const MAX_ARCHIVE_BYTES = 128 * 1024 * 1024;
const MAX_COMPRESSED_ENTRY_BYTES = 64 * 1024 * 1024;
const MAX_UNCOMPRESSED_ENTRY_BYTES = 128 * 1024 * 1024;
const MAX_TOTAL_UNCOMPRESSED_BYTES = 256 * 1024 * 1024;
const MAX_COMPRESSION_RATIO = 100;
const fatalUtf8 = new TextDecoder('utf-8', { fatal: true });

function asText(data, name) {
  if (typeof data === 'string') return data;
  if (Buffer.isBuffer(data) || data instanceof Uint8Array) {
    try {
      return fatalUtf8.decode(Buffer.from(data));
    } catch (error) {
      throw new Error(`XML entry ${name || '(unknown)'} is not valid UTF-8.`);
    }
  }
  throw new TypeError(`ZIP entry ${name || '(unknown)'} has unsupported data.`);
}

function withText(entry, text) {
  return {
    ...entry,
    data: typeof entry.data === 'string' ? text : Buffer.from(text, 'utf8'),
  };
}

function decodeXml(value, context) {
  const source = String(value);
  if (/&(?!#x[0-9A-Fa-f]+;|#[0-9]+;|amp;|lt;|gt;|quot;|apos;)/.test(source)) {
    throw new Error(`Malformed XML entity in ${context}.`);
  }
  return source.replace(
    /&(#x[0-9A-Fa-f]+|#[0-9]+|amp|lt|gt|quot|apos);/g,
    (match, entity) => {
      if (entity === 'amp') return '&';
      if (entity === 'lt') return '<';
      if (entity === 'gt') return '>';
      if (entity === 'quot') return '"';
      if (entity === 'apos') return "'";
      const radix = entity.startsWith('#x') ? 16 : 10;
      const digits = entity.slice(radix === 16 ? 2 : 1);
      const codePoint = Number.parseInt(digits, radix);
      if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff ||
          (codePoint >= 0xd800 && codePoint <= 0xdfff) ||
          (codePoint < 0x20 && codePoint !== 0x09 && codePoint !== 0x0a && codePoint !== 0x0d)) {
        throw new Error(`Malformed XML entity in ${context}.`);
      }
      return String.fromCodePoint(codePoint);
    },
  );
}

const XML_NAME = /^[A-Za-z_][A-Za-z0-9_.-]*(?::[A-Za-z_][A-Za-z0-9_.-]*)?/;
const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';

function assertXmlCharacters(value, context) {
  if (/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value) ||
      /[\ud800-\udbff](?![\udc00-\udfff])|(^|[^\ud800-\udbff])[\udc00-\udfff]/.test(value)) {
    throw new Error(`Invalid XML character in ${context}.`);
  }
}

function qualifiedName(value, context) {
  const match = XML_NAME.exec(value);
  if (!match || match[0].length !== value.length) {
    throw new Error(`Malformed XML name in ${context}.`);
  }
  const parts = value.split(':');
  return {
    name: value,
    prefix: parts.length === 2 ? parts[0] : '',
    localName: parts.length === 2 ? parts[1] : parts[0],
  };
}

function parseStrictStartTag(raw, context) {
  let inner = raw.slice(1, -1);
  const selfClosing = /\/\s*$/.test(inner);
  if (selfClosing) inner = inner.replace(/\/\s*$/, '');
  let index = 0;
  while (/\s/.test(inner[index] || '')) index += 1;
  const nameMatch = XML_NAME.exec(inner.slice(index));
  if (!nameMatch) throw new Error(`Malformed XML start tag in ${context}.`);
  const name = qualifiedName(nameMatch[0], context);
  index += nameMatch[0].length;
  const attributes = [];
  const names = new Set();
  while (index < inner.length) {
    if (!/\s/.test(inner[index] || '')) {
      throw new Error(`Missing XML attribute separator in ${context}.`);
    }
    while (/\s/.test(inner[index] || '')) index += 1;
    if (index >= inner.length) break;
    const attributeMatch = XML_NAME.exec(inner.slice(index));
    if (!attributeMatch) throw new Error(`Malformed XML attribute in ${context}.`);
    const attributeName = qualifiedName(attributeMatch[0], context);
    if (names.has(attributeName.name)) {
      throw new Error(`Duplicate XML attribute ${attributeName.name} in ${context}.`);
    }
    names.add(attributeName.name);
    index += attributeName.name.length;
    while (/\s/.test(inner[index] || '')) index += 1;
    if (inner[index] !== '=') throw new Error(`Malformed XML attribute in ${context}.`);
    index += 1;
    while (/\s/.test(inner[index] || '')) index += 1;
    const quote = inner[index];
    if (quote !== '"' && quote !== "'") throw new Error(`Unquoted XML attribute in ${context}.`);
    index += 1;
    const end = inner.indexOf(quote, index);
    if (end < 0) throw new Error(`Unterminated XML attribute in ${context}.`);
    const encodedValue = inner.slice(index, end);
    if (encodedValue.includes('<')) throw new Error(`Invalid XML attribute value in ${context}.`);
    assertXmlCharacters(encodedValue, context);
    attributes.push({ ...attributeName, value: decodeXml(encodedValue, context) });
    index = end + 1;
  }
  return { ...name, attributes, selfClosing };
}

function findMarkupEnd(xml, start, context) {
  let quote = '';
  for (let index = start; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote) {
      if (character === quote) quote = '';
    } else if (character === '"' || character === "'") {
      quote = character;
    } else if (character === '>') {
      return index;
    } else if (character === '<') {
      throw new Error(`Nested XML markup delimiter in ${context}.`);
    }
  }
  throw new Error(`Unterminated XML markup in ${context}.`);
}

function validateXmlText(value, context, outsideRoot) {
  assertXmlCharacters(value, context);
  if (value.includes(']]>')) throw new Error(`Invalid XML text terminator in ${context}.`);
  decodeXml(value, context);
  if (outsideRoot && value.trim()) throw new Error(`XML text exists outside the root in ${context}.`);
}

function validateXmlStructure(xml, context, namespaceRules) {
  assertXmlCharacters(xml, context);
  const rules = namespaceRules || {};
  const stack = [];
  const baseScope = new Map([['xml', XML_NAMESPACE]]);
  let root = null;
  let rootClosed = false;
  let index = xml.charCodeAt(0) === 0xfeff ? 1 : 0;
  let declarationSeen = false;

  while (index < xml.length) {
    const opening = xml.indexOf('<', index);
    if (opening < 0) {
      validateXmlText(xml.slice(index), context, stack.length === 0);
      index = xml.length;
      break;
    }
    validateXmlText(xml.slice(index, opening), context, stack.length === 0);

    if (xml.startsWith('<!--', opening)) {
      const end = xml.indexOf('-->', opening + 4);
      if (end < 0 || xml.slice(opening + 4, end).includes('--')) {
        throw new Error(`Malformed XML comment in ${context}.`);
      }
      throw new Error(`XML comments are not permitted in inspected part ${context}.`);
    }
    if (xml.startsWith('<![CDATA[', opening)) {
      throw new Error(`CDATA is not permitted in inspected XML part ${context}.`);
    }
    if (/^<!DOCTYPE\b/i.test(xml.slice(opening, opening + 12))) {
      throw new Error(`DOCTYPE is not permitted in inspected XML part ${context}.`);
    }
    if (xml.startsWith('<!', opening)) {
      throw new Error(`Unsupported XML declaration in ${context}.`);
    }
    if (xml.startsWith('<?', opening)) {
      const end = xml.indexOf('?>', opening + 2);
      if (end < 0) throw new Error(`Unterminated XML processing instruction in ${context}.`);
      const raw = xml.slice(opening, end + 2);
      if (declarationSeen || root ||
          !/^<\?xml\s+version=(?:"1\.0"|'1\.0')(?:\s+encoding=(?:"UTF-8"|'UTF-8'|"utf-8"|'utf-8'))?(?:\s+standalone=(?:"(?:yes|no)"|'(?:yes|no)'))?\s*\?>$/.test(raw)) {
        throw new Error(`Unsupported XML processing instruction in ${context}.`);
      }
      declarationSeen = true;
      index = end + 2;
      continue;
    }
    if (xml.startsWith('</', opening)) {
      const end = xml.indexOf('>', opening + 2);
      if (end < 0) throw new Error(`Unterminated XML closing tag in ${context}.`);
      const closingName = xml.slice(opening + 2, end).trim();
      qualifiedName(closingName, context);
      const current = stack.pop();
      if (!current || current.name !== closingName) {
        throw new Error(`Unbalanced XML closing tag ${closingName} in ${context}.`);
      }
      if (stack.length === 0) rootClosed = true;
      index = end + 1;
      continue;
    }

    const end = findMarkupEnd(xml, opening + 1, context);
    const rawTag = xml.slice(opening, end + 1);
    const parsed = parseStrictStartTag(rawTag, context);
    if (rootClosed && stack.length === 0) throw new Error(`Multiple XML roots in ${context}.`);
    const parentScope = stack.length ? stack[stack.length - 1].scope : baseScope;
    let scope = parentScope;
    const namespaceAttributes = parsed.attributes.filter(
      (attribute) => attribute.name === 'xmlns' || attribute.prefix === 'xmlns',
    );
    if (namespaceAttributes.length) {
      scope = new Map(parentScope);
      for (const attribute of namespaceAttributes) {
        const prefix = attribute.name === 'xmlns' ? '' : attribute.localName;
        if (prefix === 'xml' || prefix === 'xmlns' || (prefix && !attribute.value)) {
          throw new Error(`Invalid XML namespace declaration in ${context}.`);
        }
        scope.set(prefix, attribute.value);
      }
    }
    const namespace = scope.get(parsed.prefix || '') || '';
    if (parsed.prefix && !scope.has(parsed.prefix)) {
      throw new Error(`Unbound XML namespace prefix ${parsed.prefix} in ${context}.`);
    }
    if (Object.prototype.hasOwnProperty.call(rules, parsed.localName) &&
        namespace !== rules[parsed.localName]) {
      throw new Error(`Wrong XML namespace for ${parsed.localName} in ${context}.`);
    }
    const expandedAttributes = new Set();
    for (const attribute of parsed.attributes) {
      if (attribute.name === 'xmlns' || attribute.prefix === 'xmlns') continue;
      const attributeNamespace = attribute.prefix ? scope.get(attribute.prefix) : '';
      if (attribute.prefix && attributeNamespace === undefined) {
        throw new Error(`Unbound XML attribute prefix ${attribute.prefix} in ${context}.`);
      }
      const expanded = `{${attributeNamespace || ''}}${attribute.localName}`;
      if (expandedAttributes.has(expanded)) {
        throw new Error(`Duplicate expanded XML attribute in ${context}.`);
      }
      expandedAttributes.add(expanded);
    }
    const element = { ...parsed, namespace, scope, raw: rawTag };
    if (!root) root = element;
    if (!parsed.selfClosing) stack.push(element);
    else if (stack.length === 0) rootClosed = true;
    index = end + 1;
  }
  if (!root || stack.length || !rootClosed) {
    throw new Error(`XML document is not balanced in ${context}.`);
  }
  return root;
}

function parseAttributes(tag, localName, context) {
  const opening = new RegExp(
    `^<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?${localName}\\b([\\s\\S]*?)\\/?>$`,
  ).exec(tag);
  if (!opening) throw new Error(`Malformed ${localName} element in ${context}.`);
  let source = opening[1];
  if (source.endsWith('/')) source = source.slice(0, -1);
  const attributes = Object.create(null);
  let index = 0;
  while (index < source.length) {
    while (/\s/.test(source[index] || '')) index += 1;
    if (index >= source.length) break;
    const nameMatch = /^[A-Za-z_][A-Za-z0-9_.:-]*/.exec(source.slice(index));
    if (!nameMatch) throw new Error(`Malformed attribute in ${context}.`);
    const name = nameMatch[0];
    index += name.length;
    while (/\s/.test(source[index] || '')) index += 1;
    if (source[index] !== '=') throw new Error(`Malformed attribute ${name} in ${context}.`);
    index += 1;
    while (/\s/.test(source[index] || '')) index += 1;
    const quote = source[index];
    if (quote !== '"' && quote !== "'") {
      throw new Error(`Malformed attribute ${name} in ${context}.`);
    }
    index += 1;
    const end = source.indexOf(quote, index);
    if (end < 0) throw new Error(`Unterminated attribute ${name} in ${context}.`);
    if (Object.prototype.hasOwnProperty.call(attributes, name)) {
      throw new Error(`Duplicate attribute ${name} in ${context}.`);
    }
    attributes[name] = decodeXml(source.slice(index, end), context);
    index = end + 1;
  }
  return attributes;
}

function matchingTags(xml, localName, context) {
  const opener = new RegExp(`<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?${localName}\\b`, 'g');
  const tagPattern = new RegExp(
    `<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?${localName}\\b[^<>]*\\/>`,
    'g',
  );
  const openers = xml.match(opener) || [];
  const tags = xml.match(tagPattern) || [];
  if (openers.length !== tags.length) {
    throw new Error(`Malformed ${localName} element in ${context}.`);
  }
  return tags;
}

function assertRootNamespace(xml, localName, namespace, context, childRules) {
  const root = validateXmlStructure(xml, context, {
    ...(childRules || {}),
    [localName]: namespace,
  });
  if (root.localName !== localName || root.namespace !== namespace) {
    throw new Error(`Incorrect XML namespace for ${localName} in ${context}.`);
  }
  const attributes = Object.create(null);
  root.attributes.forEach((attribute) => { attributes[attribute.name] = attribute.value; });
  return { attributes, prefix: root.prefix };
}

function relationshipAttributeName(attributes, context) {
  const candidates = Object.keys(attributes).filter((name) => /:id$/.test(name));
  if (candidates.length !== 1 || !attributes[candidates[0]]) {
    if (attributes.location) {
      throw new Error(`Location-only hyperlink has no relationship id in ${context}.`);
    }
    throw new Error(`Hyperlink is missing its relationship id in ${context}.`);
  }
  return candidates[0];
}

function relationshipId(attributes, context) {
  const name = relationshipAttributeName(attributes, context);
  return attributes[name];
}

function relationshipEntryName(worksheetName) {
  return path.posix.join(
    path.posix.dirname(worksheetName),
    '_rels',
    `${path.posix.basename(worksheetName)}.rels`,
  );
}

function resolveRelationshipTarget(sourceName, target, context) {
  if (!target || target.includes('\\') || target.includes('\0')) {
    throw new Error(`Malformed relationship target in ${context}.`);
  }
  const resolved = target.startsWith('/')
    ? path.posix.normalize(target.slice(1))
    : path.posix.normalize(path.posix.join(path.posix.dirname(sourceName), target));
  if (!resolved || resolved === '..' || resolved.startsWith('../') || path.posix.isAbsolute(resolved)) {
    throw new Error(`Relationship target escapes the XLSX package in ${context}.`);
  }
  return resolved;
}

function parseRelationships(xml, context) {
  assertRootNamespace(xml, 'Relationships', PACKAGE_RELATIONSHIP_NS, context, {
    Relationship: PACKAGE_RELATIONSHIP_NS,
  });
  const relationships = new Map();
  for (const raw of matchingTags(xml, 'Relationship', context)) {
    const attributes = parseAttributes(raw, 'Relationship', context);
    const id = attributes.Id;
    if (!id || !attributes.Type || !attributes.Target) {
      throw new Error(`Malformed relationship in ${context}.`);
    }
    if (relationships.has(id)) throw new Error(`Duplicate relationship id ${id} in ${context}.`);
    relationships.set(id, { raw, attributes });
  }
  return relationships;
}

function contentTypeOverrides(entriesByName) {
  const entry = entriesByName.get('[Content_Types].xml');
  if (!entry) throw new Error('XLSX package is missing [Content_Types].xml.');
  const xml = asText(entry.data, entry.name);
  assertRootNamespace(xml, 'Types', CONTENT_TYPES_NS, entry.name, {
    Default: CONTENT_TYPES_NS,
    Override: CONTENT_TYPES_NS,
  });
  const overrides = new Map();
  for (const raw of matchingTags(xml, 'Override', entry.name)) {
    const attributes = parseAttributes(raw, 'Override', entry.name);
    const partName = attributes.PartName;
    if (!partName || !partName.startsWith('/') || !attributes.ContentType) {
      throw new Error('Malformed content-type override in [Content_Types].xml.');
    }
    const normalized = partName.slice(1);
    if (path.posix.normalize(normalized) !== normalized || overrides.has(normalized)) {
      throw new Error('Unsafe or duplicate content-type override in [Content_Types].xml.');
    }
    overrides.set(normalized, attributes.ContentType);
  }
  if (overrides.get('xl/workbook.xml') !== WORKBOOK_CONTENT_TYPE) {
    throw new Error('XLSX workbook content type is missing or incorrect.');
  }
  return overrides;
}

function validateOpcSkeleton(entriesByName) {
  const overrides = contentTypeOverrides(entriesByName);
  const rootEntry = entriesByName.get('_rels/.rels');
  if (!rootEntry) throw new Error('XLSX package is missing its root relationship part.');
  const relationships = parseRelationships(asText(rootEntry.data, rootEntry.name), rootEntry.name);
  const officeDocuments = Array.from(relationships.values()).filter(
    (relationship) => relationship.attributes.Type === OFFICE_DOCUMENT_RELATIONSHIP,
  );
  if (officeDocuments.length !== 1) {
    throw new Error('XLSX package must have exactly one officeDocument workbook relationship.');
  }
  const officeDocument = officeDocuments[0].attributes;
  if (officeDocument.TargetMode === 'External' ||
      resolveRelationshipTarget('', officeDocument.Target, 'root relationship') !== 'xl/workbook.xml') {
    throw new Error('Root officeDocument relationship does not target xl/workbook.xml.');
  }
  if (!entriesByName.has('xl/workbook.xml')) {
    throw new Error('Root officeDocument workbook part is missing.');
  }
  return overrides;
}

function assertNoHyperlinkFormulas(xml, context) {
  const formulaPattern = /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?f\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?f\s*>/g;
  let match;
  while ((match = formulaPattern.exec(xml)) !== null) {
    const formula = decodeXml(match[1].replace(/<[^>]*>/g, ''), context);
    if (/\bHYPERLINK\s*\(/i.test(formula)) {
      throw new Error(`HYPERLINK formula is not permitted in ${context}.`);
    }
  }
  const formulaOpeners = xml.match(/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?f\b/g) || [];
  const pairedFormulas = xml.match(formulaPattern) || [];
  const selfClosingFormulas = xml.match(/<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?f\b[^<>]*\/>/g) || [];
  if (formulaOpeners.length !== pairedFormulas.length + selfClosingFormulas.length) {
    throw new Error(`Malformed formula element in ${context}.`);
  }
}

function textNodes(xml, context) {
  const values = [];
  const pattern = /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?t\s*>/g;
  let match;
  while ((match = pattern.exec(xml)) !== null) {
    if (/<[^>]+>/.test(match[1])) throw new Error(`Malformed text node in ${context}.`);
    values.push(decodeXml(match[1], context));
  }
  return values.join('');
}

function sharedStrings(entriesByName) {
  const entry = entriesByName.get('xl/sharedStrings.xml');
  if (!entry) return [];
  const xml = asText(entry.data, entry.name);
  assertRootNamespace(xml, 'sst', SPREADSHEET_NS, entry.name, {
    si: SPREADSHEET_NS,
    r: SPREADSHEET_NS,
    rPr: SPREADSHEET_NS,
    t: SPREADSHEET_NS,
  });
  const strings = [];
  const pattern = /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?si\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?si\s*>/g;
  let match;
  while ((match = pattern.exec(xml)) !== null) strings.push(textNodes(match[1], entry.name));
  return strings;
}

function worksheetCells(worksheetXml, entriesByName, context) {
  const strings = sharedStrings(entriesByName);
  const cells = new Map();
  const pattern = /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c\b([^>]*?)(?<!\/)>([\s\S]*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?c\s*>/g;
  let match;
  while ((match = pattern.exec(worksheetXml)) !== null) {
    const rawOpening = `<c${match[1]}>`;
    const attributes = parseAttributes(rawOpening, 'c', context);
    const ref = attributes.r;
    if (!ref || !CELL_REFERENCE.test(ref) || cells.has(ref)) {
      throw new Error(`Malformed or duplicate worksheet cell in ${context}.`);
    }
    const body = match[2];
    let value = '';
    if (attributes.t === 'inlineStr') {
      value = textNodes(body, context);
    } else {
      const valueMatch = /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?v\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?v\s*>/.exec(body);
      const rawValue = valueMatch ? decodeXml(valueMatch[1], context) : '';
      if (attributes.t === 's' && rawValue !== '') {
        if (!/^(?:0|[1-9][0-9]*)$/.test(rawValue) || Number(rawValue) >= strings.length) {
          throw new Error(`Invalid shared-string index in ${context}.`);
        }
        value = strings[Number(rawValue)];
      } else {
        value = rawValue;
      }
    }
    cells.set(ref.replace(/\$/g, ''), value);
  }
  return cells;
}

function cellValue(cells, column, row) {
  return String(cells.get(`${column}${row}`) || '');
}

function markerRow(cells, label) {
  const matches = [];
  for (const [ref, value] of cells) {
    const match = /^A([1-9][0-9]*)$/.exec(ref);
    if (match && value === label) matches.push(Number(match[1]));
  }
  if (matches.length !== 1) {
    throw new Error(`START_HERE must contain exactly one ${label} section marker.`);
  }
  return matches[0];
}

function assertHeaderRow(cells, row, expected, label) {
  expected.forEach((value, index) => {
    const column = String.fromCharCode(65 + index);
    if (cellValue(cells, column, row) !== value) {
      throw new Error(`START_HERE ${label} header does not match the exporter contract.`);
    }
  });
}

function locationParts(location) {
  const match = typeof location === 'string' ? NAVIGATION_LOCATION.exec(location) : null;
  if (!match || !CELL_REFERENCE.test(match[1])) return null;
  return {
    gid: /^gid=([0-9]+)/.exec(location)[1],
    range: match[1],
  };
}

function startHereAllowedLinks(worksheet, entriesByName, declaredSheetOrder, anchorTarget) {
  const declaredSheetNames = new Set(declaredSheetOrder);
  const cells = worksheetCells(worksheet.worksheetXml, entriesByName, 'START_HERE');
  const reviewMarker = markerRow(cells, 'REVIEW FIRST');
  const directoryMarker = markerRow(cells, 'WORKBOOK DIRECTORY');
  if (directoryMarker <= reviewMarker + 3) {
    throw new Error('START_HERE review and directory sections are malformed.');
  }
  assertHeaderRow(
    cells,
    reviewMarker + 1,
    ['Severity', 'Fact', 'Tab', 'Campaign ID', 'Campaign', 'Cost', 'Conversions', 'Detail'],
    'review',
  );
  assertHeaderRow(
    cells,
    directoryMarker + 1,
    ['Group', 'Tab', 'Status', 'Rows', 'Purpose', 'Row grain', 'Date range', 'Recommended use'],
    'directory',
  );

  const reviewRows = new Set();
  for (let row = reviewMarker + 2; row <= directoryMarker - 2; row += 1) {
    if (cellValue(cells, 'C', row)) reviewRows.add(row);
  }
  const directoryRows = new Set();
  const directoryTabs = [];
  for (let row = directoryMarker + 2; ; row += 1) {
    const tab = cellValue(cells, 'B', row);
    if (!tab) break;
    directoryRows.add(row);
    directoryTabs.push(tab);
  }
  const declaredDirectoryTabs = declaredSheetOrder.filter((name) => name !== 'START_HERE');
  const declaredDirectorySet = new Set(declaredDirectoryTabs);
  const expectedDirectoryTabs = START_HERE_DIRECTORY_ORDER.filter(
    (name) => declaredDirectorySet.has(name),
  );
  if (expectedDirectoryTabs.length !== declaredDirectoryTabs.length) {
    throw new Error('START_HERE directory contains a sheet outside the exporter topology.');
  }
  if (directoryTabs.length !== expectedDirectoryTabs.length ||
      directoryTabs.some((tab, index) => tab !== expectedDirectoryTabs[index]) ||
      new Set(directoryTabs).size !== directoryTabs.length) {
    throw new Error(
      'START_HERE directory topology must list every declared sheet exactly once in generated order.',
    );
  }

  const linksByRef = new Map();
  for (const link of worksheet.links) {
    if (!linksByRef.has(link.ref)) linksByRef.set(link.ref, []);
    linksByRef.get(link.ref).push(link);
  }
  for (const row of directoryRows) {
    const links = linksByRef.get(`B${row}`) || [];
    const location = links.length === 1 ? locationParts(links[0].location) : null;
    if (links.length !== 1 || links[0].target !== anchorTarget ||
        !location || location.range !== 'A1') {
      throw new Error(`START_HERE directory row ${row} is missing its exact navigation link.`);
    }
  }
  const allowed = new Set();
  for (const link of worksheet.links) {
    if (link.target !== anchorTarget || (linksByRef.get(link.ref) || []).length !== 1) continue;
    const refMatch = /^([BCE])([1-9][0-9]*)$/.exec(link.ref);
    const location = locationParts(link.location);
    if (!refMatch || !location) continue;
    const column = refMatch[1];
    const row = Number(refMatch[2]);
    if (column === 'B' && directoryRows.has(row) && location.range === 'A1' &&
        declaredSheetNames.has(cellValue(cells, 'B', row))) {
      allowed.add(link);
      continue;
    }
    if (column === 'C' && reviewRows.has(row) && location.range === 'A1' &&
        declaredSheetNames.has(cellValue(cells, 'C', row))) {
      allowed.add(link);
      continue;
    }
    if (column === 'E' && reviewRows.has(row) && /^A[1-9][0-9]*$/.test(location.range) &&
        cellValue(cells, 'D', row) && cellValue(cells, 'E', row) &&
        declaredSheetNames.has(cellValue(cells, 'C', row))) {
      const source = (linksByRef.get(`C${row}`) || []).find((candidate) => {
        const sourceLocation = locationParts(candidate.location);
        return candidate.target === anchorTarget && sourceLocation &&
          sourceLocation.range === 'A1' && sourceLocation.gid === location.gid;
      });
      if (source) allowed.add(link);
    }
  }
  return allowed;
}

function workbookSheetMap(entriesByName) {
  const workbookEntry = entriesByName.get('xl/workbook.xml');
  const workbookRelsEntry = entriesByName.get('xl/_rels/workbook.xml.rels');
  if (!workbookEntry || !workbookRelsEntry) {
    throw new Error('XLSX package is missing workbook metadata or relationships.');
  }
  const workbookXml = asText(workbookEntry.data, workbookEntry.name);
  const workbookRoot = assertRootNamespace(
    workbookXml,
    'workbook',
    SPREADSHEET_NS,
    workbookEntry.name,
    { sheets: SPREADSHEET_NS, sheet: SPREADSHEET_NS },
  );
  const relationships = parseRelationships(
    asText(workbookRelsEntry.data, workbookRelsEntry.name),
    workbookRelsEntry.name,
  );
  const result = new Map();
  const sheetNames = new Set();
  for (const raw of matchingTags(workbookXml, 'sheet', workbookEntry.name)) {
    const attributes = parseAttributes(raw, 'sheet', workbookEntry.name);
    const name = attributes.name;
    const idNames = Object.keys(attributes).filter((key) => /:id$/.test(key));
    if (!name || idNames.length !== 1) throw new Error('Malformed workbook sheet declaration.');
    const relationshipPrefix = idNames[0].split(':')[0];
    if (workbookRoot.attributes[`xmlns:${relationshipPrefix}`] !== DOCUMENT_RELATIONSHIP_NS) {
      throw new Error('Workbook sheet relationship uses the wrong XML namespace.');
    }
    if (sheetNames.has(name)) throw new Error(`Duplicate workbook sheet name ${name}.`);
    sheetNames.add(name);
    const relationship = relationships.get(attributes[idNames[0]]);
    if (!relationship) throw new Error(`Workbook sheet ${name} has a missing relationship.`);
    if (relationship.attributes.Type !== WORKSHEET_RELATIONSHIP) {
      throw new Error(`Workbook sheet ${name} has the wrong relationship type.`);
    }
    if (relationship.attributes.TargetMode === 'External') {
      throw new Error(`Workbook sheet ${name} has an external worksheet relationship.`);
    }
    const worksheetName = resolveRelationshipTarget(
      workbookEntry.name,
      relationship.attributes.Target,
      `workbook sheet ${name}`,
    );
    if (!entriesByName.has(worksheetName)) {
      throw new Error(`Workbook sheet ${name} is missing worksheet data.`);
    }
    if (result.has(worksheetName)) throw new Error('Multiple sheets reference one worksheet entry.');
    result.set(worksheetName, name);
  }
  return result;
}

function validateEntries(entries) {
  if (!Array.isArray(entries) || entries.length === 0) {
    throw new TypeError('Workbook entries must be a non-empty array.');
  }
  const byName = new Map();
  for (const entry of entries) {
    if (!entry || typeof entry.name !== 'string' || !entry.name ||
        (!Buffer.isBuffer(entry.data) && !(entry.data instanceof Uint8Array) &&
          typeof entry.data !== 'string')) {
      throw new TypeError('Every workbook entry needs a name and string or Buffer data.');
    }
    const partName = entry.name.endsWith('/') ? entry.name.slice(0, -1) : entry.name;
    const components = partName.split('/');
    if (!partName || entry.name.includes('\\') || entry.name.startsWith('/') ||
        /^[A-Za-z]:/.test(entry.name) || /[\x00-\x1f\x7f]/.test(entry.name) ||
        components.some((component) => !component || component === '.' || component === '..') ||
        path.posix.normalize(entry.name) !== entry.name || entry.name.startsWith('../')) {
      throw new Error('XLSX package contains an unsafe entry name.');
    }
    if (byName.has(entry.name)) throw new Error(`Duplicate XLSX entry ${entry.name}.`);
    byName.set(entry.name, entry);
  }
  return byName;
}

function collectWorkbookLinks(entries) {
  const entriesByName = validateEntries(entries);
  const contentTypes = validateOpcSkeleton(entriesByName);
  const sheetMap = workbookSheetMap(entriesByName);
  for (const worksheetName of sheetMap.keys()) {
    if (contentTypes.get(worksheetName) !== WORKSHEET_CONTENT_TYPE) {
      throw new Error(`Worksheet ${worksheetName} has a missing or incorrect content type.`);
    }
  }
  const worksheets = [];

  for (const [worksheetName, sheetName] of sheetMap) {
    const worksheetEntry = entriesByName.get(worksheetName);
    const worksheetXml = asText(worksheetEntry.data, worksheetName);
    const worksheetRoot = assertRootNamespace(
      worksheetXml,
      'worksheet',
      SPREADSHEET_NS,
      worksheetName,
      {
        sheetData: SPREADSHEET_NS,
        row: SPREADSHEET_NS,
        c: SPREADSHEET_NS,
        f: SPREADSHEET_NS,
        v: SPREADSHEET_NS,
        is: SPREADSHEET_NS,
        t: SPREADSHEET_NS,
        hyperlinks: SPREADSHEET_NS,
        hyperlink: SPREADSHEET_NS,
      },
    );
    assertNoHyperlinkFormulas(worksheetXml, sheetName);
    const rawLinks = matchingTags(worksheetXml, 'hyperlink', sheetName);
    const relsName = relationshipEntryName(worksheetName);
    const relsEntry = entriesByName.get(relsName);
    if (rawLinks.length && !relsEntry) {
      throw new Error(`Hyperlinks in ${sheetName} have no relationship file.`);
    }
    const relsXml = relsEntry ? asText(relsEntry.data, relsName) : '';
    const relationships = relsEntry ? parseRelationships(relsXml, relsName) : new Map();
    const referencedRelationshipIds = new Set();
    const links = [];

    for (const raw of rawLinks) {
      const attributes = parseAttributes(raw, 'hyperlink', sheetName);
      const ref = attributes.ref;
      if (!ref || !CELL_OR_RANGE_REFERENCE.test(ref)) {
        throw new Error(`Malformed hyperlink cell reference in ${sheetName}.`);
      }
      const id = relationshipId(attributes, sheetName);
      const relationshipPrefix = relationshipAttributeName(attributes, sheetName).split(':')[0];
      if (worksheetRoot.attributes[`xmlns:${relationshipPrefix}`] !== DOCUMENT_RELATIONSHIP_NS) {
        throw new Error(`Hyperlink ${ref} in ${sheetName} uses the wrong XML namespace.`);
      }
      const relationship = relationships.get(id);
      if (!relationship) {
        throw new Error(`Hyperlink ${ref} in ${sheetName} has a missing relationship.`);
      }
      if (relationship.attributes.Type !== HYPERLINK_RELATIONSHIP) {
        throw new Error(`Hyperlink ${ref} in ${sheetName} has the wrong relationship type.`);
      }
      if (relationship.attributes.TargetMode !== 'External') {
        throw new Error(`Hyperlink ${ref} in ${sheetName} is not an external relationship.`);
      }
      referencedRelationshipIds.add(id);
      links.push({
        raw,
        attributes,
        id,
        ref,
        location: attributes.location,
        target: relationship.attributes.Target,
        relationship,
      });
    }

    for (const [id, relationship] of relationships) {
      if (relationship.attributes.Type === HYPERLINK_RELATIONSHIP &&
          !referencedRelationshipIds.has(id)) {
        throw new Error(`Orphaned hyperlink relationship ${id} in ${sheetName}.`);
      }
    }

    worksheets.push({
      worksheetName,
      worksheetEntry,
      worksheetXml,
      sheetName,
      relsName,
      relsEntry,
      relsXml,
      relationships,
      links,
    });
  }

  for (const entry of entries) {
    if (/^xl\/worksheets\/[^/]+\.xml$/.test(entry.name) && !sheetMap.has(entry.name)) {
      const xml = asText(entry.data, entry.name);
      assertRootNamespace(xml, 'worksheet', SPREADSHEET_NS, entry.name, {
        sheetData: SPREADSHEET_NS,
        row: SPREADSHEET_NS,
        c: SPREADSHEET_NS,
        f: SPREADSHEET_NS,
        v: SPREADSHEET_NS,
        is: SPREADSHEET_NS,
        t: SPREADSHEET_NS,
        hyperlinks: SPREADSHEET_NS,
        hyperlink: SPREADSHEET_NS,
      });
      assertNoHyperlinkFormulas(xml, entry.name);
      if (matchingTags(xml, 'hyperlink', entry.name).length) {
        throw new Error(`Orphaned worksheet ${entry.name} contains hyperlinks.`);
      }
    }
  }

  const declaredWorksheetRels = new Set(
    worksheets.map((worksheet) => worksheet.relsName),
  );
  for (const entry of entries) {
    if (!entry.name.endsWith('.rels') || declaredWorksheetRels.has(entry.name)) continue;
    const relationships = parseRelationships(asText(entry.data, entry.name), entry.name);
    for (const relationship of relationships.values()) {
      if (relationship.attributes.Type === HYPERLINK_RELATIONSHIP) {
        throw new Error(`Hyperlink relationship exists outside a declared worksheet in ${entry.name}.`);
      }
    }
  }

  const infoWorksheet = worksheets.find((worksheet) => worksheet.sheetName === '_export_info');
  const startWorksheet = worksheets.find((worksheet) => worksheet.sheetName === 'START_HERE');
  if (!infoWorksheet || !startWorksheet) {
    throw new Error('Exporter XLSX must contain START_HERE and _export_info worksheets.');
  }
  const anchors = infoWorksheet.links.filter((link) => link.ref === 'H5');
  if (anchors.length !== 1) {
    throw new Error('Exporter XLSX must contain exactly one _export_info H5 trust anchor.');
  }
  const anchor = anchors[0];
  if (anchor.location !== undefined || !CANONICAL_GOOGLE_SHEET.test(anchor.target)) {
    throw new Error('_export_info H5 trust anchor is not a canonical Google Sheet link.');
  }
  const infoCells = worksheetCells(infoWorksheet.worksheetXml, entriesByName, '_export_info');
  if (cellValue(infoCells, 'H', 5) !== anchor.target) {
    throw new Error('_export_info H5 trust anchor target does not match its cell value.');
  }
  const declaredSheetOrder = Array.from(sheetMap.values());
  const allowedStartLinks = startHereAllowedLinks(
    startWorksheet,
    entriesByName,
    declaredSheetOrder,
    anchor.target,
  );
  for (const worksheet of worksheets) {
    for (const link of worksheet.links) {
      link.allowed = link === anchor || allowedStartLinks.has(link);
    }
  }
  return { entriesByName, worksheets };
}

function auditWorkbookEntries(entries) {
  const collected = collectWorkbookLinks(entries);
  const report = {
    total: 0,
    allowed: 0,
    unexpected: 0,
    allowedBySheet: {},
    unexpectedBySheet: {},
  };
  for (const worksheet of collected.worksheets) {
    for (const link of worksheet.links) {
      report.total += 1;
      const bucket = link.allowed ? 'allowed' : 'unexpected';
      const sheetBucket = link.allowed ? 'allowedBySheet' : 'unexpectedBySheet';
      report[bucket] += 1;
      report[sheetBucket][worksheet.sheetName] =
        (report[sheetBucket][worksheet.sheetName] || 0) + 1;
    }
  }
  return report;
}

function sheetDataSnapshot(xml) {
  const match = /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?sheetData\b[\s\S]*?<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?sheetData\s*>/.exec(xml) ||
    /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?sheetData\b[^<>]*\/>/.exec(xml);
  if (!match) throw new Error('Worksheet is missing sheetData.');
  return match[0];
}

function removeEmptyHyperlinksContainer(xml) {
  return xml.replace(
    /<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?hyperlinks\b[^>]*>\s*<\/(?:[A-Za-z_][A-Za-z0-9_.-]*:)?hyperlinks\s*>/g,
    '',
  );
}

function assertAllSheetDataPreserved(beforeEntries, afterEntries) {
  const afterByName = new Map(afterEntries.map((entry) => [entry.name, entry]));
  for (const entry of beforeEntries) {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(entry.name)) continue;
    const after = afterByName.get(entry.name);
    if (!after || sheetDataSnapshot(asText(entry.data, entry.name)) !==
        sheetDataSnapshot(asText(after.data, after.name))) {
      throw new Error(`Worksheet sheetData changed in ${entry.name}.`);
    }
  }
}

function nonHyperlinkRelationshipNodes(entry) {
  const relationships = parseRelationships(asText(entry.data, entry.name), entry.name);
  return Array.from(relationships.values())
    .filter((relationship) => relationship.attributes.Type !== HYPERLINK_RELATIONSHIP)
    .map((relationship) => relationship.raw);
}

function assertNonHyperlinkRelationshipsPreserved(beforeEntries, afterEntries) {
  const afterByName = new Map(afterEntries.map((entry) => [entry.name, entry]));
  for (const entry of beforeEntries) {
    if (!entry.name.endsWith('.rels')) continue;
    const after = afterByName.get(entry.name);
    if (!after || JSON.stringify(nonHyperlinkRelationshipNodes(entry)) !==
        JSON.stringify(nonHyperlinkRelationshipNodes(after))) {
      throw new Error(`Non-hyperlink relationships changed in ${entry.name}.`);
    }
  }
}

function sanitizeWorkbookEntries(entries) {
  const collected = collectWorkbookLinks(entries);
  const before = auditWorkbookEntries(entries);
  const replacements = new Map();

  for (const worksheet of collected.worksheets) {
    const unexpected = worksheet.links.filter((link) => !link.allowed);
    if (!unexpected.length) continue;
    const beforeSheetData = sheetDataSnapshot(worksheet.worksheetXml);
    let worksheetXml = worksheet.worksheetXml;
    for (const link of unexpected) worksheetXml = worksheetXml.replace(link.raw, '');
    worksheetXml = removeEmptyHyperlinksContainer(worksheetXml);
    if (sheetDataSnapshot(worksheetXml) !== beforeSheetData) {
      throw new Error(`Sanitization changed sheetData in ${worksheet.sheetName}.`);
    }
    replacements.set(
      worksheet.worksheetName,
      withText(worksheet.worksheetEntry, worksheetXml),
    );

    const retainedIds = new Set(
      worksheet.links.filter((link) => link.allowed).map((link) => link.id),
    );
    let relsXml = worksheet.relsXml;
    for (const link of unexpected) {
      if (!retainedIds.has(link.id)) relsXml = relsXml.replace(link.relationship.raw, '');
    }
    replacements.set(worksheet.relsName, withText(worksheet.relsEntry, relsXml));
  }

  const output = entries.map((entry) => replacements.get(entry.name) || { ...entry });
  assertAllSheetDataPreserved(entries, output);
  assertNonHyperlinkRelationshipsPreserved(entries, output);
  const after = auditWorkbookEntries(output);
  if (after.unexpected !== 0 || after.allowed !== before.allowed ||
      after.total !== before.allowed) {
    throw new Error('Sanitized workbook failed its post-write hyperlink audit.');
  }
  return {
    entries: output,
    report: {
      ...before,
      removed: before.unexpected,
      remaining: after.total,
      sheetDataVerified: true,
      nonHyperlinkRelationshipsVerified: true,
    },
  };
}

let crcTable;
function crc32(buffer) {
  if (!crcTable) {
    crcTable = Array.from({ length: 256 }, (_, index) => {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) {
        value = (value >>> 1) ^ ((value & 1) ? 0xedb88320 : 0);
      }
      return value >>> 0;
    });
  }
  let value = 0xffffffff;
  for (const byte of buffer) value = (value >>> 8) ^ crcTable[(value ^ byte) & 0xff];
  return (value ^ 0xffffffff) >>> 0;
}

function findEndOfCentralDirectory(buffer) {
  const first = Math.max(0, buffer.length - 0xffff - 22);
  for (let offset = buffer.length - 22; offset >= first; offset -= 1) {
    if (buffer.readUInt32LE(offset) !== ZIP_END_OF_CENTRAL_DIRECTORY) continue;
    const commentLength = buffer.readUInt16LE(offset + 20);
    if (offset + 22 + commentLength === buffer.length) return offset;
  }
  throw new Error('Input is not a readable ZIP/XLSX package.');
}

function decodeZipName(bytes, flags) {
  if ((flags & 0x0800) === 0 && Array.from(bytes).some((byte) => byte > 0x7f)) {
    throw new Error('ZIP entry name is non-ASCII without the UTF-8 flag.');
  }
  try {
    return fatalUtf8.decode(bytes);
  } catch (error) {
    throw new Error('ZIP entry name is not valid UTF-8.');
  }
}

function parseZipExtraFields(extra, context) {
  let offset = 0;
  while (offset < extra.length) {
    if (offset + 4 > extra.length) throw new Error(`Malformed ZIP extra metadata in ${context}.`);
    const id = extra.readUInt16LE(offset);
    const size = extra.readUInt16LE(offset + 2);
    offset += 4;
    if (offset + size > extra.length) throw new Error(`Malformed ZIP extra metadata in ${context}.`);
    if (id === 0x0001) throw new Error(`ZIP64 extra metadata is not supported in ${context}.`);
    offset += size;
  }
}

function assertSupportedZipFlags(flags, compressionMethod, context) {
  const allowed = 0x0808 | (compressionMethod === 8 ? 0x0006 : 0);
  if ((flags & ~allowed) !== 0 || (compressionMethod === 0 && (flags & 0x0006) !== 0)) {
    throw new Error(`Unsupported ZIP flags in ${context}.`);
  }
}

function assertSafeZipEntryType(versionMadeBy, externalAttributes, name) {
  const creatorSystem = versionMadeBy >>> 8;
  const mode = (externalAttributes >>> 16) & 0xffff;
  const type = mode & 0o170000;
  if (type === 0o120000) throw new Error(`ZIP entry ${name} is an unsafe Unix symlink.`);
  if (type !== 0 && type !== 0o100000 && type !== 0o040000) {
    throw new Error(`ZIP entry ${name} has an unsafe non-regular Unix type.`);
  }
  if (creatorSystem === 3 && (type === 0o040000) !== name.endsWith('/') && type !== 0) {
    throw new Error(`ZIP entry ${name} has contradictory Unix file-type metadata.`);
  }
}

function readZipEntries(filePath) {
  const archiveSize = fs.statSync(filePath).size;
  if (archiveSize > MAX_ARCHIVE_BYTES) {
    throw new Error('Input XLSX exceeds the supported archive size limit.');
  }
  const archive = fs.readFileSync(filePath);
  const eocd = findEndOfCentralDirectory(archive);
  const disk = archive.readUInt16LE(eocd + 4);
  const centralDisk = archive.readUInt16LE(eocd + 6);
  const entriesOnDisk = archive.readUInt16LE(eocd + 8);
  const entryCount = archive.readUInt16LE(eocd + 10);
  const centralSize = archive.readUInt32LE(eocd + 12);
  const centralOffset = archive.readUInt32LE(eocd + 16);
  if (disk !== 0 || centralDisk !== 0 || entriesOnDisk !== entryCount) {
    throw new Error('Multi-disk ZIP files are not supported.');
  }
  if (entryCount === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) {
    throw new Error('ZIP64 XLSX files are not supported.');
  }
  if (centralOffset + centralSize > eocd) throw new Error('Invalid ZIP central directory.');

  const entries = [];
  const names = new Set();
  const localIntervals = [];
  let totalUncompressedSize = 0;
  let offset = centralOffset;
  for (let index = 0; index < entryCount; index += 1) {
    if (offset + 46 > archive.length ||
        archive.readUInt32LE(offset) !== ZIP_CENTRAL_DIRECTORY_HEADER) {
      throw new Error('Invalid ZIP central directory entry.');
    }
    const versionMadeBy = archive.readUInt16LE(offset + 4);
    const versionNeeded = archive.readUInt16LE(offset + 6);
    const flags = archive.readUInt16LE(offset + 8);
    const compressionMethod = archive.readUInt16LE(offset + 10);
    const modTime = archive.readUInt16LE(offset + 12);
    const modDate = archive.readUInt16LE(offset + 14);
    const expectedCrc = archive.readUInt32LE(offset + 16);
    const compressedSize = archive.readUInt32LE(offset + 20);
    const uncompressedSize = archive.readUInt32LE(offset + 24);
    const nameLength = archive.readUInt16LE(offset + 28);
    const extraLength = archive.readUInt16LE(offset + 30);
    const commentLength = archive.readUInt16LE(offset + 32);
    const diskStart = archive.readUInt16LE(offset + 34);
    const internalAttributes = archive.readUInt16LE(offset + 36);
    const externalAttributes = archive.readUInt32LE(offset + 38);
    const localOffset = archive.readUInt32LE(offset + 42);
    const end = offset + 46 + nameLength + extraLength + commentLength;
    if (end > centralOffset + centralSize || diskStart !== 0) {
      throw new Error('Invalid ZIP entry metadata.');
    }
    const centralNameBytes = archive.subarray(offset + 46, offset + 46 + nameLength);
    const name = decodeZipName(centralNameBytes, flags);
    if (!name || names.has(name)) throw new Error('ZIP contains an empty or duplicate entry name.');
    names.add(name);
    if (compressionMethod !== 0 && compressionMethod !== 8) {
      throw new Error(`ZIP entry ${name} uses an unsupported compression method.`);
    }
    assertSupportedZipFlags(flags, compressionMethod, name);
    assertSafeZipEntryType(versionMadeBy, externalAttributes, name);
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff ||
        localOffset === 0xffffffff) {
      throw new Error(`ZIP64 entry metadata is not supported for ${name}.`);
    }
    const centralExtraStart = offset + 46 + nameLength;
    parseZipExtraFields(
      archive.subarray(centralExtraStart, centralExtraStart + extraLength),
      `central entry ${name}`,
    );
    if (compressedSize > MAX_COMPRESSED_ENTRY_BYTES ||
        uncompressedSize > MAX_UNCOMPRESSED_ENTRY_BYTES) {
      throw new Error(`ZIP entry ${name} exceeds the supported size limit.`);
    }
    totalUncompressedSize += uncompressedSize;
    if (totalUncompressedSize > MAX_TOTAL_UNCOMPRESSED_BYTES) {
      throw new Error('Input XLSX exceeds the supported uncompressed size limit.');
    }
    if (compressionMethod === 8 && uncompressedSize > 0 &&
        (compressedSize === 0 || uncompressedSize / compressedSize > MAX_COMPRESSION_RATIO)) {
      throw new Error(`ZIP entry ${name} exceeds the supported compression ratio.`);
    }
    if (localOffset + 30 > centralOffset ||
        archive.readUInt32LE(localOffset) !== ZIP_LOCAL_FILE_HEADER) {
      throw new Error(`ZIP entry ${name} has an invalid local header.`);
    }
    const localFlags = archive.readUInt16LE(localOffset + 6);
    const localMethod = archive.readUInt16LE(localOffset + 8);
    const localCrc = archive.readUInt32LE(localOffset + 14);
    const localCompressedSize = archive.readUInt32LE(localOffset + 18);
    const localUncompressedSize = archive.readUInt32LE(localOffset + 22);
    const localNameLength = archive.readUInt16LE(localOffset + 26);
    const localExtraLength = archive.readUInt16LE(localOffset + 28);
    const localNameStart = localOffset + 30;
    const localNameEnd = localNameStart + localNameLength;
    const localExtraEnd = localNameEnd + localExtraLength;
    if (localExtraEnd > centralOffset ||
        !archive.subarray(localNameStart, localNameEnd).equals(centralNameBytes) ||
        localMethod !== compressionMethod || localFlags !== flags) {
      throw new Error(`ZIP entry ${name} local metadata does not match its central entry.`);
    }
    parseZipExtraFields(
      archive.subarray(localNameEnd, localExtraEnd),
      `local entry ${name}`,
    );
    const hasDataDescriptor = (flags & 0x0008) !== 0;
    if (!hasDataDescriptor && (localCrc !== expectedCrc ||
        localCompressedSize !== compressedSize ||
        localUncompressedSize !== uncompressedSize)) {
      throw new Error(`ZIP entry ${name} local CRC or size metadata differs from its central entry.`);
    }
    if (hasDataDescriptor &&
        ((localCrc !== 0 && localCrc !== expectedCrc) ||
          (localCompressedSize !== 0 && localCompressedSize !== compressedSize) ||
          (localUncompressedSize !== 0 && localUncompressedSize !== uncompressedSize))) {
      throw new Error(`ZIP entry ${name} local placeholder metadata contradicts its central entry.`);
    }
    const dataOffset = localExtraEnd;
    const dataEnd = dataOffset + compressedSize;
    if (dataEnd > centralOffset) {
      throw new Error(`ZIP entry ${name} overlaps the central directory boundary.`);
    }
    let localEnd = dataEnd;
    if (hasDataDescriptor) {
      if (dataEnd + 16 > centralOffset || archive.readUInt32LE(dataEnd) !== 0x08074b50) {
        throw new Error(`ZIP entry ${name} has a missing or unsupported data descriptor.`);
      }
      if (archive.readUInt32LE(dataEnd + 4) !== expectedCrc ||
          archive.readUInt32LE(dataEnd + 8) !== compressedSize ||
          archive.readUInt32LE(dataEnd + 12) !== uncompressedSize) {
        throw new Error(`ZIP entry ${name} data descriptor CRC or sizes do not match.`);
      }
      localEnd += 16;
    }
    localIntervals.push({ start: localOffset, end: localEnd, name });
    const compressed = archive.subarray(dataOffset, dataEnd);
    const data = compressionMethod === 0
      ? Buffer.from(compressed)
      : zlib.inflateRawSync(compressed, {
        maxOutputLength: Math.min(MAX_UNCOMPRESSED_ENTRY_BYTES, uncompressedSize + 1),
      });
    if (data.length !== uncompressedSize || crc32(data) !== expectedCrc) {
      throw new Error(`ZIP entry ${name} failed its integrity check.`);
    }
    entries.push({
      name,
      data,
      compressionMethod,
      versionMadeBy,
      versionNeeded,
      modTime,
      modDate,
      internalAttributes,
      externalAttributes,
    });
    offset = end;
  }
  if (offset !== centralOffset + centralSize) throw new Error('ZIP central directory size mismatch.');
  localIntervals.sort((left, right) => left.start - right.start);
  for (let index = 1; index < localIntervals.length; index += 1) {
    if (localIntervals[index].start < localIntervals[index - 1].end) {
      throw new Error(
        `ZIP local entries ${localIntervals[index - 1].name} and ` +
          `${localIntervals[index].name} overlap.`,
      );
    }
  }
  return entries;
}

function writeZipEntries(filePath, entries) {
  validateEntries(entries);
  if (fs.existsSync(filePath)) {
    throw new Error('Output XLSX already exists; refusing to overwrite it.');
  }
  if (entries.length >= 0xffff) throw new Error('Output XLSX exceeds the supported entry count.');
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;
  let totalUncompressedSize = 0;

  for (const entry of entries) {
    const name = Buffer.from(entry.name, 'utf8');
    const data = typeof entry.data === 'string' ? Buffer.from(entry.data, 'utf8') : Buffer.from(entry.data);
    const method = entry.compressionMethod === 0 || entry.name.endsWith('/') ? 0 : 8;
    const compressed = method === 0 ? data : zlib.deflateRawSync(data, { level: 9 });
    const crc = crc32(data);
    totalUncompressedSize += data.length;
    if (data.length > MAX_UNCOMPRESSED_ENTRY_BYTES ||
        compressed.length > MAX_COMPRESSED_ENTRY_BYTES ||
        totalUncompressedSize > MAX_TOTAL_UNCOMPRESSED_BYTES) {
      throw new Error('Output XLSX exceeds the supported size limit.');
    }
    const flags = 0x0800;
    const modTime = Number.isInteger(entry.modTime) ? entry.modTime : 0;
    const modDate = Number.isInteger(entry.modDate) ? entry.modDate : 0x0021;
    const local = Buffer.alloc(30);
    local.writeUInt32LE(ZIP_LOCAL_FILE_HEADER, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt16LE(modTime, 10);
    local.writeUInt16LE(modDate, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(compressed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);
    localParts.push(local, name, compressed);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(ZIP_CENTRAL_DIRECTORY_HEADER, 0);
    central.writeUInt16LE(Number.isInteger(entry.versionMadeBy) ? entry.versionMadeBy : 0x0314, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(flags, 8);
    central.writeUInt16LE(method, 10);
    central.writeUInt16LE(modTime, 12);
    central.writeUInt16LE(modDate, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(Number.isInteger(entry.internalAttributes) ? entry.internalAttributes : 0, 36);
    central.writeUInt32LE(Number.isInteger(entry.externalAttributes) ? entry.externalAttributes : 0, 38);
    central.writeUInt32LE(localOffset, 42);
    centralParts.push(central, name);
    localOffset += local.length + name.length + compressed.length;
    if (localOffset > MAX_ARCHIVE_BYTES) {
      throw new Error('Output XLSX exceeds the supported archive size limit.');
    }
  }

  const centralDirectory = Buffer.concat(centralParts);
  if (localOffset + centralDirectory.length + 22 > MAX_ARCHIVE_BYTES) {
    throw new Error('Output XLSX exceeds the supported archive size limit.');
  }
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(ZIP_END_OF_CENTRAL_DIRECTORY, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralDirectory.length, 12);
  eocd.writeUInt32LE(localOffset, 16);
  eocd.writeUInt16LE(0, 20);
  const output = Buffer.concat([...localParts, centralDirectory, eocd]);

  const directory = path.dirname(filePath);
  fs.mkdirSync(directory, { recursive: true });
  const temporary = path.join(
    directory,
    `.${path.basename(filePath)}.${process.pid}.${Date.now()}.tmp`,
  );
  try {
    fs.writeFileSync(temporary, output, { flag: 'wx', mode: 0o600 });
    fs.linkSync(temporary, filePath);
  } finally {
    try { fs.unlinkSync(temporary); } catch (error) {
      if (error && error.code !== 'ENOENT') throw error;
    }
  }
}

function assertEntriesMatch(expected, actual) {
  if (expected.length !== actual.length) {
    throw new Error('Reopened XLSX entry count does not match sanitized data.');
  }
  const actualByName = new Map(actual.map((entry) => [entry.name, entry]));
  for (const expectedEntry of expected) {
    const actualEntry = actualByName.get(expectedEntry.name);
    const expectedData = typeof expectedEntry.data === 'string'
      ? Buffer.from(expectedEntry.data, 'utf8')
      : Buffer.from(expectedEntry.data);
    if (!actualEntry || !Buffer.from(actualEntry.data).equals(expectedData)) {
      throw new Error(`Reopened XLSX data mismatch in ${expectedEntry.name}.`);
    }
  }
}

function defaultOutputPath(inputPath) {
  return /\.xlsx$/i.test(inputPath)
    ? inputPath.replace(/\.xlsx$/i, '-sanitized.xlsx')
    : `${inputPath}-sanitized.xlsx`;
}

function redactTargets(message) {
  return String(message || '').replace(/https?:\/\/[^\s"'<>]+/gi, '[redacted-target]');
}

function sha256File(filePath) {
  return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex');
}

function verificationPath(outputPath) {
  return path.join(
    path.dirname(outputPath),
    `.${path.basename(outputPath)}.verified-${process.pid}-${crypto.randomBytes(8).toString('hex')}.tmp.xlsx`,
  );
}

function runCli(argv = process.argv.slice(2), scriptPath = process.argv[1]) {
  if (argv.length < 1 || argv.length > 2) {
    const scriptName = path.basename(scriptPath || 'sanitize-downloaded-xlsx.js');
    throw new Error(`Usage: node ${scriptName} input.xlsx [output.xlsx]`);
  }
  const inputPath = path.resolve(argv[0]);
  const outputPath = path.resolve(argv[1] || defaultOutputPath(argv[0]));
  if (inputPath === outputPath) throw new Error('Input and output XLSX paths must be different.');
  if (!/\.xlsx$/i.test(inputPath) || !/\.xlsx$/i.test(outputPath)) {
    throw new Error('Input and output paths must use the .xlsx extension.');
  }
  if (fs.existsSync(outputPath)) {
    throw new Error('Output XLSX already exists; refusing to overwrite it.');
  }
  const inputDigest = sha256File(inputPath);
  const entries = readZipEntries(inputPath);
  const result = sanitizeWorkbookEntries(entries);
  const temporaryPath = verificationPath(outputPath);
  let outputDigest;
  try {
    writeZipEntries(temporaryPath, result.entries);
    const reopened = readZipEntries(temporaryPath);
    assertEntriesMatch(result.entries, reopened);
    assertAllSheetDataPreserved(entries, reopened);
    assertNonHyperlinkRelationshipsPreserved(entries, reopened);
    const finalAudit = auditWorkbookEntries(reopened);
    if (finalAudit.unexpected !== 0 || finalAudit.allowed !== result.report.allowed ||
        finalAudit.total !== result.report.allowed) {
      throw new Error('Written XLSX failed its reopened hyperlink audit.');
    }
    outputDigest = sha256File(temporaryPath);
    try {
      fs.linkSync(temporaryPath, outputPath);
    } catch (error) {
      if (error && error.code === 'EEXIST') {
        throw new Error('Output XLSX already exists; refusing to overwrite it.');
      }
      throw error;
    }
  } finally {
    try { fs.unlinkSync(temporaryPath); } catch (error) {
      if (error && error.code !== 'ENOENT') throw error;
    }
  }
  console.log(`Sanitized XLSX written: ${path.basename(outputPath)}`);
  console.log(
    `Hyperlinks audited: ${result.report.total}; preserved: ${result.report.allowed}; ` +
      `removed: ${result.report.removed}.`,
  );
  console.log('Final unexpected hyperlinks: 0');
  console.log('Worksheet sheetData verification: PASS');
  console.log('Non-hyperlink relationship preservation: PASS');
  console.log('Reopened ZIP byte-for-byte entry verification: PASS');
  console.log(`Input SHA-256: ${inputDigest}`);
  console.log(`Output SHA-256: ${outputDigest}`);
  return { outputPath, report: result.report, inputDigest, outputDigest };
}

module.exports = {
  auditWorkbookEntries,
  sanitizeWorkbookEntries,
  readZipEntries,
  writeZipEntries,
  runCli,
};

if (require.main === module) {
  try {
    runCli();
  } catch (error) {
    console.error(`XLSX sanitization failed: ${redactTargets(error && error.message)}`);
    process.exitCode = 1;
  }
}
