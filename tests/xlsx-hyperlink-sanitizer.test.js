'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const zlib = require('node:zlib');
const {
  auditWorkbookEntries,
  sanitizeWorkbookEntries,
  readZipEntries,
  runCli,
} = require('../tools/sanitize-downloaded-xlsx');

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const DOC_REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const SHEET_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const OFFICE_DOCUMENT_RELATIONSHIP = `${DOC_REL_NS}/officeDocument`;
const WORKBOOK_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const WORKSHEET_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const DRAWING_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.drawing+xml';
const WORKBOOK_URL = [
  'https://docs.google.com', 'spreadsheets', 'd', 'example_sheet_id', 'edit',
].join('/');

function relationship(id, type, target, mode = '') {
  return `<Relationship Id="${id}" Type="${type}" Target="${target}"${
    mode ? ` TargetMode="${mode}"` : ''
  }/>`;
}

function relationships(rows) {
  return `<Relationships xmlns="${REL_NS}">${rows.join('')}</Relationships>`;
}

function worksheet(sheetData, links = []) {
  const hyperlinkXml = links.length ? `<hyperlinks>${links.join('')}</hyperlinks>` : '';
  return `<worksheet xmlns="${SHEET_NS}" xmlns:r="${DOC_REL_NS}"><sheetData>${
    sheetData
  }</sheetData>${hyperlinkXml}</worksheet>`;
}

function hyperlink(ref, id, location = '') {
  return `<hyperlink ref="${ref}" r:id="${id}"${
    location ? ` location="${location.replace('&', '&amp;')}"` : ''
  }/>`;
}

function inlineCell(ref, value) {
  return `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`;
}

function inlineRow(number, values) {
  return `<row r="${number}">${Object.entries(values).map(
    ([column, value]) => inlineCell(`${column}${number}`, value),
  ).join('')}</row>`;
}

function startHereSheetData() {
  return [
    inlineRow(15, { A: 'REVIEW FIRST' }),
    inlineRow(16, {
      A: 'Severity', B: 'Fact', C: 'Tab', D: 'Campaign ID', E: 'Campaign',
      F: 'Cost', G: 'Conversions', H: 'Detail',
    }),
    inlineRow(17, {
      A: 'WARN', B: 'Spend with zero conversions', C: 'campaign', D: '123',
      E: 'Campaign A', F: '10', G: '0', H: 'Review',
    }),
    inlineRow(20, { A: 'WORKBOOK DIRECTORY' }),
    inlineRow(21, {
      A: 'Group', B: 'Tab', C: 'Status', D: 'Rows', E: 'Purpose',
      F: 'Row grain', G: 'Date range', H: 'Recommended use',
    }),
    inlineRow(22, { A: 'Performance', B: 'campaign', C: 'OK', D: '1' }),
    inlineRow(23, { A: 'Guidance', B: '_export_info', C: 'OK', D: '1' }),
  ].join('');
}

function contentTypes() {
  const overrides = [
    ['/xl/workbook.xml', WORKBOOK_CONTENT_TYPE],
    ['/xl/worksheets/sheet1.xml', WORKSHEET_CONTENT_TYPE],
    ['/xl/worksheets/sheet2.xml', WORKSHEET_CONTENT_TYPE],
    ['/xl/worksheets/sheet3.xml', WORKSHEET_CONTENT_TYPE],
    ['/xl/drawings/drawing1.xml', DRAWING_CONTENT_TYPE],
    ['/xl/drawings/drawing2.xml', DRAWING_CONTENT_TYPE],
    ['/xl/drawings/drawing3.xml', DRAWING_CONTENT_TYPE],
  ].map(([part, type]) => `<Override PartName="${part}" ContentType="${type}"/>`).join('');
  return `<Types xmlns="${CONTENT_TYPES_NS}">` +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `${overrides}</Types>`;
}

function drawing() {
  return '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>';
}

function fixtureEntries() {
  const hyperlinkType = `${DOC_REL_NS}/hyperlink`;
  return [
    { name: '[Content_Types].xml', data: contentTypes() },
    {
      name: '_rels/.rels',
      data: relationships([
        relationship('rId1', OFFICE_DOCUMENT_RELATIONSHIP, 'xl/workbook.xml'),
      ]),
    },
    {
      name: 'xl/workbook.xml',
      data: `<workbook xmlns="${SHEET_NS}" xmlns:r="${DOC_REL_NS}"><sheets>` +
        '<sheet name="START_HERE" sheetId="1" r:id="rId1"/>' +
        '<sheet name="_export_info" sheetId="2" r:id="rId2"/>' +
        '<sheet name="campaign" sheetId="3" r:id="rId3"/>' +
        '</sheets></workbook>',
    },
    {
      name: 'xl/_rels/workbook.xml.rels',
      data: relationships([
        relationship('rId1', `${DOC_REL_NS}/worksheet`, 'worksheets/sheet1.xml'),
        relationship('rId2', `${DOC_REL_NS}/worksheet`, 'worksheets/sheet2.xml'),
        relationship('rId3', `${DOC_REL_NS}/worksheet`, 'worksheets/sheet3.xml'),
      ]),
    },
    {
      name: 'xl/worksheets/sheet1.xml',
      data: worksheet(startHereSheetData(), [
        hyperlink('C17', 'rId1', 'gid=11&range=A1'),
        hyperlink('E17', 'rId2', 'gid=11&range=A3'),
        hyperlink('B22', 'rId3', 'gid=22&range=A1'),
        hyperlink('B23', 'rId4', 'gid=23&range=A1'),
      ]),
    },
    {
      name: 'xl/worksheets/_rels/sheet1.xml.rels',
      data: relationships([
        relationship('rId1', hyperlinkType, WORKBOOK_URL, 'External'),
        relationship('rId2', hyperlinkType, WORKBOOK_URL, 'External'),
        relationship('rId3', hyperlinkType, WORKBOOK_URL, 'External'),
        relationship('rId4', hyperlinkType, WORKBOOK_URL, 'External'),
        relationship('rId5', `${DOC_REL_NS}/drawing`, '../drawings/drawing1.xml'),
      ]),
    },
    {
      name: 'xl/worksheets/sheet2.xml',
      data: worksheet(inlineRow(5, { G: 'workbook_url', H: WORKBOOK_URL }), [
        hyperlink('H5', 'rId1'),
      ]),
    },
    {
      name: 'xl/worksheets/_rels/sheet2.xml.rels',
      data: relationships([
        relationship('rId1', hyperlinkType, WORKBOOK_URL, 'External'),
        relationship('rId2', `${DOC_REL_NS}/drawing`, '../drawings/drawing2.xml'),
      ]),
    },
    {
      name: 'xl/worksheets/sheet3.xml',
      data: worksheet('<row r="1"><c r="A1" t="inlineStr"><is><t>campaign.id</t></is></c></row>', [
        hyperlink('A1', 'rId1'),
        hyperlink('J2', 'rId2'),
      ]),
    },
    {
      name: 'xl/worksheets/_rels/sheet3.xml.rels',
      data: relationships([
        relationship('rId1', hyperlinkType, 'http://campaign.id', 'External'),
        relationship('rId2', hyperlinkType, 'https://example.test/path', 'External'),
        relationship('rId3', `${DOC_REL_NS}/drawing`, '../drawings/drawing3.xml'),
      ]),
    },
    { name: 'xl/drawings/drawing1.xml', data: drawing() },
    { name: 'xl/drawings/drawing2.xml', data: drawing() },
    { name: 'xl/drawings/drawing3.xml', data: drawing() },
  ];
}

function byName(entries, name) {
  return entries.find((entry) => entry.name === name).data;
}

let crcTable;
function fixtureCrc32(buffer) {
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

function buildIndependentStoredZip(entries, options = {}) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, 'utf8');
    const data = typeof entry.data === 'string' ? Buffer.from(entry.data, 'utf8') : Buffer.from(entry.data);
    const compressed = options.deflate ? zlib.deflateRawSync(data) : data;
    const method = options.deflate ? 8 : 0;
    const dataDescriptor = Boolean(options.dataDescriptor);
    const flags = 0x0800 | (dataDescriptor ? 0x0008 : 0);
    const localExtra = entry.localExtra ? Buffer.from(entry.localExtra) : Buffer.alloc(0);
    const centralExtra = entry.centralExtra ? Buffer.from(entry.centralExtra) : localExtra;
    const crc = fixtureCrc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(flags, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(dataDescriptor ? 0 : crc, 14);
    local.writeUInt32LE(dataDescriptor ? 0 : compressed.length, 18);
    local.writeUInt32LE(dataDescriptor ? 0 : data.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(localExtra.length, 28);
    const descriptor = dataDescriptor ? Buffer.alloc(16) : Buffer.alloc(0);
    if (dataDescriptor) {
      descriptor.writeUInt32LE(0x08074b50, 0);
      descriptor.writeUInt32LE(crc, 4);
      descriptor.writeUInt32LE(compressed.length, 8);
      descriptor.writeUInt32LE(data.length, 12);
    }
    localParts.push(local, name, localExtra, compressed, descriptor);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(entry.versionMadeBy === undefined ? 20 : entry.versionMadeBy, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(flags, 8);
    central.writeUInt16LE(method, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(compressed.length, 20);
    central.writeUInt32LE(data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt16LE(centralExtra.length, 30);
    central.writeUInt32LE(entry.externalAttributes || 0, 38);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, name, centralExtra);
    offset += local.length + name.length + localExtra.length + compressed.length + descriptor.length;
  }
  const centralDirectory = Buffer.concat(centralParts);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralDirectory.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...localParts, centralDirectory, eocd]);
}

function centralDirectoryOffset(archive) {
  return archive.readUInt32LE(archive.length - 6);
}

function firstEntryOffsets(archive) {
  const central = centralDirectoryOffset(archive);
  return { local: archive.readUInt32LE(central + 42), central };
}

function removeInfoAnchor(entries) {
  const clone = entries.map((entry) => ({ ...entry }));
  const sheet = clone.find((entry) => entry.name === 'xl/worksheets/sheet2.xml');
  sheet.data = sheet.data.replace(/<hyperlinks>[\s\S]*?<\/hyperlinks>/, '');
  const rels = clone.find((entry) => entry.name === 'xl/worksheets/_rels/sheet2.xml.rels');
  rels.data = rels.data.replace(
    /<Relationship Id="rId1"[^>]*\/officeDocument\/2006\/relationships\/hyperlink"[^>]*\/>/,
    '',
  );
  return clone;
}

test('audit separates exact workbook navigation from converter-inferred links', () => {
  const report = auditWorkbookEntries(fixtureEntries());

  assert.equal(report.total, 7);
  assert.equal(report.allowed, 5);
  assert.equal(report.unexpected, 2);
  assert.deepEqual(report.unexpectedBySheet, { campaign: 2 });
});

test('navigation trust is anchored only by one canonical _export_info H5 link', () => {
  const withoutAnchor = removeInfoAnchor(fixtureEntries());
  const start = withoutAnchor.find((entry) => entry.name === 'xl/worksheets/sheet1.xml');
  start.data = start.data.replace(
    '</hyperlinks>', `${hyperlink('B999', 'rId9', 'gid=999&range=A1')}</hyperlinks>`,
  );
  const startRels = withoutAnchor.find(
    (entry) => entry.name === 'xl/worksheets/_rels/sheet1.xml.rels',
  );
  startRels.data = startRels.data.replace(
    '</Relationships>',
    `${relationship('rId9', `${DOC_REL_NS}/hyperlink`, WORKBOOK_URL, 'External')}</Relationships>`,
  );
  assert.throws(
    () => auditWorkbookEntries(withoutAnchor),
    /_export_info|trust anchor|anchor.*missing/i,
  );

  const duplicate = fixtureEntries();
  const info = duplicate.find((entry) => entry.name === 'xl/worksheets/sheet2.xml');
  info.data = info.data.replace('</hyperlinks>', `${hyperlink('H5', 'rId9')}</hyperlinks>`);
  const infoRels = duplicate.find(
    (entry) => entry.name === 'xl/worksheets/_rels/sheet2.xml.rels',
  );
  infoRels.data = infoRels.data.replace(
    '</Relationships>',
    `${relationship('rId9', `${DOC_REL_NS}/hyperlink`, WORKBOOK_URL, 'External')}</Relationships>`,
  );
  assert.throws(() => auditWorkbookEntries(duplicate), /duplicate|exactly one|trust anchor/i);

  const noncanonical = fixtureEntries();
  noncanonical.find((entry) => entry.name.endsWith('sheet2.xml.rels')).data =
    noncanonical.find((entry) => entry.name.endsWith('sheet2.xml.rels')).data
      .replace(WORKBOOK_URL, 'https://example.test/workbook');
  assert.throws(() => auditWorkbookEntries(noncanonical), /canonical|trust anchor|_export_info/i);
});

test('START_HERE preserves only links in generated review and directory cells', () => {
  const entries = fixtureEntries();
  const start = entries.find((entry) => entry.name === 'xl/worksheets/sheet1.xml');
  start.data = start.data
    .replace('</sheetData>', `${inlineRow(999, { B: 'campaign' })}</sheetData>`)
    .replace('</hyperlinks>', `${hyperlink('B999', 'rId9', 'gid=999&range=A1')}</hyperlinks>`);
  const rels = entries.find((entry) => entry.name.endsWith('sheet1.xml.rels'));
  rels.data = rels.data.replace(
    '</Relationships>',
    `${relationship('rId9', `${DOC_REL_NS}/hyperlink`, WORKBOOK_URL, 'External')}</Relationships>`,
  );

  const before = auditWorkbookEntries(entries);
  assert.equal(before.allowed, 5);
  assert.equal(before.unexpected, 3);
  const sanitized = sanitizeWorkbookEntries(entries);
  assert.equal(sanitized.report.removed, 3);
  assert.equal(auditWorkbookEntries(sanitized.entries).allowed, 5);
});

test('START_HERE directory topology rejects duplicate, omitted, and unlinked tabs', () => {
  const duplicate = fixtureEntries();
  const duplicateStart = duplicate.find((entry) => entry.name === 'xl/worksheets/sheet1.xml');
  duplicateStart.data = duplicateStart.data.replace('<t>_export_info</t>', '<t>campaign</t>');
  assert.throws(
    () => auditWorkbookEntries(duplicate),
    /directory.*(?:duplicate|topology|sheet set|order)|duplicate.*directory/i,
  );

  const missingLink = fixtureEntries();
  const missingStart = missingLink.find((entry) => entry.name === 'xl/worksheets/sheet1.xml');
  missingStart.data = missingStart.data.replace(hyperlink('B23', 'rId4', 'gid=23&range=A1'), '');
  const missingRels = missingLink.find(
    (entry) => entry.name === 'xl/worksheets/_rels/sheet1.xml.rels',
  );
  missingRels.data = missingRels.data.replace(
    relationship('rId4', `${DOC_REL_NS}/hyperlink`, WORKBOOK_URL, 'External'),
    '',
  );
  assert.throws(
    () => auditWorkbookEntries(missingLink),
    /directory.*(?:missing|link|topology)/i,
  );

  const wrongOrder = fixtureEntries();
  const orderStart = wrongOrder.find((entry) => entry.name === 'xl/worksheets/sheet1.xml');
  orderStart.data = orderStart.data
    .replace('<t>_export_info</t>', '<t>__TEMP__</t>')
    .replace('<t>campaign</t>', '<t>_export_info</t>')
    .replace('<t>__TEMP__</t>', '<t>campaign</t>');
  assert.throws(() => auditWorkbookEntries(wrongOrder), /directory.*(?:order|topology)/i);
});

test('audit requires a valid OPC/XLSX package skeleton and safe part names', () => {
  for (const missing of ['[Content_Types].xml', '_rels/.rels']) {
    const entries = fixtureEntries().filter((entry) => entry.name !== missing);
    assert.throws(() => auditWorkbookEntries(entries), /content.?types|root relationship|OPC|XLSX/i);
  }

  const wrongRoot = fixtureEntries();
  wrongRoot.find((entry) => entry.name === '_rels/.rels').data = relationships([
    relationship('rId1', `${DOC_REL_NS}/customXml`, 'xl/workbook.xml'),
  ]);
  assert.throws(() => auditWorkbookEntries(wrongRoot), /officeDocument|workbook relationship/i);

  const wrongContentType = fixtureEntries();
  wrongContentType.find((entry) => entry.name === '[Content_Types].xml').data =
    wrongContentType.find((entry) => entry.name === '[Content_Types].xml').data
      .replace(WORKSHEET_CONTENT_TYPE, 'application/xml');
  assert.throws(() => auditWorkbookEntries(wrongContentType), /content.?type|worksheet/i);

  for (const unsafeName of [
    '..', 'xl/../evil.xml', 'xl/./evil.xml', 'C:/evil.xml', 'xl/evil\0.xml',
    `xl/control-${String.fromCharCode(1)}.xml`,
  ]) {
    const entries = fixtureEntries().concat({ name: unsafeName, data: 'unsafe' });
    assert.throws(() => auditWorkbookEntries(entries), /unsafe|dot|entry name/i);
  }

  const wrongNamespace = fixtureEntries();
  wrongNamespace.find((entry) => entry.name === '[Content_Types].xml').data =
    wrongNamespace.find((entry) => entry.name === '[Content_Types].xml').data
      .replace(CONTENT_TYPES_NS, 'https://example.test/wrong-content-types-namespace');
  assert.throws(() => auditWorkbookEntries(wrongNamespace), /namespace|content.?types/i);
});

test('XML package parts reject invalid UTF-8 instead of replacement decoding', () => {
  const entries = fixtureEntries();
  entries.find((entry) => entry.name === '[Content_Types].xml').data =
    Buffer.from([0x3c, 0x54, 0x79, 0x70, 0x65, 0x73, 0xc3, 0x28, 0x2f, 0x3e]);
  assert.throws(() => auditWorkbookEntries(entries), /UTF-?8|encoding/i);
});

test('inspected XML must be balanced and every relevant child must use the required namespace', () => {
  const missingClose = fixtureEntries();
  missingClose.find((entry) => entry.name === '[Content_Types].xml').data =
    missingClose.find((entry) => entry.name === '[Content_Types].xml').data.replace('</Types>', '');
  assert.throws(() => auditWorkbookEntries(missingClose), /XML|closing|balanced|Types/i);

  const badOverride = fixtureEntries();
  badOverride.find((entry) => entry.name === '[Content_Types].xml').data =
    badOverride.find((entry) => entry.name === '[Content_Types].xml').data.replace(
      '<Override ',
      '<bad:Override xmlns:bad="urn:wrong" ',
    );
  assert.throws(() => auditWorkbookEntries(badOverride), /namespace|Override|content.?type/i);

  const badRelationship = fixtureEntries();
  badRelationship.find((entry) => entry.name === '_rels/.rels').data =
    badRelationship.find((entry) => entry.name === '_rels/.rels').data.replace(
      '<Relationship ',
      '<bad:Relationship xmlns:bad="urn:wrong" ',
    );
  assert.throws(() => auditWorkbookEntries(badRelationship), /namespace|Relationship/i);
});

test('commented XML elements never participate in package trust or worksheet parsing', () => {
  const commentedRootRelationship = fixtureEntries();
  const rootRels = commentedRootRelationship.find((entry) => entry.name === '_rels/.rels');
  const officeRelationship = relationship(
    'rId1',
    OFFICE_DOCUMENT_RELATIONSHIP,
    'xl/workbook.xml',
  );
  rootRels.data = rootRels.data.replace(officeRelationship, `<!--${officeRelationship}-->`);
  assert.throws(
    () => auditWorkbookEntries(commentedRootRelationship),
    /comment|officeDocument|workbook relationship|missing/i,
  );

  const commentedOverride = fixtureEntries();
  const types = commentedOverride.find((entry) => entry.name === '[Content_Types].xml');
  const workbookOverride =
    `<Override PartName="/xl/workbook.xml" ContentType="${WORKBOOK_CONTENT_TYPE}"/>`;
  types.data = types.data.replace(workbookOverride, `<!--${workbookOverride}-->`);
  assert.throws(
    () => auditWorkbookEntries(commentedOverride),
    /comment|workbook content.?type|missing/i,
  );

  const commentedMarker = fixtureEntries();
  const start = commentedMarker.find((entry) => entry.name === 'xl/worksheets/sheet1.xml');
  const reviewMarker = inlineRow(15, { A: 'REVIEW FIRST' });
  start.data = start.data.replace(reviewMarker, `<!--${reviewMarker}-->`);
  assert.throws(() => auditWorkbookEntries(commentedMarker), /comment|REVIEW FIRST|section marker/i);

  const commentedAnchor = fixtureEntries();
  const info = commentedAnchor.find((entry) => entry.name === 'xl/worksheets/sheet2.xml');
  const anchor = hyperlink('H5', 'rId1');
  info.data = info.data.replace(anchor, `<!--${anchor}-->`);
  assert.throws(() => auditWorkbookEntries(commentedAnchor), /comment|anchor|orphan|_export_info/i);
});

test('sanitizer removes unexpected hyperlinks while preserving navigation and sheet data', () => {
  const original = fixtureEntries();
  const beforeSheetData = original
    .filter((entry) => /^xl\/worksheets\/sheet\d+\.xml$/.test(entry.name))
    .map((entry) => [entry.name, entry.data.match(/<sheetData>[\s\S]*?<\/sheetData>/)[0]]);

  const result = sanitizeWorkbookEntries(original);
  const afterAudit = auditWorkbookEntries(result.entries);

  assert.equal(result.report.removed, 2);
  assert.equal(afterAudit.allowed, 5);
  assert.equal(afterAudit.unexpected, 0);
  assert.match(byName(result.entries, 'xl/worksheets/_rels/sheet1.xml.rels'), /drawing1\.xml/);
  assert.match(byName(result.entries, 'xl/worksheets/_rels/sheet3.xml.rels'), /drawing3\.xml/);
  assert.doesNotMatch(byName(result.entries, 'xl/worksheets/sheet3.xml'), /<hyperlinks?>/);
  assert.doesNotMatch(
    byName(result.entries, 'xl/worksheets/_rels/sheet3.xml.rels'),
    /campaign\.id|example\.test|\/hyperlink/,
  );
  assert.equal(byName(result.entries, 'xl/drawings/drawing1.xml'), drawing());
  for (const [name, sheetData] of beforeSheetData) {
    assert.equal(byName(result.entries, name).match(/<sheetData>[\s\S]*?<\/sheetData>/)[0], sheetData);
  }
});

test('sanitizer fails closed on orphaned, wrong-type, and location-only hyperlinks', () => {
  const orphaned = fixtureEntries();
  orphaned.find((entry) => entry.name.endsWith('sheet3.xml.rels')).data =
    orphaned.find((entry) => entry.name.endsWith('sheet3.xml.rels')).data
      .replace('Id="rId2"', 'Id="rId9"');
  assert.throws(() => sanitizeWorkbookEntries(orphaned), /missing|orphan|relationship/i);

  const wrongType = fixtureEntries();
  wrongType.find((entry) => entry.name.endsWith('sheet3.xml.rels')).data =
    wrongType.find((entry) => entry.name.endsWith('sheet3.xml.rels')).data
      .replace(`${DOC_REL_NS}/hyperlink`, `${DOC_REL_NS}/drawing`);
  assert.throws(() => sanitizeWorkbookEntries(wrongType), /relationship type|hyperlink/i);

  const locationOnly = fixtureEntries();
  locationOnly.find((entry) => entry.name === 'xl/worksheets/sheet3.xml').data =
    locationOnly.find((entry) => entry.name === 'xl/worksheets/sheet3.xml').data
      .replace(' r:id="rId1"', ' location="A1"');
  assert.throws(() => sanitizeWorkbookEntries(locationOnly), /relationship id|location-only/i);
});

test('audit rejects HYPERLINK formulas instead of silently rewriting formulas', () => {
  const entries = fixtureEntries();
  entries.find((entry) => entry.name === 'xl/worksheets/sheet3.xml').data =
    entries.find((entry) => entry.name === 'xl/worksheets/sheet3.xml').data
      .replace('</sheetData>', '<row r="3"><c r="A3"><f>HYPERLINK(&quot;https://example.test&quot;)</f></c></row></sheetData>');

  assert.throws(() => auditWorkbookEntries(entries), /HYPERLINK.*formula|formula.*HYPERLINK/i);
});

test('ZIP reader reconciles local integrity metadata with the central directory', () => {
  const crcMismatch = buildIndependentStoredZip(fixtureEntries());
  const crcOffsets = firstEntryOffsets(crcMismatch);
  crcMismatch.writeUInt32LE(crcMismatch.readUInt32LE(crcOffsets.local + 14) ^ 1, crcOffsets.local + 14);
  const crcPath = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'gads-zip-crc-')), 'bad.xlsx');
  fs.writeFileSync(crcPath, crcMismatch);
  assert.throws(() => readZipEntries(crcPath), /local.*CRC|metadata.*central|integrity/i);

  const sizeMismatch = buildIndependentStoredZip(fixtureEntries());
  const sizeOffsets = firstEntryOffsets(sizeMismatch);
  sizeMismatch.writeUInt32LE(
    sizeMismatch.readUInt32LE(sizeOffsets.local + 18) + 1,
    sizeOffsets.local + 18,
  );
  const sizePath = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'gads-zip-size-')), 'bad.xlsx');
  fs.writeFileSync(sizePath, sizeMismatch);
  assert.throws(() => readZipEntries(sizePath), /local.*size|metadata.*central|integrity/i);
});

test('ZIP reader validates signed data descriptors and rejects unsupported flags', () => {
  const descriptorArchive = buildIndependentStoredZip(fixtureEntries(), { dataDescriptor: true });
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'gads-zip-descriptor-'));
  const validPath = path.join(directory, 'valid.xlsx');
  fs.writeFileSync(validPath, descriptorArchive);
  assert.doesNotThrow(() => readZipEntries(validPath));

  const brokenDescriptor = Buffer.from(descriptorArchive);
  const first = firstEntryOffsets(brokenDescriptor);
  const nameLength = brokenDescriptor.readUInt16LE(first.local + 26);
  const extraLength = brokenDescriptor.readUInt16LE(first.local + 28);
  const compressedSize = brokenDescriptor.readUInt32LE(first.central + 20);
  const descriptorOffset = first.local + 30 + nameLength + extraLength + compressedSize;
  brokenDescriptor.writeUInt32LE(
    brokenDescriptor.readUInt32LE(descriptorOffset + 4) ^ 1,
    descriptorOffset + 4,
  );
  const brokenPath = path.join(directory, 'broken.xlsx');
  fs.writeFileSync(brokenPath, brokenDescriptor);
  assert.throws(() => readZipEntries(brokenPath), /data descriptor|descriptor.*CRC|integrity/i);

  const flagsArchive = buildIndependentStoredZip(fixtureEntries());
  const flagOffsets = firstEntryOffsets(flagsArchive);
  flagsArchive.writeUInt16LE(flagsArchive.readUInt16LE(flagOffsets.local + 6) | 0x0040, flagOffsets.local + 6);
  flagsArchive.writeUInt16LE(
    flagsArchive.readUInt16LE(flagOffsets.central + 8) | 0x0040,
    flagOffsets.central + 8,
  );
  const flagsPath = path.join(directory, 'unsupported-flags.xlsx');
  fs.writeFileSync(flagsPath, flagsArchive);
  assert.throws(() => readZipEntries(flagsPath), /unsupported ZIP flags|flag/i);
});

test('ZIP reader rejects overlap, Unix symlinks, and contradictory ZIP64 metadata', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'gads-zip-safety-'));

  const overlap = buildIndependentStoredZip(fixtureEntries());
  const overlapOffsets = firstEntryOffsets(overlap);
  const oversized = overlapOffsets.central;
  overlap.writeUInt32LE(oversized, overlapOffsets.local + 18);
  overlap.writeUInt32LE(oversized, overlapOffsets.central + 20);
  const overlapPath = path.join(directory, 'overlap.xlsx');
  fs.writeFileSync(overlapPath, overlap);
  assert.throws(() => readZipEntries(overlapPath), /central directory|overlap|boundary/i);

  const symlinkEntries = fixtureEntries();
  symlinkEntries[0] = {
    ...symlinkEntries[0],
    versionMadeBy: (3 << 8) | 20,
    externalAttributes: (0o120777 << 16) >>> 0,
  };
  const symlinkPath = path.join(directory, 'symlink.xlsx');
  fs.writeFileSync(symlinkPath, buildIndependentStoredZip(symlinkEntries));
  assert.throws(() => readZipEntries(symlinkPath), /symlink|non-regular|unsafe.*type/i);

  const disguisedSymlinkEntries = fixtureEntries();
  disguisedSymlinkEntries[0] = {
    ...disguisedSymlinkEntries[0],
    versionMadeBy: 20,
    externalAttributes: (0o120777 << 16) >>> 0,
  };
  const disguisedSymlinkPath = path.join(directory, 'dos-claimed-symlink.xlsx');
  fs.writeFileSync(disguisedSymlinkPath, buildIndependentStoredZip(disguisedSymlinkEntries));
  assert.throws(
    () => readZipEntries(disguisedSymlinkPath),
    /symlink|non-regular|unsafe.*type/i,
  );

  const zip64Extra = Buffer.alloc(12);
  zip64Extra.writeUInt16LE(0x0001, 0);
  zip64Extra.writeUInt16LE(8, 2);
  zip64Extra.writeBigUInt64LE(1n, 4);
  const zip64Entries = fixtureEntries();
  zip64Entries[0] = {
    ...zip64Entries[0],
    localExtra: zip64Extra,
    centralExtra: zip64Extra,
  };
  const zip64Path = path.join(directory, 'contradictory-zip64.xlsx');
  fs.writeFileSync(zip64Path, buildIndependentStoredZip(zip64Entries));
  assert.throws(() => readZipEntries(zip64Path), /ZIP64|extra metadata/i);
});

test('ZIP reader enforces compression ratio for every nonempty deflated entry', () => {
  const entries = fixtureEntries();
  entries.push({ name: 'xl/highly-compressible.bin', data: 'A'.repeat(10000) });
  const archive = buildIndependentStoredZip(entries, { deflate: true });
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'gads-zip-ratio-'));
  const input = path.join(directory, 'ratio.xlsx');
  fs.writeFileSync(input, archive);
  assert.throws(() => readZipEntries(input), /compression ratio/i);
});

test('CLI sanitizes and reopens an independently built XLSX without leaking targets', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'gads-xlsx-sanitizer-'));
  const input = path.join(directory, 'live-download.xlsx');
  const output = path.join(directory, 'live-download-sanitized.xlsx');
  fs.writeFileSync(input, buildIndependentStoredZip(fixtureEntries()));
  const originalBytes = fs.readFileSync(input);
  const messages = [];
  const originalLog = console.log;
  console.log = (message) => messages.push(String(message));
  try {
    const result = runCli([input, output]);
    assert.equal(result.outputPath, output);
  } finally {
    console.log = originalLog;
  }

  assert.deepEqual(fs.readFileSync(input), originalBytes, 'input download was modified');
  const audit = auditWorkbookEntries(readZipEntries(output));
  assert.equal(audit.allowed, 5);
  assert.equal(audit.unexpected, 0);
  assert.match(messages.join('\n'), /Final unexpected hyperlinks:\s*0/i);
  assert.match(messages.join('\n'), /sheetData.*PASS/i);
  assert.match(messages.join('\n'), /non-hyperlink relationship.*PASS/i);
  assert.match(messages.join('\n'), /reopen.*ZIP.*PASS|PASS.*reopen.*ZIP/i);
  assert.match(messages.join('\n'), /SHA-256:\s*[a-f0-9]{64}/i);
  assert.doesNotMatch(messages.join('\n'), new RegExp(directory.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  assert.doesNotMatch(messages.join('\n'), /campaign\.id|example\.test|example_sheet_id/i);
  assert.throws(() => runCli([input, output]), /already exists|refus.*overwrite/i);
  fs.rmSync(directory, { recursive: true, force: true });
});

test('CLI verification failure never publishes or leaves a temporary XLSX', () => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'gads-xlsx-publish-'));
  const input = path.join(directory, 'input.xlsx');
  const output = path.join(directory, 'public.xlsx');
  fs.writeFileSync(input, buildIndependentStoredZip(fixtureEntries()));
  const originalRead = fs.readFileSync;
  fs.readFileSync = function injectedRead(filePath, ...args) {
    const bytes = originalRead.call(fs, filePath, ...args);
    if (String(filePath).includes('.verified-') && Buffer.isBuffer(bytes)) {
      return bytes.subarray(0, Math.max(1, bytes.length - 11));
    }
    return bytes;
  };
  try {
    assert.throws(() => runCli([input, output]), /ZIP|reopen|verification|readable/i);
  } finally {
    fs.readFileSync = originalRead;
  }
  assert.equal(fs.existsSync(output), false, 'unverified final output was published');
  assert.deepEqual(
    fs.readdirSync(directory).sort(),
    ['input.xlsx'],
    'temporary verification files were not cleaned',
  );
  fs.rmSync(directory, { recursive: true, force: true });
});

test('CLI usage names the packaged sanitizer that the user actually ran', () => {
  assert.throws(
    () => runCli([], '/tmp/sanitize-downloaded-xlsx-v1.0.0.js'),
    /Usage: node sanitize-downloaded-xlsx-v1\.0\.0\.js input\.xlsx \[output\.xlsx\]/,
  );
});
