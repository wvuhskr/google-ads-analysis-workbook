'use strict';

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const SCRIPT_PATH = path.resolve(
  __dirname,
  '..',
  'google-ads-analysis-workbook.js',
);

function loadExporter(overrides = {}) {
  const source = fs.readFileSync(SCRIPT_PATH, 'utf8');
  const context = {
    console,
    Date,
    JSON,
    Math,
    Object,
    RegExp,
    String,
    Number,
    Boolean,
    Array,
    Error,
    Set,
    Map,
    ...overrides,
  };

  vm.createContext(context);
  vm.runInContext(source, context, { filename: SCRIPT_PATH });
  const aliases = {
    assertAdvertiserAccount: 'assertAdvertiserAccountRuntime_',
    createRowBuffer: 'createSafeRowBuffer_',
  };
  const api = new Proxy(Object.create(null), {
    get(_target, property) {
      if (typeof property !== 'string') return undefined;
      const direct = aliases[property] || property;
      if (Object.prototype.hasOwnProperty.call(context, direct)) return context[direct];
      return context[`${property}_`];
    },
  });
  return { context, source, api };
}

function normalizeLinkSegments(text, links = []) {
  return links.map((link) => ({
    start: Number(link.start ?? 0),
    end: Number(link.end ?? text.length),
    url: String(link.url),
  }));
}

function autoLinkForText(text) {
  if (/^https?:\/\/\S+$/i.test(text)) {
    return [{ start: 0, end: text.length, url: text }];
  }
  if (!/^[a-z_][a-z0-9_]*(?:\.[a-z_][a-z0-9_]*)+$/i.test(text)) return [];
  return [{ start: 0, end: text.length, url: `http://${text}` }];
}

function simulatedDownloadedLinks(cell) {
  if (!cell || cell.value === '') return [];
  if (cell.richText && cell.richText.linkDirectiveWasSet) {
    return cell.richText.links.map((link) => ({ ...link }));
  }
  // The live Google Sheets XLSX converter ignored both quotePrefix and
  // setShowHyperlink(false), then inferred links again from displayed text.
  return autoLinkForText(String(cell.value));
}

class PersistentRichTextValue {
  constructor(text, links = [], metadata = {}) {
    this.text = String(text ?? '');
    this.links = normalizeLinkSegments(this.text, links);
    this.linkDirectiveWasSet = Boolean(metadata.linkDirectiveWasSet);
  }

  getText() { return this.text; }

  getLinkUrl(start, end) {
    if (arguments.length === 0) {
      if (this.links.length !== 1) return null;
      const [link] = this.links;
      return link.start === 0 && link.end === this.text.length ? link.url : null;
    }
    const matching = this.links.filter((link) => link.start <= start && link.end >= end);
    if (matching.length !== 1) return null;
    return matching[0].url;
  }

  getRuns() {
    // A single style run can still contain multiple distinct links. This mirrors
    // the API's documented getLinkUrl() ambiguity and forces callers that need a
    // fail-closed answer to inspect substring offsets.
    return [this];
  }

  copy() {
    const builder = new PersistentRichTextBuilder();
    builder.text = this.text;
    builder.links = this.links.map((link) => ({ ...link }));
    builder.linkDirectiveWasSet = true;
    return builder;
  }
}

class PersistentRichTextBuilder {
  constructor(options = {}) {
    this.options = options;
    this.text = '';
    this.links = [];
    this.linkDirectiveWasSet = false;
  }

  setText(value) {
    this.text = String(value ?? '');
    this.links = [];
    return this;
  }

  setLinkUrl(...args) {
    this.linkDirectiveWasSet = true;
    if (args.length === 1) {
      const [url] = args;
      this.links = url === null || url === undefined
        ? []
        : [{ start: 0, end: this.text.length, url: String(url) }];
      return this;
    }
    const [start, end, url] = args;
    const numericStart = Number(start);
    const numericEnd = Number(end);
    this.links = this.links.filter((link) => (
      link.end <= numericStart || link.start >= numericEnd
    ));
    if (url !== null && url !== undefined) {
      this.links.push({ start: numericStart, end: numericEnd, url: String(url) });
    }
    return this;
  }

  build() {
    return new PersistentRichTextValue(this.text, this.links, {
      linkDirectiveWasSet: this.linkDirectiveWasSet,
    });
  }
}

function normalizeCell(input) {
  const spec = input && typeof input === 'object' && !Array.isArray(input) &&
    Object.prototype.hasOwnProperty.call(input, 'value')
    ? input
    : { value: input };
  const value = spec.value ?? '';
  const text = String(value);
  const links = spec.links === undefined ? autoLinkForText(text) : spec.links;
  return {
    value,
    richText: new PersistentRichTextValue(text, links),
    showHyperlink: spec.showHyperlink ?? true,
    numberFormat: 'General',
    quotePrefix: false,
    note: '',
    background: spec.background ?? '#FFFFFF',
    fontColor: spec.fontColor ?? '#1F1F1F',
    fontWeight: spec.fontWeight ?? 'normal',
    verticalAlignment: spec.verticalAlignment ?? 'bottom',
    wrap: spec.wrap ?? false,
  };
}

function columnLabel(column) {
  let value = Number(column);
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return label;
}

class PersistentRange {
  constructor(sheet, row, column, rowCount, columnCount) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.rowCount = rowCount;
    this.columnCount = columnCount;
  }

  eachCell(callback) {
    for (let r = 0; r < this.rowCount; r += 1) {
      for (let c = 0; c < this.columnCount; c += 1) {
        callback(this.sheet.ensureCell(this.row + r, this.column + c), r, c);
      }
    }
    return this;
  }

  getDisplayValues() {
    const rows = [];
    for (let r = 0; r < this.rowCount; r += 1) {
      const row = [];
      for (let c = 0; c < this.columnCount; c += 1) {
        const value = this.sheet.ensureCell(this.row + r, this.column + c).value;
        row.push(value === null || value === undefined ? '' : String(value));
      }
      rows.push(row);
    }
    return rows;
  }

  getRichTextValues() {
    const rows = [];
    for (let r = 0; r < this.rowCount; r += 1) {
      const row = [];
      for (let c = 0; c < this.columnCount; c += 1) {
        const cell = this.sheet.ensureCell(this.row + r, this.column + c);
        row.push(cell.value === '' ? null : cell.richText);
      }
      rows.push(row);
    }
    return rows;
  }

  setRichTextValues(values) {
    if (values.length !== this.rowCount || values.some((row) => row.length !== this.columnCount)) {
      throw new Error('Rich-text matrix does not match range dimensions.');
    }
    return this.eachCell((cell, r, c) => {
      const incoming = values[r][c];
      const text = incoming && typeof incoming.getText === 'function'
        ? incoming.getText()
        : String(incoming && incoming.text !== undefined ? incoming.text : '');
      const directiveWasSet = Boolean(incoming && incoming.linkDirectiveWasSet);
      let links;
      if (this.sheet.options.stickyLinks && cell.richText.links.length) {
        links = cell.richText.links;
      } else if (directiveWasSet) {
        links = incoming.links || [];
      } else {
        links = autoLinkForText(text);
      }
      cell.value = text;
      if (this.sheet.options.richTextClearsQuotePrefix) cell.quotePrefix = false;
      cell.richText = new PersistentRichTextValue(text, links, {
        linkDirectiveWasSet: directiveWasSet,
      });
    });
  }

  setValues(values) {
    if (values.length !== this.rowCount || values.some((row) => row.length !== this.columnCount)) {
      throw new Error('Value matrix does not match range dimensions.');
    }
    return this.eachCell((cell, r, c) => {
      const raw = values[r][c] ?? '';
      const quoted = typeof raw === 'string' && raw.startsWith("'");
      const value = quoted ? raw.slice(1) : raw;
      const text = String(value);
      cell.value = value;
      cell.quotePrefix = quoted;
      cell.richText = new PersistentRichTextValue(
        text,
        quoted ? [] : autoLinkForText(text),
        // quotePrefix is a different cell property from an explicit rich-text
        // link directive. Keeping them distinct reproduces the live converter.
        { linkDirectiveWasSet: false },
      );
    });
  }

  setShowHyperlink(value) {
    return this.eachCell((cell) => { cell.showHyperlink = Boolean(value); });
  }

  setNumberFormat(value) {
    return this.eachCell((cell) => { cell.numberFormat = value; });
  }

  setNote(value) {
    return this.eachCell((cell) => { cell.note = String(value ?? ''); });
  }

  setBackground(value) {
    return this.eachCell((cell) => { cell.background = String(value); });
  }

  setFontColor(value) {
    return this.eachCell((cell) => { cell.fontColor = String(value); });
  }

  setFontWeight(value) {
    return this.eachCell((cell) => { cell.fontWeight = String(value); });
  }

  setVerticalAlignment(value) {
    return this.eachCell((cell) => { cell.verticalAlignment = String(value); });
  }

  setWrap(value) {
    return this.eachCell((cell) => { cell.wrap = Boolean(value); });
  }

  createFilter() {
    const filter = {
      remove: () => { this.sheet.filter = null; },
    };
    this.sheet.filter = filter;
    return filter;
  }

  getRow() { return this.row; }
  getColumn() { return this.column; }
  getNumRows() { return this.rowCount; }
  getNumColumns() { return this.columnCount; }
  getA1Notation() {
    const start = `${columnLabel(this.column)}${this.row}`;
    const end = `${columnLabel(this.column + this.columnCount - 1)}${this.row + this.rowCount - 1}`;
    return start === end ? start : `${start}:${end}`;
  }
}

class PersistentSheet {
  constructor(name, rows, options = {}) {
    this.name = name;
    this.options = options;
    this.cells = rows.map((row) => row.map(normalizeCell));
    this.filter = null;
    this.frozenRows = 0;
    this.frozenColumns = 0;
    this.columnWidths = {};
    this.hiddenColumns = new Set();
    this.rowHeights = {};
  }

  ensureCell(row, column) {
    while (this.cells.length < row) this.cells.push([]);
    const targetRow = this.cells[row - 1];
    while (targetRow.length < column) targetRow.push(normalizeCell(''));
    return targetRow[column - 1];
  }

  getName() { return this.name; }
  getLastRow() {
    for (let row = this.cells.length; row >= 1; row -= 1) {
      if (this.cells[row - 1].some((cell) => cell.value !== '')) return row;
    }
    return 0;
  }
  getLastColumn() {
    return this.cells.reduce((max, row) => {
      for (let column = row.length; column >= 1; column -= 1) {
        if (row[column - 1].value !== '') return Math.max(max, column);
      }
      return max;
    }, 0);
  }
  getRange(row, column, rowCount = 1, columnCount = 1) {
    return new PersistentRange(this, row, column, rowCount, columnCount);
  }
  getFilter() { return this.filter; }
  setFrozenRows(value) { this.frozenRows = value; return this; }
  setFrozenColumns(value) { this.frozenColumns = value; return this; }
  setRowHeight(row, value) { this.rowHeights[row] = value; return this; }
  setColumnWidth(column, value) { this.columnWidths[column] = value; return this; }
  hideColumns(column) { this.hiddenColumns.add(column); return this; }
}

function createPersistentRichTextHarness(options = {}) {
  const spreadsheetApp = {
    flushCount: 0,
    flush() { this.flushCount += 1; },
    newRichTextValue() {
      const builder = new PersistentRichTextBuilder(options);
      if (options.omitBuilderSetLinkUrl) builder.setLinkUrl = undefined;
      return builder;
    },
  };
  if (options.omitFlush) spreadsheetApp.flush = undefined;
  return {
    SpreadsheetApp: spreadsheetApp,
    createSheet(name, rows, sheetOptions = {}) {
      const mergedOptions = { ...options, ...sheetOptions };
      const sheet = new PersistentSheet(name, rows, mergedOptions);
      if (mergedOptions.omitGetRichTextValues) {
        sheet.getRange = function getRange(row, column, rowCount = 1, columnCount = 1) {
          const range = new PersistentRange(this, row, column, rowCount, columnCount);
          range.getRichTextValues = undefined;
          return range;
        };
      }
      return sheet;
    },
    linkedCell(value, links, extra = {}) {
      return { value, links, ...extra };
    },
  };
}

module.exports = {
  loadExporter,
  SCRIPT_PATH,
  createPersistentRichTextHarness,
  PersistentRichTextValue,
  simulatedDownloadedLinks,
};
