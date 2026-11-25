// fake-gas.js - loads the Apps Script files in src/ into a Node vm context with
// in-memory stand-ins for SpreadsheetApp, PropertiesService, UrlFetchApp and friends.
// The spreadsheet is built from sample-data/*.tsv. Nothing here talks to a network.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

const ROOT = path.join(__dirname, '..');
const SRC_FILES = ['Config.js', 'TestCaseGenerator.js', 'YouTrackSchemaManager.js', 'YouTrackAPI.js', 'Main.js', 'UIManager.js'];

function readTsv(file) {
  const lines = fs.readFileSync(file, 'utf8').replace(/\r/g, '').split('\n').filter(l => l !== '');
  // Sheets returns checkboxes as booleans and empty cells as ''.
  return lines.map(l => l.split('\t').map(v => (v === 'TRUE' ? true : v === 'FALSE' ? false : v)));
}

class FakeSheet {
  constructor(name, rows, log) { this.name = name; this.rows = rows; this.log = log; }
  getName() { return this.name; }
  getLastRow() { return this.rows.length; }
  getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
  getMaxRows() { return Math.max(this.rows.length, 100); }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
  getRange(row, col, numRows = 1, numCols = 1) {
    if (typeof row === 'string') return new FakeRange(this, 2, 1, 1, 1); // A1 notation: only used for validation sources
    return new FakeRange(this, row, col, numRows, numCols);
  }
  cell(r, c) { const row = this.rows[r - 1]; return row && row[c - 1] !== undefined ? row[c - 1] : ''; }
  set(r, c, v) {
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
    this.log.push({ sheet: this.name, row: r, col: c, value: v });
  }
  clear() { this.rows = []; this.log.push({ sheet: this.name, clear: true }); }
  appendRow(values) { this.rows.push(values.slice()); this.log.push({ sheet: this.name, append: values }); }
}

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) { Object.assign(this, { sheet, row, col, numRows, numCols }); }
  getSheet() { return this.sheet; }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const line = [];
      for (let c = 0; c < this.numCols; c++) line.push(this.sheet.cell(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  setValue(v) { this.sheet.set(this.row, this.col, v); return this; }
  setFormula(f) { this.sheet.set(this.row, this.col, f); return this; }
  setDataValidation(rule) { this.sheet.log.push({ sheet: this.sheet.name, validation: rule, col: this.col }); return this; }
}

/** A fake YouTrack that answers the four calls the script makes and records each request. */
function fakeYouTrack(base, { states = {} } = {}) {
  const requests = [];
  let next = 100;
  const respond = (code, body) => ({ getResponseCode: () => code, getContentText: () => JSON.stringify(body) });
  return {
    requests,
    fetch(url, options = {}) {
      const method = (options.method || 'get').toLowerCase();
      const payload = options.payload ? JSON.parse(options.payload) : undefined;
      requests.push({ method, url, payload, headers: options.headers });
      const p = url.replace(base, '').split('?')[0];
      if (method === 'get' && /^\/api\/admin\/projects\/[^/]+\/customFields$/.test(p)) {
        return respond(200, [
          { field: { name: 'Priority' }, isRequired: true, $type: 'EnumProjectCustomField',
            bundle: { values: [{ name: 'Critical' }, { name: 'Major' }, { name: 'Normal' }, { name: 'Minor' }] } },
          { field: { name: 'Platforms' }, isRequired: false, $type: 'EnumProjectCustomField',
            bundle: { values: [{ name: 'Console' }, { name: 'PC' }] } },
          { field: { name: 'State' }, isRequired: true, $type: 'StateProjectCustomField',
            bundle: { values: [{ name: 'Open' }, { name: 'Fixed' }, { name: 'Verified' }] } },
        ]);
      }
      if (method === 'post' && p === '/api/issues') {
        next += 1;
        return respond(200, { id: '0-' + next, idReadable: 'DEMO-' + next });
      }
      if (method === 'post' && p === '/api/commands') return respond(200, {});
      const m = p.match(/^\/api\/issues\/([^/]+)$/);
      if (method === 'get' && m) {
        return respond(200, { customFields: [{ name: 'State', value: { name: states[m[1]] || 'Open' } }] });
      }
      return respond(404, { error: 'not faked: ' + method + ' ' + p });
    },
  };
}

/**
 * Builds a fresh context with the sample spreadsheet loaded and the src/ files evaluated.
 * @param {object} opts
 * @param {boolean} [opts.network=true] false: any UrlFetchApp call throws.
 * @param {number} [opts.now] fixed time (ms) for `new Date()` with no arguments.
 * @param {object} [opts.states] YouTrack State by permanent ID, for status sync.
 */
function load(opts = {}) {
  const writes = [];
  const logs = [];
  const sheets = {};
  for (const f of fs.readdirSync(path.join(ROOT, 'sample-data'))) {
    if (f.endsWith('.tsv')) {
      const name = f.slice(0, -4);
      sheets[name] = new FakeSheet(name, readTsv(path.join(ROOT, 'sample-data', f)), writes);
    }
  }
  sheets._youtrack_schema_cache = new FakeSheet('_youtrack_schema_cache', [], writes);
  const props = new Map([['YOUTRACK_TOKEN', 'test-token']]);
  const alerts = [];
  const fetchAttempts = [];
  const base = 'https://youtrack.example.com';
  const yt = fakeYouTrack(base, opts);

  const ctx = {
    console: { log: m => logs.push(String(m)), warn: m => logs.push(String(m)), error: m => logs.push(String(m)) },
    Logger: { log: m => logs.push(String(m)) },
    SpreadsheetApp: {
      openById: () => ({ getSheetByName: n => sheets[n] || null }),
      flush: () => {},
      getUi: () => ({ alert: m => alerts.push(m) }),
      newDataValidation: () => {
        const rule = {};
        const b = {
          requireValueInRange: (r) => { rule.range = r; return b; },
          setAllowInvalid: (v) => { rule.allowInvalid = v; return b; },
          setHelpText: (t) => { rule.help = t; return b; },
          build: () => rule,
        };
        return b;
      },
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: k => (props.has(k) ? props.get(k) : null),
        setProperty: (k, v) => props.set(k, v),
        deleteProperty: k => props.delete(k),
      }),
    },
    UrlFetchApp: {
      fetch: (url, o) => {
        // Recorded before the throw: the script catches some errors, so a throw alone could be swallowed.
        fetchAttempts.push(url);
        if (opts.network === false) throw new Error('network call during a read-only run: ' + url);
        return yt.fetch(url, o);
      },
    },
    Utilities: {
      DigestAlgorithm: { MD5: 'md5' },
      computeDigest: (alg, s) => [...crypto.createHash(alg).update(s, 'utf8').digest()].map(b => (b > 127 ? b - 256 : b)),
      sleep: () => {},
    },
  };
  vm.createContext(ctx);
  if (opts.now !== undefined) {
    vm.runInContext(`(() => { const R = Date; const T = ${Number(opts.now)};
      globalThis.Date = class extends R { constructor(...a) { if (a.length === 0) super(T); else super(...a); }
        static now() { return T; } }; })();`, ctx);
  }
  for (const f of SRC_FILES) {
    vm.runInContext(fs.readFileSync(path.join(ROOT, 'src', f), 'utf8'), ctx, { filename: 'src/' + f });
  }
  const config = sheets._generator_config;
  const headers = config.rows[0];
  const rowOf = id => config.rows.findIndex(r => r[0] === id) + 1; // 1-based sheet row
  const cellOf = (id, header) => config.cell(rowOf(id), headers.indexOf(header) + 1);
  return { ctx, sheets, writes, logs, alerts, yt, props, fetchAttempts, cellOf, text: () => logs.join('\n') };
}

module.exports = { load };
