// hygiene.test.js - no credential or real deployment ID in any file of the repository.
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SKIP = new Set(['.git', 'node_modules']);
const SELF = path.relative(ROOT, __filename);

function files(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap(e => {
    if (SKIP.has(e.name)) return [];
    const p = path.join(dir, e.name);
    return e.isDirectory() ? files(p) : [p];
  });
}

const PATTERNS = [
  ['YouTrack permanent token', /perm[:-][A-Za-z0-9+/=._-]{12,}/],
  ['Bearer token literal', /Bearer [A-Za-z0-9+/=._-]{16,}/],
  ['Google file or script ID (long base64url run)', /\b[A-Za-z0-9_-]{40,}\b/],
  ['GitHub token', /\bgh[pousr]_[A-Za-z0-9]{20,}/],
  ['private key', /-----BEGIN [A-Z ]*PRIVATE KEY-----/],
];

test('no file holds a token, key or real Google ID', () => {
  const hits = [];
  for (const f of files(ROOT)) {
    const rel = path.relative(ROOT, f);
    if (rel === SELF) continue;
    const text = fs.readFileSync(f, 'utf8');
    for (const [name, re] of PATTERNS) if (re.test(text)) hits.push(`${rel}: ${name}`);
  }
  assert.deepEqual(hits, []);
});

test('.clasp.json (holds the real script ID) is ignored, and only the example is shipped', () => {
  assert.match(fs.readFileSync(path.join(ROOT, '.gitignore'), 'utf8'), /^\.clasp\.json$/m);
  assert.equal(fs.existsSync(path.join(ROOT, '.clasp.json')), false);
  const example = JSON.parse(fs.readFileSync(path.join(ROOT, '.clasp.json.example'), 'utf8'));
  assert.match(example.scriptId, /^REPLACE_/);
});
