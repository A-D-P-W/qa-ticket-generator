// generator.test.js - runs the real src/ code against sample-data/ and a fake YouTrack.
'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const { load } = require('./fake-gas');

const issuesCreated = yt => yt.requests.filter(r => r.method === 'post' && r.url.includes('/api/issues?'));
const commands = yt => yt.requests.filter(r => r.url.endsWith('/api/commands'));

test('diagnostics is read-only and reports exactly the two broken jobs', () => {
  const g = load({ network: false }); // any UrlFetchApp call would throw
  g.ctx.QAGEN_runDiagnostics();
  const out = g.text();
  assert.match(out, /Found 10 active job\(s\)/);
  assert.match(out, /DIAGNOSTICS COMPLETE: Found 2 issue\(s\)/);
  assert.match(out, /JobID 'LVL-002'\): Missing required parameter\(s\)\. The blueprint requires: \[Level_Name\]/);
  assert.match(out, /JobID 'BOSS-001'\): A matching rule was not found/);
  assert.equal(g.writes.length, 0, 'no cell was written');
  assert.deepEqual(g.fetchAttempts, [], 'no HTTP request was attempted');
  assert.equal(g.props.has('YOUTRACK_SCHEMA_HASH'), false, 'no script property was written');
});

test('diagnostics previews a TEST job without creating it', () => {
  const g = load({ network: false });
  g.ctx.QAGEN_runDiagnostics();
  assert.match(g.text(), /MOCK PREVIEW for Job 'SMK-002'/);
  assert.match(g.text(), /Summary: \[Playthrough\] Level 1 \(Tutorial\)/);
});

test('parameters: global < job column < Parameters cell; gaps stay visible', () => {
  const g = load({ network: false });
  const map = g.ctx._buildParameterMap({ Platform: 'Console', Parameters: 'Platform:Mobile; Note:a:b' });
  assert.equal(map.get('Build_Version'), '0.1.0'); // from _Global_Parameters
  assert.equal(map.get('Platform'), 'Mobile'); // Parameters cell wins over the column and the global
  assert.equal(map.get('Note'), 'a:b'); // split on the first colon only
  assert.equal(g.ctx._applyParameters('x ${Missing} y', {}), 'x [UNDEFINED_PARAMETER:Missing] y');
});

test('live run creates, links and writes back only the jobs that pass every gate', () => {
  const g = load();
  g.ctx.QAGEN_runTestGeneration();
  const created = issuesCreated(g.yt).map(r => r.payload.summary);
  assert.deepEqual(created, [
    '[Smoke] Build 0.1.0 on PC', // SMK-001
    '[QA Checklist] - UI/UX: Functionality', // OPT-001 (3 items: one combined ticket)
    '[Text] UI strings (French)', // TXT-001 parent (30 items)
    '[Text] UI strings (French) (Part 1 of 2)',
    '[Text] UI strings (French) (Part 2 of 2)',
    '[Smoke] Milestone build 0.1.0', // SCH-001..003: a manual run ignores triggers
    '[Smoke] Weekly build 0.1.0',
    '[Smoke] Every 7 days 0.1.0',
  ]);
  // Gates that stopped a job: missing parameter, unmet dependency, no blueprint.
  for (const id of ['LVL-002', 'DEP-001', 'BOSS-001']) assert.equal(g.cellOf(id, 'YouTrack_Ticket_ID'), '', id);
  assert.match(g.text(), /Missing required parameter: \[Level_Name\]\. Halting generation for Job_ID: \[LVL-002\]/);
  assert.match(g.text(), /Dependency 'SMK-001' has status ""/);

  // Sub-tasks are linked from the child side with the exact command keyword.
  assert.deepEqual(commands(g.yt).map(r => [r.payload.query, r.payload.issues[0].idReadable]), [
    ['add subtask of DEMO-103', 'DEMO-104'],
    ['add subtask of DEMO-103', 'DEMO-105'],
  ]);

  // Checklist chunking: 25 items, then 5.
  const parts = issuesCreated(g.yt).filter(r => /Part \d of 2/.test(r.payload.summary));
  assert.deepEqual(parts.map(r => (r.payload.description.match(/^- \[ \] /gm) || []).length), [25, 5]);

  // Write-back: both IDs, a link and a status, on the job's own row.
  assert.equal(g.cellOf('SMK-001', 'YouTrack_Ticket_ID'), 'DEMO-101');
  assert.equal(g.cellOf('SMK-001', 'YouTrack_Permanent_ID'), '0-101');
  assert.equal(g.cellOf('SMK-001', 'YouTrack_Ticket_URL'), '=HYPERLINK("https://youtrack.example.com/issue/DEMO-101","View Ticket")');
  assert.equal(g.cellOf('TXT-001', 'YouTrack_Ticket_URL'), '=HYPERLINK("https://youtrack.example.com/issue/DEMO-103","View Suite")');
  assert.equal(g.cellOf('SMK-001', 'YouTrack_Status'), 'Open');
});

test('payload: schema-driven default field, Markdown, real newlines, token from Script Properties', () => {
  const g = load();
  g.ctx.QAGEN_runTestGeneration();
  const first = issuesCreated(g.yt)[0];
  assert.equal(first.payload.project.id, 'REPLACE_WITH_PROJECT_ID');
  assert.equal(first.payload.usesMarkdown, true);
  assert.deepEqual(first.payload.customFields.find(f => f.name === 'Platforms'),
    { name: 'Platforms', $type: 'MultiEnumIssueCustomField', value: [{ name: 'PC' }] });
  assert.ok(first.payload.description.includes('\n2. Launch it.'), 'literal \\n from the sheet became a newline');
  assert.ok(!first.payload.description.includes('\\n'));
  assert.equal(first.headers.Authorization, 'Bearer test-token');
  // The schema cache sheet was filled once and its hash stored.
  assert.equal(g.sheets._youtrack_schema_cache.rows.length, 4);
  assert.match(g.props.get('YOUTRACK_SCHEMA_HASH'), /^[0-9a-f]{32}$/);
});

test('a second run creates no duplicates; a dependency opens once its status is Verified', () => {
  const g = load();
  g.ctx.QAGEN_runTestGeneration();
  const before = issuesCreated(g.yt).length;
  g.ctx.QAGEN_runTestGeneration();
  assert.equal(issuesCreated(g.yt).length, before, 'rows with a permanent ID are skipped');

  const row = g.sheets._generator_config.rows.findIndex(r => r[0] === 'SMK-001') + 1;
  g.sheets._generator_config.set(row, g.sheets._generator_config.rows[0].indexOf('YouTrack_Status') + 1, 'Verified');
  g.ctx.QAGEN_runTestGeneration();
  const after = issuesCreated(g.yt).slice(before).map(r => r.payload.summary);
  assert.deepEqual(after, ['[Save/Load] Inventory']); // DEP-001
});

test('campaign run takes only the jobs tagged with that campaign', () => {
  const g = load();
  g.ctx.QAGEN_executeCampaign('Localization');
  assert.deepEqual(issuesCreated(g.yt).map(r => r.payload.summary), [
    '[Text] UI strings (French)',
    '[Text] UI strings (French) (Part 1 of 2)',
    '[Text] UI strings (French) (Part 2 of 2)',
  ]);
});

test('scheduled triggers on Monday 2026-03-02 (and the recurring limitation)', () => {
  const monday = new Date(2026, 2, 2, 12, 0, 0).getTime();
  const g = load({ now: monday });
  g.ctx.QAGEN_evaluateTriggers();
  assert.deepEqual(issuesCreated(g.yt).map(r => r.payload.summary), [
    '[Smoke] Milestone build 0.1.0', // On Date 2026-03-01, never run
    '[Smoke] Weekly build 0.1.0', // Weekly, Monday
  ]);
  // Known limitation, kept as the original behaves: a Recurring job with no Last_Run never fires.
  assert.equal(g.cellOf('SCH-003', 'YouTrack_Ticket_ID'), '');
  assert.notEqual(g.cellOf('SCH-002', 'Last_Run'), '', 'autonomous runs stamp Last_Run');
});

test('status sync copies the YouTrack State into the sheet', () => {
  const g = load({ states: { '0-1': 'Fixed' } });
  g.ctx.QAGEN_syncYouTrackStatuses();
  assert.equal(g.cellOf('DONE-001', 'YouTrack_Status'), 'Fixed');
  assert.match(g.text(), /Updated 1 ticket status\(es\)/);
});

test('data validation dropdowns come from _data_ref', () => {
  const g = load({ network: false });
  g.ctx.QAGEN_applyDataValidationRules();
  const applied = g.writes.filter(w => w.validation).map(w => g.sheets._generator_config.rows[0][w.col - 1]);
  assert.deepEqual(applied, ['Category', 'Sub_Category', 'Priority', 'Impact_Level']);
  assert.match(g.alerts[0], /applied successfully/);
});
