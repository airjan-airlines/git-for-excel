'use strict';
// Mapping from the engine's diff to GitXL's `changes[]` rows, and the push flow against a stubbed client.
const test = require('node:test');
const assert = require('node:assert/strict');
const gitxl = require('../src/sync/gitxl');

test('a changed formula maps to property=formula with = prefixes', () => {
  const row = gitxl.mapChange('Q4_Model.xlsx', {
    id: 'cell:P&L!D18', kind: 'cell', type: 'modified', detail: 'formula', label: 'P&L!D18',
    before: { value: '2000', formula: 'D6*0.5', type: 'n', style: 's1' },
    after: { value: '1400', formula: 'D6*0.35', type: 'n', style: 's1' },
  });
  assert.deepEqual(row, {
    file_path: 'Q4_Model.xlsx', file_type: 'xlsx',
    unit_id: 'cell:P&L!D18', unit_label: 'P&L!D18', unit_type: 'cell',
    change_type: 'modified', before: '=D6*0.5', after: '=D6*0.35', property: 'formula',
  });
});

test('a typed value maps to property=value; a recalculated one does too', () => {
  const mk = (detail) => gitxl.mapChange('a.xlsx', {
    id: 'cell:Sheet1!B3', kind: 'cell', type: 'modified', detail, label: 'Sheet1!B3',
    before: { value: '5', formula: null }, after: { value: '4242', formula: null },
  });
  assert.equal(mk('input').property, 'value');
  assert.equal(mk('calculated').property, 'value');
  assert.deepEqual([mk('input').before, mk('input').after], ['5', '4242']);
});

test('added and removed cells carry a null on the missing side', () => {
  const added = gitxl.mapChange('a.xlsx', { id: 'cell:S!A1', kind: 'cell', type: 'added', label: 'S!A1', before: null, after: { value: 'Notes', formula: null } });
  assert.deepEqual([added.change_type, added.before, added.after], ['added', null, 'Notes']);
  const removed = gitxl.mapChange('a.xlsx', { id: 'cell:S!A1', kind: 'cell', type: 'removed', label: 'S!A1', before: { value: 'Old', formula: null }, after: null });
  assert.deepEqual([removed.change_type, removed.before, removed.after], ['removed', 'Old', null]);
});

test('non-cell units keep their kind and still produce a readable before/after', () => {
  const para = gitxl.mapChange('Policy.docx', {
    id: 'docx:body:modified:2:2', kind: 'paragraph', type: 'modified', detail: 'text', label: 'Paragraph 3',
    before: { text: 'book 14 days ahead' }, after: { text: 'book 21 days ahead' },
  });
  assert.deepEqual([para.unit_type, para.file_type, para.property, para.before, para.after],
    ['paragraph', 'docx', 'text', 'book 14 days ahead', 'book 21 days ahead']);
  // "moved" is not one of GitXL's change types, so it is reported as a modification
  assert.equal(gitxl.mapChange('a.pptx', { id: 'x', kind: 'slide-order', type: 'moved', label: 'Slide 2', before: {}, after: {} }).change_type, 'modified');
});

test('stats count additions and deletions; an edit counts as one of each', () => {
  assert.deepEqual(gitxl.countStats([{ change_type: 'added' }, { change_type: 'added' }, { change_type: 'removed' }, { change_type: 'modified' }]),
    { additions: 3, deletions: 2 });
});

test('merge commits name the branch they brought in', () => {
  assert.equal(gitxl.mergeSource("Merge branch 'travel-budget' into main"), 'travel-budget');
  assert.equal(gitxl.mergeSource("Merge branch 'main' of C:/x/base44.git"), 'main');
  assert.equal(gitxl.mergeSource("Merge remote-tracking branch 'base44/main'"), 'main');
  assert.equal(gitxl.mergeSource('Update budget'), null);
});

// --- push flow against a stubbed Base44 client ---
function stubClient(seed = {}) {
  const db = { Repository: [], Branch: [], Commit: [], ...seed };
  let n = 0;
  const match = (row, q) => Object.entries(q || {}).every(([k, v]) => row[k] === v);
  return {
    db,
    async list(entity, { q } = {}) { return (db[entity] || []).filter((r) => match(r, q)); },
    async create(entity, data) { const rec = { ...data, id: `rec_${++n}` }; (db[entity] ||= []).push(rec); return rec; },
    async update(entity, id, data) {
      const row = (db[entity] || []).find((r) => r.id === id);
      Object.assign(row, data);
      return row;
    },
  };
}

// A fake Repo exposing only what push() uses.
function fakeRepo(commits, branches = ['main']) {
  return {
    root: 'C:/work/Budgets',
    // `git rev-parse <branch>` resolves to the branch tip, i.e. its newest commit
    g: { try: async (_c, ref) => ({ ok: true, out: (commits.filter((c) => c.branch === ref).pop() || commits[commits.length - 1]).id }) },
    async branches() { return { list: branches.map((name) => ({ name, current: name === 'main' })) }; },
    async log(branch) { return commits.filter((c) => c.branch === branch).slice().reverse(); },
    async diffCommit(id, p) {
      const c = commits.find((x) => x.id === id);
      return { format: 'xlsx', changes: c.diff || [] };
    },
  };
}

const COMMITS = [
  { id: 'a'.repeat(40), parents: [], author: 'Ana Lyst', email: 'ana@acme.test', message: 'Add the budget', branch: 'main', changed: [{ path: 'Budget.xlsx', status: 'added' }], diff: [{ id: 'cell:S!A1', kind: 'cell', type: 'added', label: 'S!A1', before: null, after: { value: 'Revenue', formula: null } }] },
  { id: 'b'.repeat(40), parents: ['a'.repeat(40)], author: 'Ana Lyst', email: 'ana@acme.test', message: 'Update income', branch: 'main', changed: [{ path: 'Budget.xlsx', status: 'modified' }], diff: [{ id: 'cell:S!B3', kind: 'cell', type: 'modified', detail: 'input', label: 'S!B3', before: { value: '5' }, after: { value: '4242' } }] },
];

test('push creates the repository, its commits oldest-first, and its branches', async () => {
  const client = stubClient();
  const steps = [];
  const out = await gitxl.push(fakeRepo(COMMITS), client, { owner: 'Ana Lyst', onProgress: (p) => steps.push(p.step) });

  assert.equal(client.db.Repository.length, 1);
  const r = client.db.Repository[0];
  assert.equal(r.name, 'Budgets', 'named after the folder');
  assert.equal(r.default_branch, 'main');
  assert.equal(r.file_type, 'xlsx');

  assert.equal(client.db.Commit.length, 2);
  assert.deepEqual(client.db.Commit.map((c) => c.message), ['Add the budget', 'Update income'], 'parents first');
  const second = client.db.Commit[1];
  assert.equal(second.hash, 'b'.repeat(7), 'seven-character hash, as GitXL stores');
  assert.equal(second.parent_hash, 'a'.repeat(7));
  assert.equal(second.author_email, 'ana@acme.test');
  assert.equal(second.repository_id, r.id);
  assert.deepEqual(second.files, [{ path: 'Budget.xlsx', change_type: 'modified', summary: '1 cell modified' }]);
  assert.deepEqual(second.changes[0].unit_id, 'cell:S!B3');
  assert.deepEqual(second.stats, { additions: 1, deletions: 1 });
  assert.deepEqual(client.db.Commit[0].stats, { additions: 1, deletions: 0 });

  assert.equal(client.db.Branch.length, 1);
  assert.equal(client.db.Branch[0].head_commit_id, second.id, 'branch points at the tip commit record');
  assert.equal(client.db.Branch[0].protected, true, 'the default branch is protected');
  assert.deepEqual(out, { repositoryId: r.id, repositoryName: 'Budgets', branches: 1, commitsPushed: 2, commitsTotal: 2 });
  assert.deepEqual(steps, ['repository', 'commit', 'commit', 'branch']);
});

test('pushing twice adds nothing and does not duplicate the repository', async () => {
  const client = stubClient();
  await gitxl.push(fakeRepo(COMMITS), client, {});
  const again = await gitxl.push(fakeRepo(COMMITS), client, {});
  assert.equal(again.commitsPushed, 0);
  assert.equal(client.db.Repository.length, 1);
  assert.equal(client.db.Commit.length, 2);
  assert.equal(client.db.Branch.length, 1, 'the branch is updated, not re-created');
});

test('a file that cannot be compared still produces a commit, with the reason in its summary', async () => {
  const client = stubClient();
  const broken = [{ ...COMMITS[0], diff: null }];
  const repo = fakeRepo(broken);
  repo.diffCommit = async () => { throw new Error('zip is corrupt'); };
  await gitxl.push(repo, client, {});
  assert.match(client.db.Commit[0].files[0].summary, /could not be compared \(zip is corrupt\)/);
  assert.deepEqual(client.db.Commit[0].changes, []);
});
