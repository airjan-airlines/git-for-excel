'use strict';
// The payloads Base44 stores and renders. The local source is exercised against a real repository of real
// workbooks; the GitHub source is exercised against a stubbed API so its URLs and parsing are pinned down
// without touching the network.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const JSZip = require('jszip');
const { Repo } = require('../src/main/repo');
const R = require('../src/sync/reader');

const DIR = process.env.OG_XLSX_DIR || path.join(os.homedir(), 'Downloads');
const PLAIN = /<c r="([A-Z]+\d+)"(?![^>]*\bt=)[^>]*><v>(-?[\d.]+)<\/v><\/c>/g;
const sheets = (z) => Object.keys(z.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'og-reader-'));
const NAME = 'Budget.xlsx';
const opts = { skip: fs.existsSync(DIR) && fs.readdirSync(DIR).some((n) => /\.xlsx$/i.test(n)) ? false : `no .xlsx files in ${DIR}` };

async function cellsOf(buf, want = 2) {
  const zip = await JSZip.loadAsync(buf);
  const out = [];
  for (const s of sheets(zip)) {
    const xml = await zip.file(s).async('string');
    for (const m of xml.matchAll(PLAIN)) { out.push(m[1]); if (out.length >= want) return out; }
  }
  return out;
}
async function setCell(file, addr, value) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const s of sheets(zip)) {
    const xml = await zip.file(s).async('string');
    const re = new RegExp(`<c r="${addr}"(?![^>]*\\bt=)[^>]*><v>(-?[\\d.]+)</v></c>`);
    const m = re.exec(xml);
    if (!m) continue;
    zip.file(s, xml.replace(m[0], m[0].replace(`<v>${m[1]}</v>`, `<v>${value}</v>`)));
    fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    return;
  }
  throw new Error('cell not found ' + addr);
}

let SRC = null;
test.before(async () => {
  for (const n of fs.readdirSync(DIR).filter((x) => /\.xlsx$/i.test(x) && !x.startsWith('~$'))) {
    const f = path.join(DIR, n);
    if ((await cellsOf(fs.readFileSync(f), 2)).length >= 2) { SRC = f; return; }
  }
});

// Build a small repo: initial commit, an edit on main, a branch with another edit.
async function fixture() {
  const root = tmp();
  fs.copyFileSync(SRC, path.join(root, NAME));
  const repo = await new Repo(root).open();
  const [cellA, cellB] = await cellsOf(fs.readFileSync(SRC), 2);
  const c1 = await repo.commit({ message: 'Add the budget', paths: [NAME], author: 'Ana Lyst' });
  await setCell(path.join(root, NAME), cellA, 4242);
  const c2 = await repo.commit({ message: 'Update registration income', description: 'Board asked for a revised figure.', paths: [NAME], author: 'Ana Lyst' });
  await repo.createAndSwitch('travel-revision');
  await setCell(path.join(root, NAME), cellB, 9999);
  const c3 = await repo.commit({ message: 'Raise the travel line', paths: [NAME], author: 'Ben Ops' });
  await repo.switchBranch('main');
  return { root, repo, cellA, cellB, c1, c2, c3 };
}

test('local source: branch and commit payloads match what the web app needs', opts, async () => {
  assert.ok(SRC, 'a workbook with two editable cells');
  const f = await fixture();
  const src = R.localSource(f.root);

  const branches = await R.listBranches(src, 'repo-1');
  assert.deepEqual(branches.map((b) => b.name).sort(), ['main', 'travel-revision']);
  const main = branches.find((b) => b.name === 'main');
  assert.equal(main.isDefault, true);
  assert.match(main.head, /^[0-9a-f]{40}$/);
  assert.ok(!Number.isNaN(Date.parse(main.updatedAt)));

  const commits = await R.listCommits(src, 'repo-1', 'main', 10);
  assert.deepEqual(commits.map((c) => c.message), ['Update registration income', 'Add the budget']);
  const top = commits[0];
  assert.equal(top.repositoryId, 'repo-1');
  assert.equal(top.author, 'Ana Lyst');
  assert.equal(top.description, 'Board asked for a revised figure.');
  assert.equal(top.isMerge, false);
  assert.deepEqual(top.parentShas, [commits[1].sha]);
  assert.deepEqual(top.files, [{ path: NAME, status: 'modified' }]);
  assert.ok(!Number.isNaN(Date.parse(top.createdAt)));
  // the payload must survive a JSON round trip unchanged - it is stored as an entity
  assert.deepEqual(JSON.parse(JSON.stringify(top)), top);

  // the first commit reports the file as added
  assert.deepEqual(commits[1].files, [{ path: NAME, status: 'added' }]);
  // branch commits are visible too
  const branchCommits = await R.listCommits(src, 'repo-1', 'travel-revision', 10);
  assert.equal(branchCommits[0].message, 'Raise the travel line');
  assert.equal(branchCommits[0].author, 'Ben Ops');
});

test('local source: getDiff returns the cell-level change a reviewer reads', opts, async () => {
  const f = await fixture();
  const src = R.localSource(f.root);
  const commits = await R.listCommits(src, 'repo-1', 'main', 10);

  const d = await R.getDiff(src, { sha: commits[0].sha, path: NAME });
  assert.equal(d.format, 'xlsx');
  assert.equal(d.status, 'modified');
  assert.equal(d.changeCount, 1);
  const cell = d.changes[0];
  assert.equal(cell.kind, 'cell');
  assert.equal(cell.detail, 'input');
  assert.ok(cell.label.endsWith('!' + f.cellA));
  assert.equal(String(cell.after.value), '4242');
  assert.match(d.summary, /1 cell modified/);
  assert.deepEqual(JSON.parse(JSON.stringify(d)), d, 'diff payload is plain JSON');

  // the very first commit has no parent: everything reads as added
  const first = await R.getDiff(src, { sha: commits[1].sha, path: NAME });
  assert.equal(first.parentSha, null);
  assert.ok(first.changes.every((c) => c.type === 'added'));
  assert.ok(first.changeCount > 1);

  await assert.rejects(R.getDiff(src, { sha: commits[0].sha, path: 'Nope.xlsx' }), /not in commit/);
});

// --- GitHub source against a stubbed API (no network) ---
function stubGitHub(f) {
  const buf = fs.readFileSync(SRC);
  const calls = [];
  const routes = {
    '/repos/acme/budgets': { default_branch: 'main' },
    '/repos/acme/budgets/branches?per_page=100': [
      { name: 'main', commit: { sha: 'a'.repeat(40) } },
      { name: 'travel', commit: { sha: 'b'.repeat(40) } },
    ],
    '/repos/acme/budgets/commits?sha=main&per_page=50': [{
      sha: 'a'.repeat(40),
      parents: [{ sha: 'c'.repeat(40) }],
      commit: { author: { name: 'Ana Lyst', email: 'ana@acme.test', date: '2026-10-10T12:00:00Z' }, message: 'Update registration income\n\nBoard asked.' },
    }],
    [`/repos/acme/budgets/commits/${'a'.repeat(40)}`]: {
      files: [{ filename: NAME, status: 'modified' }, { filename: 'notes.txt', status: 'modified' }],
    },
    [`/repos/acme/budgets/contents/${NAME}?ref=${'a'.repeat(40)}`]: { sha: 'd'.repeat(40) },
    [`/repos/acme/budgets/git/blobs/${'d'.repeat(40)}`]: { encoding: 'base64', content: buf.toString('base64') },
  };
  const fetch = async (url, init) => {
    const u = url.replace('https://api.github.com', '');
    calls.push({ url: u, headers: init.headers });
    if (!(u in routes)) return { ok: false, status: 404, text: async () => `no stub for ${u}` };
    return { ok: true, status: 200, json: async () => routes[u], text: async () => JSON.stringify(routes[u]) };
  };
  return { source: R.githubSource({ owner: 'acme', repo: 'budgets', token: 'tok_123', fetch }), calls, routes };
}

test('github source: authenticates, lists branches, and reads a workbook through the blob API', opts, async () => {
  assert.ok(SRC);
  const { source, calls } = stubGitHub();

  const branches = await R.listBranches(source, 'repo-1');
  assert.deepEqual(branches.map((b) => [b.name, b.isDefault]), [['main', true], ['travel', false]]);
  assert.equal(calls[0].headers.authorization, 'Bearer tok_123', 'token is sent');
  assert.equal(calls[0].headers['x-github-api-version'], '2022-11-28');

  const commits = await R.listCommits(source, 'repo-1', 'main', 50);
  assert.equal(commits.length, 1);
  assert.equal(commits[0].message, 'Update registration income');
  assert.equal(commits[0].description, 'Board asked.');
  assert.deepEqual(commits[0].parentShas, ['c'.repeat(40)]);
  assert.deepEqual(commits[0].files, [{ path: NAME, status: 'modified' }], 'non-Office files are ignored');

  // reading a workbook goes contents -> blob (the contents API alone caps at 1 MB)
  const bytes = await source.readFile('a'.repeat(40), NAME);
  assert.deepEqual(bytes, fs.readFileSync(SRC), 'bytes round-trip through base64');
  assert.ok(calls.some((c) => c.url.includes('/git/blobs/')), 'used the blob API');
});

test('github source: surfaces API errors with their status', opts, async () => {
  const { source } = stubGitHub();
  await assert.rejects(source.commits('missing-branch', 50), /GitHub 404/);
});
