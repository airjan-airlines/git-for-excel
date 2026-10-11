'use strict';
// Base44 as the git host: two people share a workbook through an in-memory stand-in for the Base44 app
// (entities + file storage). Git, the bundles and the Office merge driver are all real.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const JSZip = require('jszip');
const { Repo } = require('../src/main/repo');
const hosted = require('../src/sync/hosted');

const DIR = process.env.OG_XLSX_DIR || path.join(os.homedir(), 'Downloads');
const PLAIN = /<c r="([A-Z]+\d+)"(?![^>]*\bt=)[^>]*><v>(-?[\d.]+)<\/v><\/c>/g;
const sheets = (z) => Object.keys(z.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();
async function cellsOf(buf, want = 2) {
  const zip = await JSZip.loadAsync(buf);
  const out = [];
  for (const s of sheets(zip)) {
    for (const m of (await zip.file(s).async('string')).matchAll(PLAIN)) { out.push(m[1]); if (out.length >= want) return out; }
  }
  return out;
}
async function setCell(file, addr, value) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const s of sheets(zip)) {
    const xml = await zip.file(s).async('string');
    const m = new RegExp(`<c r="${addr}"(?![^>]*\\bt=)[^>]*><v>(-?[\\d.]+)</v></c>`).exec(xml);
    if (!m) continue;
    zip.file(s, xml.replace(m[0], m[0].replace(`<v>${m[1]}</v>`, `<v>${value}</v>`)));
    fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    return;
  }
  throw new Error('cell not found: ' + addr);
}
async function cellValue(repo, rev, name, addr) {
  const t = await repo.treeAt(rev, name);
  const hit = Object.entries(t.units).find(([k, u]) => u.kind === 'cell' && k.endsWith('!' + addr));
  return hit && hit[1].data.value;
}

// The slice of the Base44 client hosted.js uses: entities and file storage, in memory.
function fakeBase44() {
  const db = { Repository: [], Branch: [], Commit: [] };
  const files = new Map();
  let n = 0;
  const match = (row, q) => Object.entries(q || {}).every(([k, v]) => row[k] === v);
  return {
    db, files,
    async list(entity, { q } = {}) { return (db[entity] || []).filter((r) => match(r, q)); },
    async get(entity, id) { const r = (db[entity] || []).find((x) => x.id === id); if (!r) throw new Error('404'); return { ...r }; },
    async create(entity, data) { const rec = { ...data, id: `rec_${++n}` }; (db[entity] ||= []).push(rec); return rec; },
    async update(entity, id, data) { const r = db[entity].find((x) => x.id === id); Object.assign(r, data); return r; },
    async upload(buf, name) { const url = `https://files.test/${++n}_${name}`; files.set(url, Buffer.from(buf)); return url; },
    async download(url) { return files.get(url); },
  };
}

let SRC = null;
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'og-hosted-'));
const NAME = 'Budget.xlsx';
const opts = { skip: fs.existsSync(DIR) && fs.readdirSync(DIR).some((n) => /\.xlsx$/i.test(n)) ? false : `no .xlsx files in ${DIR}` };

test.before(async () => {
  if (!fs.existsSync(DIR)) return;
  for (const n of fs.readdirSync(DIR).filter((x) => /\.xlsx$/i.test(x) && !x.startsWith('~$'))) {
    const f = path.join(DIR, n);
    if ((await cellsOf(fs.readFileSync(f), 2)).length >= 2) { SRC = f; return; }
  }
});

test('two people share a workbook through Base44: publish, clone, push, pull, merge', opts, async () => {
  assert.ok(SRC, 'found a workbook with two editable cells');
  const [cellA, cellB] = await cellsOf(fs.readFileSync(SRC), 2);
  const b44 = fakeBase44();

  // Ana creates the repository on Base44 from her folder and publishes it
  const anaDir = tmp();
  fs.copyFileSync(SRC, path.join(anaDir, NAME));
  const ana = await new Repo(anaDir).open();
  await ana.commit({ message: 'Add the budget', paths: [NAME], author: 'Ana' });
  const linkedA = await hosted.link(ana, b44, { name: 'Team Budget', owner: 'Ana' });
  assert.equal(linkedA.cloned, false, 'nothing hosted yet');
  const pub = await hosted.push(ana, b44, { owner: 'Ana' });
  assert.equal(pub.branch, 'main');
  const repoRec = b44.db.Repository[0];
  assert.equal(repoRec.bundle_url, pub.url, 'the bundle is what Base44 points at');
  assert.equal(b44.db.Commit.length, 1, 'the web app gets a commit record');
  assert.equal(b44.db.Branch[0].name, 'main');
  assert.equal((await ana.syncState()).upstream, 'base44/main', 'publishing sets the upstream');

  // Ben brings it down into an empty folder: the workbook appears with its history
  const benDir = tmp();
  const ben = await new Repo(benDir).open();
  const linkedB = await hosted.link(ben, b44, { repositoryId: repoRec.id });
  assert.equal(linkedB.cloned, true);
  assert.ok(fs.existsSync(path.join(benDir, NAME)), 'workbook checked out');
  assert.equal((await ben.log()).length, 1);

  // Both edit different cells; Ben pushes first
  await setCell(path.join(benDir, NAME), cellB, 777);
  await ben.commit({ message: 'Ben edits ' + cellB, paths: [NAME], author: 'Ben' });
  await hosted.push(ben, b44, { owner: 'Ben' });
  assert.equal(b44.db.Commit.length, 2);

  await setCell(path.join(anaDir, NAME), cellA, 4242);
  await ana.commit({ message: 'Ana edits ' + cellA, paths: [NAME], author: 'Ana' });

  // Ana is behind, so her push is refused - nothing hosted changes
  const urlBefore = b44.db.Repository[0].bundle_url;
  await assert.rejects(hosted.push(ana, b44, { owner: 'Ana' }), /Pull first/);
  assert.equal(b44.db.Repository[0].bundle_url, urlBefore);

  // She pulls: different cells, so the merge driver combines them with no conflict
  await hosted.fetch(ana, b44);
  assert.equal((await ana.syncState()).behind, 1);
  const pulled = await ana.pull('base44');
  assert.deepEqual(pulled.conflicts, []);
  assert.equal(String(await cellValue(ana, 'HEAD', NAME, cellA)), '4242');
  assert.equal(String(await cellValue(ana, 'HEAD', NAME, cellB)), '777');

  // ...and pushes the merge; Ben pulls it and has both edits
  await hosted.push(ana, b44, { owner: 'Ana' });
  const merge = b44.db.Commit.find((c) => c.is_merge);
  assert.ok(merge, 'the merge is recorded as a merge for the web app');
  assert.ok(merge.second_parent_hash);
  await hosted.fetch(ben, b44);
  await ben.pull('base44');
  assert.equal(String(await cellValue(ben, 'HEAD', NAME, cellA)), '4242');
  assert.equal(String(await cellValue(ben, 'HEAD', NAME, cellB)), '777');
  assert.equal((await ben.syncState()).behind, 0);
});

test('a push that races another push is refused instead of overwriting it', opts, async () => {
  assert.ok(SRC);
  const b44 = fakeBase44();
  const dir = tmp();
  fs.copyFileSync(SRC, path.join(dir, NAME));
  const repo = await new Repo(dir).open();
  await repo.commit({ message: 'Add the budget', paths: [NAME], author: 'Ana' });
  await hosted.link(repo, b44, { name: 'Race' });

  // someone else's push lands between our fetch and our publish
  const realUpload = b44.upload;
  b44.upload = async (buf, name) => { b44.db.Repository[0].bundle_url = await realUpload(Buffer.from('theirs'), 'theirs'); return realUpload(buf, name); };
  b44.download = async () => fs.readFileSync(await bundleOf(repo));
  await assert.rejects(hosted.push(repo, b44, {}), /Someone else pushed/);
});

// a valid bundle of the repo, used as "their" hosted copy in the race test
async function bundleOf(repo) {
  const f = path.join(tmp(), 'x.bundle');
  await repo.g.git('bundle', 'create', f, '--branches');
  return f;
}

test("a teammate's shared what-if shows up in your scenario list and can be opened", opts, async () => {
  assert.ok(SRC);
  const b44 = fakeBase44();
  const anaDir = tmp();
  fs.copyFileSync(SRC, path.join(anaDir, NAME));
  const ana = await new Repo(anaDir).open();
  await ana.commit({ message: 'Add the budget', paths: [NAME], author: 'Ana' });
  await hosted.link(ana, b44, { name: 'Scenarios' });
  await hosted.push(ana, b44, {});

  const ben = await new Repo(tmp()).open();
  await hosted.link(ben, b44, { repositoryId: b44.db.Repository[0].id });

  // Ana starts a what-if, edits, and shares it (without touching main)
  await ana.createAndSwitch('Downside-case');
  const [cellA] = await cellsOf(fs.readFileSync(SRC), 1);
  await setCell(path.join(anaDir, NAME), cellA, 1);
  await ana.commit({ message: 'Stress revenue', paths: [NAME], author: 'Ana' });
  await hosted.push(ana, b44, {});

  await hosted.fetch(ben, b44);
  const { list } = await ben.scenarios();
  const theirs = list.find((s) => s.name === 'Downside-case');
  assert.ok(theirs, 'listed');
  assert.deepEqual([theirs.local, theirs.shared, theirs.author, theirs.subject], [false, true, 'Ana', 'Stress revenue']);

  const r = await ben.switchBranch('Downside-case');
  assert.equal(r.switched, true, 'opening it creates a local copy that tracks the shared one');
  assert.equal(String(await cellValue(ben, 'HEAD', NAME, cellA)), '1');
  assert.equal((await ben.syncState()).upstream, 'base44/Downside-case');
  assert.notEqual(String(await cellValue(ben, 'main', NAME, cellA)), '1', 'main is untouched');
});
