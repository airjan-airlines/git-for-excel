'use strict';
// Version control on real .xlsx files, backed by git. Originals are read-only; every test runs in a temp copy.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const JSZip = require('jszip');
const { Repo } = require('../src/main/repo');

const DIR = process.env.OG_XLSX_DIR || path.join(os.homedir(), 'Downloads');
const REAL = fs.existsSync(DIR)
  ? fs.readdirSync(DIR).filter((n) => /\.xlsx$/i.test(n) && !n.startsWith('~$')).map((n) => path.join(DIR, n)).sort((a, b) => fs.statSync(a).size - fs.statSync(b).size)
  : [];
const opts = { skip: REAL.length ? false : `no .xlsx files in ${DIR}` };
const sha = (b) => crypto.createHash('sha256').update(b).digest('hex');
const originals = new Map(REAL.map((f) => [f, sha(fs.readFileSync(f))]));
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'og-git-'));
const NAME = 'Model.xlsx';
let SRC = REAL[0]; // replaced in before() with a workbook that has editable numeric cells

// Edit a numeric cell straight in the workbook XML (stands in for a user editing in Excel).
async function setCell(buf, addr, value) {
  const zip = await JSZip.loadAsync(buf);
  for (const s of Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort()) {
    const xml = await zip.file(s).async('string');
    const re = new RegExp(`<c r="${addr}"(?![^>]*\\bt=)[^>]*><v>(-?[\\d.]+)</v></c>`);
    const m = re.exec(xml);
    if (!m) continue;
    zip.file(s, xml.replace(m[0], m[0].replace(`<v>${m[1]}</v>`, `<v>${value}</v>`)));
    return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  }
  throw new Error(`cell ${addr} not found`);
}
async function numericCells(buf, n = 2) {
  const zip = await JSZip.loadAsync(buf);
  const found = [];
  for (const s of Object.keys(zip.files).filter((x) => /^xl\/worksheets\/sheet\d+\.xml$/.test(x)).sort()) {
    const xml = await zip.file(s).async('string');
    for (const m of xml.matchAll(/<c r="([A-Z]+\d+)"(?![^>]*\bt=)[^>]*><v>(-?[\d.]+)<\/v><\/c>/g)) {
      found.push({ addr: m[1], value: m[2] });
      if (found.length >= n) return found;
    }
  }
  return found;
}
async function cellValue(repo, rev, addr) {
  const t = await repo.treeAt(rev, NAME);
  const hit = Object.entries(t.units).find(([k, u]) => u.kind === 'cell' && k.endsWith('!' + addr));
  return hit && hit[1].data.value;
}
test.before(async () => {
  for (const f of REAL) {
    if ((await numericCells(fs.readFileSync(f), 2)).length >= 2) { SRC = f; return; }
  }
});

async function repoWith(file) {
  file = file || SRC;
  const root = tmp();
  fs.copyFileSync(file, path.join(root, NAME));
  const repo = await new Repo(root).open();
  return { root, repo, abs: path.join(root, NAME) };
}

test('a repository is a real git repository with the Office merge driver registered', opts, async () => {
  const { root, repo } = await repoWith();
  assert.ok(fs.existsSync(path.join(root, '.git')));
  assert.match(await repo.g.git('config', 'merge.officegit.driver'), /merge-driver\.js/);
  assert.match(fs.readFileSync(path.join(root, '.git', 'info', 'attributes'), 'utf8'), /\*\.xlsx .*merge=officegit/);
  assert.ok(!fs.existsSync(path.join(root, '.gitattributes')), 'the working folder stays free of config files');
  await repo.commit({ message: 'Initial', paths: [NAME], author: 'Ana Lyst' });
  assert.match(await repo.g.git('log', '--format=%an', '-1'), /Ana Lyst/);
  // plain git can read the history this app wrote
  assert.match(await repo.g.git('cat-file', '-t', 'HEAD'), /commit/);
});

test('commit, history and restore: stored bytes are the original file exactly', opts, async () => {
  const { repo, abs } = await repoWith();
  const original = fs.readFileSync(SRC);
  assert.deepEqual((await repo.status()).changes.map((c) => c.status), ['added']);
  const c1 = await repo.commit({ message: 'Initial version', paths: [NAME], author: 'Ana' });
  assert.equal((await repo.status()).changes.length, 0);

  const [cell] = await numericCells(original, 1);
  assert.ok(cell, 'workbook has a plain numeric cell');
  fs.writeFileSync(abs, await setCell(original, cell.addr, Number(cell.value) + 42));

  assert.deepEqual((await repo.status()).changes.map((c) => c.status), ['modified']);
  const changes = (await repo.diffWorking(NAME)).changes.filter((c) => c.kind === 'cell');
  assert.equal(changes.length, 1);
  assert.equal(changes[0].detail, 'input');
  assert.equal(String(changes[0].after.value), String(Number(cell.value) + 42));

  await repo.commit({ message: 'Bump a number', paths: [NAME], author: 'Ana' });
  const log = await repo.log();
  assert.deepEqual(log.map((c) => c.message), ['Bump a number', 'Initial version']);
  assert.deepEqual(log[0].changed, [{ path: NAME, status: 'modified' }]);
  assert.equal((await repo.diffCommit(log[0].id, NAME)).changes.filter((c) => c.kind === 'cell').length, 1);

  await repo.restoreFile(c1.id, NAME);
  assert.equal(sha(fs.readFileSync(abs)), originals.get(SRC), 'restore returns the original bytes');
  assert.ok((await repo.fsck()).ok);
});

test('a re-save with no real edits is not reported as a change', opts, async () => {
  const { repo, abs } = await repoWith();
  await repo.commit({ message: 'base', paths: [NAME], author: 'Ana' });
  const zip = await JSZip.loadAsync(fs.readFileSync(abs));
  zip.file('docProps/core.xml', '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><cp:lastModifiedBy>x</cp:lastModifiedBy></cp:coreProperties>');
  fs.writeFileSync(abs, await zip.generateAsync({ type: 'nodebuffer', compression: 'STORE' }));
  assert.notEqual(sha(fs.readFileSync(abs)), originals.get(SRC), 'bytes really did change');
  const st = await repo.status();
  assert.deepEqual(st.changes, []);
  assert.deepEqual(st.resaved, [NAME]);
});

test('branches: switching rewrites the workbook byte-exactly, dirty switch is refused, stash round-trips', opts, async () => {
  const { repo, abs } = await repoWith();
  const original = fs.readFileSync(SRC);
  const [cell] = await numericCells(original, 1);
  await repo.commit({ message: 'base', paths: [NAME], author: 'Ana' });

  await repo.createAndSwitch('forecast');
  assert.equal((await repo.branches()).current, 'forecast');
  const edited = await setCell(original, cell.addr, 12345);
  fs.writeFileSync(abs, edited);
  await repo.commit({ message: 'Forecast numbers', paths: [NAME], author: 'Ben' });

  await repo.switchBranch('main');
  assert.equal(sha(fs.readFileSync(abs)), sha(original), 'main still has the original');
  await repo.switchBranch('forecast');
  assert.equal(sha(fs.readFileSync(abs)), sha(edited), 'forecast has the edit');

  fs.writeFileSync(abs, await setCell(original, cell.addr, 999));
  assert.ok((await repo.switchBranch('main')).needsAction, 'uncommitted work blocks the switch');
  await repo.switchBranch('main', { stash: true });
  assert.equal((await repo.stashList()).length, 1);
  assert.equal(sha(fs.readFileSync(abs)), sha(original), 'main is clean after stashing');
  await repo.switchBranch('forecast');
  await repo.stashApply((await repo.stashList())[0].id);
  assert.equal((await repo.status()).changes.length, 1, 'stashed work came back');
  assert.equal((await repo.stashList()).length, 0, 'stash entry consumed');

  const n = await repo.unmergedCommits('forecast');
  assert.equal(n, 1, 'forecast has one version that exists nowhere else');
});

test('merge: two people edit different cells -> git merges them with the Office driver', opts, async () => {
  const { repo, abs } = await repoWith();
  const original = fs.readFileSync(SRC);
  const cells = await numericCells(original, 2);
  assert.ok(cells.length >= 2, 'test workbook has two editable numeric cells');
  await repo.commit({ message: 'base', paths: [NAME], author: 'Ana' });

  await repo.createAndSwitch('bens-edits');
  fs.writeFileSync(abs, await setCell(original, cells[1].addr, 777));
  await repo.commit({ message: 'Ben edits ' + cells[1].addr, paths: [NAME], author: 'Ben' });

  await repo.switchBranch('main');
  fs.writeFileSync(abs, await setCell(original, cells[0].addr, 111));
  await repo.commit({ message: 'Ana edits ' + cells[0].addr, paths: [NAME], author: 'Ana' });

  const m = await repo.mergeStart('bens-edits');
  assert.equal(m.status, 'ready', JSON.stringify(m.files));
  const merged = await repo.mergeComplete({}, 'Ana');

  // both edits survive, in one workbook that still opens as a valid package
  assert.equal(String(await cellValue(repo, 'HEAD', cells[0].addr)), '111');
  assert.equal(String(await cellValue(repo, 'HEAD', cells[1].addr)), '777');
  const head = (await repo.log())[0];
  assert.equal(head.parents.length, 2, 'a real git merge commit with two parents');
  assert.ok(merged.id);
  assert.equal((await repo.status()).changes.length, 0, 'working tree is clean after the merge');
  assert.ok((await repo.fsck()).ok);
});

test('merge conflict: same cell edited on both sides -> conflict, then resolved by choosing a side', opts, async () => {
  const { repo, abs } = await repoWith();
  const original = fs.readFileSync(SRC);
  const [cell] = await numericCells(original, 1);
  await repo.commit({ message: 'base', paths: [NAME], author: 'Ana' });

  await repo.createAndSwitch('bens-edits');
  fs.writeFileSync(abs, await setCell(original, cell.addr, 222));
  await repo.commit({ message: 'Ben', paths: [NAME], author: 'Ben' });

  await repo.switchBranch('main');
  fs.writeFileSync(abs, await setCell(original, cell.addr, 333));
  await repo.commit({ message: 'Ana', paths: [NAME], author: 'Ana' });

  const m = await repo.mergeStart('bens-edits');
  assert.equal(m.status, 'conflicts');
  const f = m.files.find((x) => x.path === NAME);
  assert.equal(f.state, 'conflict');
  assert.match(f.summary, /overlapping edit/);
  assert.ok(f.items.some((i) => i.label.endsWith('!' + cell.addr)), JSON.stringify(f.items.slice(0, 3)));
  assert.deepEqual([f.items[0].ours, f.items[0].theirs], ['333', '222'], 'conflict screen shows both values');

  await assert.rejects(repo.mergeComplete({}, 'Ana'), /Choose a version/);
  await repo.mergeComplete({ [NAME]: 'theirs' }, 'Ana');
  assert.equal(String(await cellValue(repo, 'HEAD', cell.addr)), '222', 'their version won');
  assert.equal((await repo.log())[0].parents.length, 2);
  assert.equal((await repo.status()).changes.length, 0);
});

test('merge can be abandoned and leaves the branch untouched', opts, async () => {
  const { repo, abs } = await repoWith();
  const original = fs.readFileSync(SRC);
  const [cell] = await numericCells(original, 1);
  await repo.commit({ message: 'base', paths: [NAME], author: 'Ana' });
  await repo.createAndSwitch('other');
  fs.writeFileSync(abs, await setCell(original, cell.addr, 1));
  await repo.commit({ message: 'other', paths: [NAME], author: 'Ben' });
  await repo.switchBranch('main');
  fs.writeFileSync(abs, await setCell(original, cell.addr, 2));
  const mine = await repo.commit({ message: 'mine', paths: [NAME], author: 'Ana' });

  await repo.mergeStart('other');
  await repo.mergeAbort();
  assert.equal(await repo.g.head(), mine.id, 'branch tip is unchanged');
  assert.equal((await repo.status()).changes.length, 0, 'working tree is clean again');
});

test('every workbook on disk can be committed and read back byte-exactly', opts, async () => {
  const root = tmp();
  for (const f of REAL) fs.copyFileSync(f, path.join(root, path.basename(f)));
  const repo = await new Repo(root).open();
  const names = REAL.map((f) => path.basename(f));
  assert.equal((await repo.status()).changes.length, names.length);
  await repo.commit({ message: 'Add all workbooks', paths: names, author: 'Ana' });
  assert.equal((await repo.status()).changes.length, 0);
  for (const f of REAL) {
    const stored = await repo.readAt('HEAD', path.basename(f));
    assert.equal(sha(stored), originals.get(f), `${path.basename(f)} round-trips byte for byte`);
  }
  assert.deepEqual((await repo.trackedFiles()).sort(), names.sort());
});

test('originals on disk were never modified', opts, () => {
  for (const [f, h] of originals) assert.equal(sha(fs.readFileSync(f)), h, f);
});
