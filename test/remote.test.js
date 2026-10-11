'use strict';
// Two people sharing one workbook through a git remote. The remote here is a local bare repo, which is the
// same protocol git uses for GitHub/GitLab/Gitea - only the URL differs.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');
const JSZip = require('jszip');
const { Repo } = require('../src/main/repo');
const { run } = require('../src/main/git');

const DIR = process.env.OG_XLSX_DIR || path.join(os.homedir(), 'Downloads');
const PLAIN = /<c r="([A-Z]+\d+)"(?![^>]*\bt=)[^>]*><v>(-?[\d.]+)<\/v><\/c>/g;
const sheets = (z) => Object.keys(z.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();

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
  throw new Error('cell not found: ' + addr);
}
async function cellValue(repo, rev, name, addr) {
  const t = await repo.treeAt(rev, name);
  const hit = Object.entries(t.units).find(([k, u]) => u.kind === 'cell' && k.endsWith('!' + addr));
  return hit && hit[1].data.value;
}

let SRC = null;
const tmp = () => fs.mkdtempSync(path.join(os.tmpdir(), 'og-remote-'));
const NAME = 'Budget.xlsx';
const opts = { skip: fs.existsSync(DIR) && fs.readdirSync(DIR).some((n) => /\.xlsx$/i.test(n)) ? false : `no .xlsx files in ${DIR}` };

test.before(async () => {
  for (const n of fs.readdirSync(DIR).filter((x) => /\.xlsx$/i.test(x) && !x.startsWith('~$'))) {
    const f = path.join(DIR, n);
    if ((await cellsOf(fs.readFileSync(f), 2)).length >= 2) { SRC = f; return; }
  }
});

test('two people share a workbook through a git remote: push, clone, pull', opts, async () => {
  assert.ok(SRC, 'found a workbook with two editable cells');
  const [cellA, cellB] = await cellsOf(fs.readFileSync(SRC), 2);

  // the shared remote (a bare repo - identical protocol to a GitHub URL)
  const remote = path.join(tmp(), 'team.git');
  await run(path.dirname(remote), ['init', '--bare', '-b', 'main', remote]);

  // Ana sets up, commits and publishes
  const anaDir = tmp();
  fs.copyFileSync(SRC, path.join(anaDir, NAME));
  const ana = await new Repo(anaDir).open();
  await ana.commit({ message: 'Add the budget', paths: [NAME], author: 'Ana' });
  assert.deepEqual((await ana.syncState()).remotes, [], 'no remote yet');
  await ana.setRemote('origin', remote);
  let st = await ana.syncState();
  assert.equal(st.upstream, null);
  assert.equal(st.publishable, true, 'app offers to publish');
  await ana.push('origin', { setUpstream: true });
  st = await ana.syncState();
  assert.equal(st.upstream, 'origin/main');
  assert.deepEqual([st.ahead, st.behind], [0, 0], 'in sync after publishing');

  // Ben clones the same remote and gets the workbook, byte for byte
  const benDir = tmp();
  await run(path.dirname(benDir), ['clone', remote, benDir]);
  const ben = await new Repo(benDir).open();
  assert.equal(crypto.createHash('sha256').update(fs.readFileSync(path.join(benDir, NAME))).digest('hex'),
    crypto.createHash('sha256').update(fs.readFileSync(SRC)).digest('hex'), 'cloned file matches the original');

  // Ben edits one cell and pushes
  await setCell(path.join(benDir, NAME), cellB, 8888);
  await ben.commit({ message: 'Ben updates the travel line', paths: [NAME], author: 'Ben' });
  await ben.push('origin');

  // Ana edits a different cell, then pulls Ben's work: the merge driver combines both
  await setCell(path.join(anaDir, NAME), cellA, 1111);
  await ana.commit({ message: 'Ana updates registration income', paths: [NAME], author: 'Ana' });
  st = await ana.syncState();
  assert.equal(st.ahead, 1, 'one version to send');

  const pulled = await ana.pull('origin');
  assert.deepEqual(pulled.conflicts, [], 'different cells merge without asking');
  assert.equal(String(await cellValue(ana, 'HEAD', NAME, cellA)), '1111', 'Ana’s edit survived');
  assert.equal(String(await cellValue(ana, 'HEAD', NAME, cellB)), '8888', 'Ben’s edit survived');
  assert.equal((await ana.status()).changes.length, 0, 'working tree clean after pull');

  await ana.push('origin');
  st = await ana.syncState();
  assert.deepEqual([st.ahead, st.behind], [0, 0], 'everything shared');

  // and the remote really has it: a fresh clone sees both edits
  const checkDir = tmp();
  await run(path.dirname(checkDir), ['clone', remote, checkDir]);
  const check = await new Repo(checkDir).open();
  assert.equal(String(await cellValue(check, 'HEAD', NAME, cellA)), '1111');
  assert.equal(String(await cellValue(check, 'HEAD', NAME, cellB)), '8888');
});

test('pulling an overlapping edit raises a conflict the user can resolve', opts, async () => {
  assert.ok(SRC);
  const [cellA] = await cellsOf(fs.readFileSync(SRC), 1);
  const remote = path.join(tmp(), 'team.git');
  await run(path.dirname(remote), ['init', '--bare', '-b', 'main', remote]);

  const anaDir = tmp();
  fs.copyFileSync(SRC, path.join(anaDir, NAME));
  const ana = await new Repo(anaDir).open();
  await ana.commit({ message: 'base', paths: [NAME], author: 'Ana' });
  await ana.setRemote('origin', remote);
  await ana.push('origin', { setUpstream: true });

  const benDir = tmp();
  await run(path.dirname(benDir), ['clone', remote, benDir]);
  const ben = await new Repo(benDir).open();
  await setCell(path.join(benDir, NAME), cellA, 222);
  await ben.commit({ message: 'Ben', paths: [NAME], author: 'Ben' });
  await ben.push('origin');

  // Ana changes the same cell
  await setCell(path.join(anaDir, NAME), cellA, 333);
  await ana.commit({ message: 'Ana', paths: [NAME], author: 'Ana' });

  const pulled = await ana.pull('origin');
  assert.equal(pulled.conflicts.length, 1);
  assert.match(pulled.conflicts[0].summary, /overlapping edit/);
  assert.ok(pulled.conflicts[0].items.some((i) => i.label.endsWith('!' + cellA)));

  await ana.mergeComplete({ [NAME]: 'theirs' }, 'Ana');
  assert.equal(String(await cellValue(ana, 'HEAD', NAME, cellA)), '222', 'chose their version');
  assert.equal((await ana.status()).changes.length, 0);
});

test('a push that would overwrite someone else’s work is refused with a clear message', opts, async () => {
  assert.ok(SRC);
  const [cellA] = await cellsOf(fs.readFileSync(SRC), 1);
  const remote = path.join(tmp(), 'team.git');
  await run(path.dirname(remote), ['init', '--bare', '-b', 'main', remote]);
  const anaDir = tmp();
  fs.copyFileSync(SRC, path.join(anaDir, NAME));
  const ana = await new Repo(anaDir).open();
  await ana.commit({ message: 'base', paths: [NAME], author: 'Ana' });
  await ana.setRemote('origin', remote);
  await ana.push('origin', { setUpstream: true });

  const benDir = tmp();
  await run(path.dirname(benDir), ['clone', remote, benDir]);
  const ben = await new Repo(benDir).open();
  await setCell(path.join(benDir, NAME), cellA, 777001);
  await ben.commit({ message: 'Ben', paths: [NAME], author: 'Ben' });
  await ben.push('origin');

  await setCell(path.join(anaDir, NAME), cellA, 777002);
  await ana.commit({ message: 'Ana', paths: [NAME], author: 'Ana' });
  await assert.rejects(ana.push('origin'), /Pull first/);
});
