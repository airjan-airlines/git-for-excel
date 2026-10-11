'use strict';
// Builds a demo project whose history has a clear pattern: every monthly close updates the same three inputs.
//   node scripts/demo-patterns.js <empty-folder> [source.xlsx] [--share "Project name"]
// Then edit just ONE of the three cells in Excel: GitXL suggests the other two ("you usually also update…").
// One version also overwrites a formula with a typed number, to show the model-risk flag in History.
const fs = require('fs');
const os = require('os');
const path = require('path');
const JSZip = require('jszip');
const { Repo } = require('../src/main/repo');
const engine = require('../src/engine');
const { rowLabel } = require('../src/main/patterns');

const args = process.argv.slice(2);
const flag = (f) => { const i = args.indexOf(f); return i >= 0 ? args.splice(i, 2)[1] : undefined; };
const shareAs = flag('--share');
const [dir, srcArg] = args;
if (!dir) throw new Error('usage: demo-patterns.js <empty-folder> [source.xlsx] [--share "Project name"]');
const SRC = srcArg || path.join(os.homedir(), 'Downloads', 'TexMUN_Budget_FY2627.xlsx');
const NAME = 'Model.xlsx';

const sheets = (z) => Object.keys(z.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();

// Replace one cell's content with a plain number, whether it held a value or a formula.
async function setNumber(file, sheetPath, addr, value) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const xml = await zip.file(sheetPath).async('string');
  const re = new RegExp(`<c r="${addr}"([^>]*?)(/>|>[\\s\\S]*?</c>)`);
  const m = re.exec(xml);
  if (!m) throw new Error(`cell ${addr} not found`);
  const attrs = m[1].replace(/\s+t="[^"]*"/, '');
  zip.file(sheetPath, xml.replace(m[0], `<c r="${addr}"${attrs}><v>${value}</v></c>`));
  fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

(async () => {
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, NAME);
  fs.copyFileSync(SRC, file);

  // Pick a sheet with three labelled number inputs, plus a formula cell to hardcode later.
  const tree = await engine.ingest(fs.readFileSync(file));
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const sheetNames = (await require('../src/engine/ingest').sheetPaths(zip));
  let pick = null;
  for (const { name, path: sp } of sheetNames) {
    const cells = Object.entries(tree.units).filter(([id, u]) => u.kind === 'cell' && id.startsWith(`cell:${name}!`));
    const seen = new Set(); // one cell per line item, so the suggestion names three different things
    const inputs = cells.filter(([id, u]) => u.data.formula == null && u.data.type === 'n' && rowLabel(tree, id))
      .map(([id]) => id).filter((id) => { const l = rowLabel(tree, id); if (seen.has(l)) return false; seen.add(l); return true; });
    const formula = cells.find(([, u]) => u.data.formula && !u.data.formula.startsWith('[shared') && u.data.type === 'n');
    if (inputs.length >= 4 && formula && sp) { pick = { sheet: name, sheetPath: sp, inputs: inputs.slice(0, 3), formula: formula[0] }; break; }
  }
  if (!pick) throw new Error('no sheet with three number inputs and a formula found in ' + SRC);
  const addr = (id) => id.slice(id.lastIndexOf('!') + 1);
  const label = (id) => `${rowLabel(tree, id)} (${addr(id)})`;

  const repo = await new Repo(dir).open();
  await repo.commit({ message: 'Add the FY26-27 model', paths: [NAME], author: 'Priya Raman' });
  const months = ['July', 'August', 'September', 'October'];
  const authors = ['Priya Raman', 'Marcus Webb', 'Priya Raman', 'Lena Park'];
  for (let i = 0; i < months.length; i++) {
    for (const [k, id] of pick.inputs.entries()) await setNumber(file, pick.sheetPath, addr(id), 100 * (i + 2) + 17 * k + 5);
    await repo.commit({ message: `${months[i]} close: update actuals`, description: 'Monthly actuals from the GL export.', paths: [NAME], author: authors[i] });
    if (i === 1) {
      await setNumber(file, pick.sheetPath, addr(pick.formula), 2500);
      await repo.commit({ message: 'Quick fix before the board pack', paths: [NAME], author: 'Marcus Webb' });
    }
  }

  console.log(`\nDemo project ready in ${dir}`);
  console.log(`Sheet "${pick.sheet}". Every monthly close updated:`);
  for (const id of pick.inputs) console.log(`   ${label(id)}`);
  console.log(`\nDemo: open ${NAME} in Excel, change ONLY ${label(pick.inputs[0])}, save and close.`);
  console.log('GitXL will suggest the other two. The "Quick fix" version in History shows the formula-overwrite flag.');

  if (shareAs) {
    const hosted = require('../src/sync/hosted');
    const client = require('../src/sync/base44').fromEnv();
    await hosted.link(repo, client, { name: shareAs, owner: 'Priya Raman' });
    await hosted.push(repo, client, { owner: 'Priya Raman' });
    console.log(`\nShared to GitXL as "${shareAs}".`);
  }
})().catch((e) => { console.error(e.message); process.exit(1); });
