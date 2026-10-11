'use strict';
// Builds a demo repository from real workbooks so you can try branching and merging:
//   node scripts/demo-setup.js <folder-with-xlsx> <demo-folder>
// Creates: main with one edit, a branch with a different edit -> ready to merge in the app.
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { Repo } = require('../src/main/repo');

const PLAIN_CELL = /<c r="([A-Z]+\d+)"(?![^>]*\bt=)[^>]*><v>(-?[\d.]+)<\/v><\/c>/g;
const sheetsOf = (zip) => Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort();

async function plainCells(file, want = 2) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const out = [];
  for (const s of sheetsOf(zip)) {
    const xml = await zip.file(s).async('string');
    for (const m of xml.matchAll(PLAIN_CELL)) {
      out.push({ sheet: s, addr: m[1], value: m[2] });
      if (out.length >= want) return out;
    }
  }
  return out;
}

async function setCell(file, addr, value) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  for (const s of sheetsOf(zip)) {
    const xml = await zip.file(s).async('string');
    const re = new RegExp(`<c r="${addr}"(?![^>]*\\bt=)[^>]*><v>(-?[\\d.]+)</v></c>`);
    const m = re.exec(xml);
    if (!m) continue;
    zip.file(s, xml.replace(m[0], m[0].replace(`<v>${m[1]}</v>`, `<v>${value}</v>`)));
    fs.writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
    return m[1];
  }
  throw new Error(`no plain numeric cell ${addr} in ${path.basename(file)}`);
}

(async () => {
  const [srcDir, demo] = process.argv.slice(2);
  if (!srcDir || !demo) throw new Error('usage: demo-setup.js <folder-with-xlsx> <demo-folder>');
  fs.rmSync(demo, { recursive: true, force: true });
  fs.mkdirSync(demo, { recursive: true });

  const picked = [];
  for (const n of fs.readdirSync(srcDir).filter((x) => /\.xlsx$/i.test(x) && !x.startsWith('~$'))) {
    const src = path.join(srcDir, n);
    if ((await plainCells(src, 2)).length < 2) continue;
    picked.push(src);
    if (picked.length === 2) break;
  }
  if (!picked.length) throw new Error(`no editable workbooks found in ${srcDir}`);
  const names = picked.map((f, i) => (i === 0 ? 'Budget.xlsx' : 'Plan.xlsx'));
  picked.forEach((f, i) => fs.copyFileSync(f, path.join(demo, names[i])));

  const repo = await new Repo(demo).open();
  await repo.commit({ message: 'Add the workbooks we track', paths: names, author: 'Arya Chakrabarti' });

  const budget = path.join(demo, names[0]);
  const cells = await plainCells(budget, 2);

  await repo.createAndSwitch('travel-budget-revision');
  await setCell(budget, cells[1].addr, 9999);
  await repo.commit({ message: 'Raise the travel line for the spring conference', paths: [names[0]], author: 'Priya Raman' });

  await repo.switchBranch('main');
  await setCell(budget, cells[0].addr, 4242);
  await repo.commit({ message: 'Update the registration income estimate', paths: [names[0]], author: 'Arya Chakrabarti' });

  console.log(`demo ready at ${demo}`);
  console.log(`  files:    ${names.join(', ')}`);
  console.log(`  edited:   ${cells[0].addr} on main, ${cells[1].addr} on travel-budget-revision`);
  console.log(`  branches: ${(await repo.branches()).list.map((b) => b.name + (b.current ? ' (current)' : '')).join(', ')}`);
})().catch((e) => { console.error(e.message); process.exit(1); });
