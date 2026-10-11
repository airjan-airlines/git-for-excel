'use strict';
// rebuild: apply a merged tree onto the "ours" package. Only the changed units are patched, everything
// else is carried over untouched. Changes the engine can't express safely throw Unsupported so the
// caller can fall back to a whole-file conflict instead of writing a risky file.
const JSZip = require('jszip');
const I = require('./ingest');
const X = require('./xml');
const { diff, lcsOps } = require('./diff');

class Unsupported extends Error {}
const unsupported = (m) => new Unsupported(m);

const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const splitAddr = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return { col: colNum(m[1]), row: +m[2] }; };

function cellNode(addr, data, old, before, sigToIdx) {
  const attrs = { '@_r': addr };
  let s = old ? X.attr(old, 's') : undefined;
  if (!before || before.style !== data.style) {
    if (sigToIdx.has(data.style)) s = String(sigToIdx.get(data.style));
    else if (data.style !== 'default' && sigToIdx.size) throw unsupported(`Cell ${addr} uses a format that does not exist in this file`);
  }
  if (s !== undefined) attrs['@_s'] = s;
  const ch = [];
  if (data.formula != null) {
    if (String(data.formula).startsWith('[shared')) throw unsupported(`Cell ${addr} uses a shared formula`);
    ch.push({ f: [{ '#text': data.formula }] });
    if (data.value != null) ch.push({ v: [{ '#text': String(data.value) }] });
    if (data.type === 'str') attrs['@_t'] = 'str'; else if (data.type !== 'n') attrs['@_t'] = data.type;
  } else if (data.type === 'str') {
    attrs['@_t'] = 'inlineStr';
    ch.push({ is: [{ t: [{ '#text': String(data.value) }], ':@': { '@_xml:space': 'preserve' } }] });
  } else {
    if (data.type && data.type !== 'n') attrs['@_t'] = data.type;
    ch.push({ v: [{ '#text': String(data.value) }] });
  }
  return { c: ch, ':@': attrs };
}

async function patchXlsx(zip, oursTree, merged, theirsZip) {
  const changes = diff(oursTree, merged);
  if (!changes.length) return;
  const sheets = Object.fromEntries((await I.sheetPaths(zip)).map((s) => [s.name, s.path]));
  const sigToIdx = await I.styleIndexBySig(zip);
  const bySheet = {};
  let formulaTouched = false;
  const wbChanges = [];
  for (const c of changes) {
    if (c.kind === 'cell') {
      const sheet = c.id.slice(5, c.id.lastIndexOf('!'));
      (bySheet[sheet] ||= []).push(c);
      formulaTouched = true; // any edited cell can change calculated results elsewhere
    } else if (c.kind === 'name') wbChanges.push(c);
    else if (c.kind === 'opaque' && c.type === 'modified') {
      const p = c.id.slice(5);
      if (!theirsZip.file(p)) throw unsupported(`Part ${p} is missing`);
      zip.file(p, await theirsZip.file(p).async('nodebuffer'));
    } else throw unsupported(`${c.label} (${c.type}) can't be merged automatically`);
  }
  for (const [sheet, list] of Object.entries(bySheet)) {
    const p = sheets[sheet];
    const nodes = X.parse(await zip.file(p).async('string'));
    const root = nodes.find((n) => X.tagOf(n) === 'worksheet');
    const sd = X.kids(root).find((n) => X.tagOf(n) === 'sheetData');
    if (!sd) throw unsupported(`Sheet ${sheet} has no cell data`);
    const rows = X.kids(sd);
    I.unshareFormulas(rows);
    for (const c of list) {
      const addr = c.id.slice(c.id.lastIndexOf('!') + 1);
      const { col, row } = splitAddr(addr);
      let rowNode = rows.find((r) => X.tagOf(r) === 'row' && +X.attr(r, 'r') === row);
      if (!rowNode) {
        if (c.type === 'removed') continue;
        rowNode = { row: [], ':@': { '@_r': String(row) } };
        const at = rows.findIndex((r) => X.tagOf(r) === 'row' && +X.attr(r, 'r') > row);
        rows.splice(at < 0 ? rows.length : at, 0, rowNode);
      }
      X.setAttr(rowNode, 'spans', undefined);
      const cells = X.kids(rowNode);
      const idx = cells.findIndex((n) => X.tagOf(n) === 'c' && X.attr(n, 'r') === addr);
      if (c.type === 'removed') { if (idx >= 0) cells.splice(idx, 1); continue; }
      const node = cellNode(addr, c.after, idx >= 0 ? cells[idx] : null, c.before, sigToIdx);
      if (idx >= 0) cells[idx] = node;
      else {
        const at = cells.findIndex((n) => X.tagOf(n) === 'c' && splitAddr(X.attr(n, 'r')).col > col);
        cells.splice(at < 0 ? cells.length : at, 0, node);
      }
    }
    zip.file(p, X.serialize(nodes));
  }
  if (wbChanges.length || formulaTouched) {
    const nodes = X.parse(await zip.file('xl/workbook.xml').async('string'));
    for (const c of wbChanges) {
      const [, scope, name] = /^name:([^:]*):(.*)$/.exec(c.id);
      const dn = X.findAll(nodes, 'definedName').find((n) => X.attr(n, 'name') === name && (X.attr(n, 'localSheetId') ?? '') === scope);
      if (c.type !== 'modified' || !dn) throw unsupported(`Named range ${name} can't be merged automatically`);
      X.kids(dn).splice(0, X.kids(dn).length, { '#text': c.after.ref });
    }
    if (formulaTouched) {
      const calc = X.find(nodes, 'calcPr');
      if (calc) X.setAttr(calc, 'fullCalcOnLoad', '1');
      else X.kids(nodes.find((n) => X.tagOf(n) === 'workbook')).push({ calcPr: [], ':@': { '@_fullCalcOnLoad': '1' } });
      // a stale calcChain triggers Office's repair prompt, so drop it (Excel rebuilds it)
      if (zip.file('xl/calcChain.xml')) {
        zip.remove('xl/calcChain.xml');
        let ct = await zip.file('[Content_Types].xml').async('string');
        ct = ct.replace(/<Override[^>]*calcChain[^>]*\/>/g, '');
        zip.file('[Content_Types].xml', ct);
        const relsP = 'xl/_rels/workbook.xml.rels';
        zip.file(relsP, (await zip.file(relsP).async('string')).replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, ''));
      }
    }
    zip.file('xl/workbook.xml', X.serialize(nodes));
  }
}

function assertPrefixes(xml, rootNode, label) {
  const declared = new Set(Object.keys(rootNode[':@'] || {}).filter((k) => k.startsWith('@_xmlns:')).map((k) => k.slice(8)));
  declared.add('xml');
  for (const m of xml.matchAll(/[<\s]([A-Za-z0-9]+):[A-Za-z]/g)) {
    if (!declared.has(m[1]) && !/^(http|https)$/.test(m[1])) throw unsupported(`${label} uses formatting this document doesn't declare (${m[1]}:)`);
  }
}

async function patchDocx(zip, oursTree, merged, theirsTree) {
  for (const name of new Set([...Object.keys(oursTree.seqs), ...Object.keys(merged.seqs)])) {
    const A = oursTree.seqs[name] || [], B = merged.seqs[name] || [];
    if (A.map((x) => x.h).join() === B.map((x) => x.h).join()) continue;
    const part = Object.keys(zip.files).find((f) => I.WORD_SEQ.test(f) && I.wordSeqName(f) === name);
    if (!part) throw unsupported(`Part for ${name} not found`);
    const nodes = X.normalize(X.parse(await zip.file(part).async('string')));
    const root = nodes.find((n) => X.tagOf(n) !== '?xml');
    const paras = I.collectParas(nodes);
    if (paras.length !== A.length) throw unsupported('Document structure changed while merging');
    const oursKeys = new Set(A.map((x) => x.h));
    let prev = null;
    const pending = [];
    for (const op of lcsOps(A.map((x) => x.h), B.map((x) => x.h))) {
      if (op.t === 'eq') { prev = paras[op.i]; continue; }
      if (op.t === 'del') { const p = paras[op.i]; p.parent.splice(p.parent.indexOf(p.node), 1); continue; }
      const item = B[op.j];
      if (!oursKeys.has(item.h) && /\sr:(id|embed|link)=/.test(item.xml)) throw unsupported('A merged paragraph contains links or images from the other version');
      assertPrefixes(item.xml, root, 'A merged paragraph');
      const node = X.parse(item.xml)[0];
      if (prev) { prev.parent.splice(prev.parent.indexOf(prev.node) + 1, 0, node); prev = { node, parent: prev.parent }; }
      else pending.push(node);
    }
    if (pending.length) {
      const first = paras.find((p) => p.parent.includes(p.node));
      if (!first) throw unsupported('Cannot place paragraph in an empty document');
      first.parent.splice(first.parent.indexOf(first.node), 0, ...pending);
    }
    zip.file(part, X.serialize(nodes));
  }
  for (const c of diff(oursTree, merged)) {
    if (c.kind === 'paragraph') continue;
    if (c.kind === 'opaque' && c.type === 'modified') continue;
    throw unsupported(`${c.label} (${c.type}) can't be merged automatically`);
  }
}

function setShapeText(sp, text) {
  const tx = X.kids(sp).find((c) => X.tagOf(c) === 'p:txBody');
  if (!tx) throw unsupported('Shape has no text frame');
  const kids = X.kids(tx);
  const firstP = kids.find((c) => X.tagOf(c) === 'a:p');
  const pPr = firstP && X.kids(firstP).find((c) => X.tagOf(c) === 'a:pPr');
  const run = firstP && X.findAll(X.kids(firstP), 'a:r')[0];
  const rPr = run && X.kids(run).find((c) => X.tagOf(c) === 'a:rPr');
  const mk = (line) => ({ 'a:p': [...(pPr ? [pPr] : []), { 'a:r': [...(rPr ? [rPr] : []), { 'a:t': [{ '#text': line }] }] }] });
  const keep = kids.filter((c) => X.tagOf(c) !== 'a:p');
  kids.splice(0, kids.length, ...keep, ...text.split('\n').map(mk));
}

async function patchPptx(zip, oursTree, merged) {
  const slidePath = Object.fromEntries((await I.pptxSlides(zip)).map((s) => [s.id, s.path]));
  const edits = {};
  for (const c of diff(oursTree, merged)) {
    if (c.kind === 'opaque' && c.type === 'modified') continue;
    if (c.kind !== 'shape' || c.type !== 'modified' || /format/.test(c.detail)) throw unsupported(`${c.label} (${c.type}) can't be merged automatically`);
    const [, sid, shid] = /^shape:([^/]+)\/(.+)$/.exec(c.id);
    (edits[sid] ||= []).push({ shid, c });
  }
  for (const [sid, list] of Object.entries(edits)) {
    const nodes = X.normalize(X.parse(await zip.file(slidePath[sid]).async('string')));
    const shapes = I.collectShapes(X.kids(X.find(nodes, 'p:spTree')));
    for (const { shid, c } of list) {
      const sp = shapes.find((s) => X.attr(I.shapeProps(s), 'id') === shid);
      if (!sp) throw unsupported(`Shape ${c.label} not found`);
      if (c.before.text !== c.after.text) setShapeText(sp, c.after.text);
      if (c.before.name !== c.after.name) X.setAttr(I.shapeProps(sp), 'name', c.after.name);
      if (c.before.geom !== c.after.geom) {
        const x = X.find(X.kids(sp), 'a:xfrm') || X.find(X.kids(sp), 'p:xfrm');
        const off = x && X.find(X.kids(x), 'a:off'), ext = x && X.find(X.kids(x), 'a:ext');
        if (!off || !ext) throw unsupported('Shape has no explicit position');
        const [ox, oy, cx, cy] = c.after.geom.split(',');
        X.setAttr(off, 'x', ox); X.setAttr(off, 'y', oy); X.setAttr(ext, 'cx', cx); X.setAttr(ext, 'cy', cy);
      }
    }
    zip.file(slidePath[sid], X.serialize(nodes));
  }
  // opaque (master/layout/theme) edits: copy from the other side
}

async function rebuild({ merged, oursTree, theirsTree, oursBuf, theirsBuf }) {
  const zip = await JSZip.loadAsync(oursBuf);
  const theirsZip = await JSZip.loadAsync(theirsBuf);
  const fmt = oursTree.format;
  if (fmt === 'xlsx') await patchXlsx(zip, oursTree, merged, theirsZip);
  else if (fmt === 'docx') await patchDocx(zip, oursTree, merged, theirsTree);
  else await patchPptx(zip, oursTree, merged);
  if (fmt !== 'xlsx') {
    for (const c of diff(oursTree, merged)) {
      if (c.kind === 'opaque' && c.type === 'modified') {
        const p = c.id.slice(5);
        if (!theirsZip.file(p)) throw unsupported(`Part ${p} is missing`);
        zip.file(p, await theirsZip.file(p).async('nodebuffer'));
      }
    }
  }
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

module.exports = { rebuild, Unsupported };
