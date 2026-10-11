'use strict';
// ingest(file) -> semantic tree:  { schema, format, units: {id: {kind,label,data}}, seqs: {name: [item]} }
const crypto = require('crypto');
const path = require('path');
const JSZip = require('jszip');
const X = require('./xml');

const SCHEMA = 1;
const sha = (s) => crypto.createHash('sha256').update(s).digest('hex');
const short = (s) => sha(s).slice(0, 16);

function stable(o) {
  if (o === null || typeof o !== 'object') return JSON.stringify(o) ?? 'null';
  if (Array.isArray(o)) return '[' + o.map(stable).join(',') + ']';
  return '{' + Object.keys(o).filter((k) => o[k] !== undefined).sort()
    .map((k) => JSON.stringify(k) + ':' + stable(o[k])).join(',') + '}';
}

const resolveTarget = (from, target) =>
  target.startsWith('/') ? target.slice(1) : path.posix.normalize(path.posix.join(path.posix.dirname(from), target));

async function readRels(zip, source) {
  const relsPath = path.posix.join(path.posix.dirname(source), '_rels', path.posix.basename(source) + '.rels');
  const f = zip.file(relsPath);
  const map = {};
  if (!f) return map;
  for (const r of X.findAll(X.parse(await f.async('string')), 'Relationship')) {
    map[X.attr(r, 'Id')] = { target: resolveTarget(source, X.attr(r, 'Target')), type: X.attr(r, 'Type') || '' };
  }
  return map;
}

const SKIP_COMMON = /^(\[Content_Types\]\.xml|docProps\/(core|app)\.xml|docProps\/thumbnail\..*|.*\.rels)$/i;

const CACHE_TAGS = new Set(['c:numCache', 'c:strCache', 'c:multiLvlStrCache']);
function stripCaches(nodes) {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const t = X.tagOf(nodes[i]);
    if (t === '#text') continue;
    if (CACHE_TAGS.has(t)) nodes.splice(i, 1); else stripCaches(nodes[i][t]);
  }
  return nodes;
}

async function opaqueHash(zip, name) {
  const buf = await zip.file(name).async('nodebuffer');
  if (/\.(xml|rels)$/i.test(name)) {
    try {
      let nodes = X.normalize(X.parse(buf.toString('utf8')));
      nodes = X.resolveRels(nodes, await readRels(zip, name));
      if (/charts\/[^/]*\.xml$/i.test(name)) stripCaches(nodes); // cached values are derived from the cells
      return short(X.serialize(nodes));
    } catch { /* fall through to raw bytes */ }
  }
  return short(buf);
}

function addOpaque(tree, name, hash, label) {
  tree.units['part:' + name] = { kind: 'opaque', label: label || `Part ${name}`, data: { hash } };
}

function detectFormat(zip) {
  if (zip.file('xl/workbook.xml')) return 'xlsx';
  if (zip.file('word/document.xml')) return 'docx';
  if (zip.file('ppt/presentation.xml')) return 'pptx';
  return null;
}

// ---------- Excel ----------
const siText = (nodes) => {
  let s = '';
  for (const n of nodes) {
    const t = X.tagOf(n);
    if (t === '#text' || t === 'rPh') continue;
    s += t === 't' ? X.textOf(n.t) : siText(n[t]);
  }
  return s;
};

async function sheetPaths(zip) {
  const wb = X.parse(await zip.file('xl/workbook.xml').async('string'));
  const rels = await readRels(zip, 'xl/workbook.xml');
  return X.findAll(wb, 'sheet').map((n) => ({ name: X.attr(n, 'name'), path: rels[X.attr(n, 'r:id')]?.target }));
}

// Excel renumbers its style table (cellXfs) on save, so a cell's style index means nothing across versions.
// Resolve each index to a signature of what it actually looks like.
async function styleInfo(zip) {
  const empty = { sigs: [], other: '' };
  if (!zip.file('xl/styles.xml')) return empty;
  const root = X.normalize(X.parse(await zip.file('xl/styles.xml').async('string'))).find((n) => X.tagOf(n) === 'styleSheet');
  if (!root) return empty;
  const K = X.kids(root);
  const section = (t) => K.find((n) => X.tagOf(n) === t);
  const items = (t, c) => (section(t) ? X.kids(section(t)).filter((n) => X.tagOf(n) === c) : []);
  const ser = (n) => (n ? X.serialize([n]) : '');
  const fmts = {};
  for (const nf of items('numFmts', 'numFmt')) fmts[X.attr(nf, 'numFmtId')] = X.attr(nf, 'formatCode');
  const fonts = items('fonts', 'font').map(ser), fills = items('fills', 'fill').map(ser), borders = items('borders', 'border').map(ser);
  const sigs = items('cellXfs', 'xf').map((xf) => {
    const id = X.attr(xf, 'numFmtId') || '0';
    const sub = (t) => ser(X.kids(xf).find((c) => X.tagOf(c) === t));
    return short(JSON.stringify([fmts[id] ?? 'builtin:' + id, fonts[+X.attr(xf, 'fontId') || 0], fills[+X.attr(xf, 'fillId') || 0], borders[+X.attr(xf, 'borderId') || 0], sub('alignment'), sub('protection')]));
  });
  const skip = new Set(['numFmts', 'fonts', 'fills', 'borders', 'cellXfs']);
  const other = short(X.serialize(K.filter((n) => !skip.has(X.tagOf(n)))));
  return { sigs, other };
}
// signature -> first cellXfs index (used when a merged cell needs a style that exists in this file)
async function styleIndexBySig(zip) {
  const m = new Map();
  (await styleInfo(zip)).sigs.forEach((sig, i) => { if (!m.has(sig)) m.set(sig, i); });
  return m;
}

async function ingestXlsx(zip, tree) {
  const modeled = new Set(['xl/workbook.xml', 'xl/sharedStrings.xml', 'xl/calcChain.xml', 'xl/styles.xml']);
  const styles = await styleInfo(zip);
  tree.units.styles = { kind: 'styles', label: 'Styles & number formats', data: { hash: short(JSON.stringify([[...new Set(styles.sigs)].sort(), styles.other])) } };
  const ss = [];
  if (zip.file('xl/sharedStrings.xml')) {
    for (const si of X.findAll(X.parse(await zip.file('xl/sharedStrings.xml').async('string')), 'si')) ss.push(siText(X.kids(si)));
  }
  const wb = X.parse(await zip.file('xl/workbook.xml').async('string'));
  for (const dn of X.findAll(wb, 'definedName')) {
    const name = X.attr(dn, 'name');
    const scope = X.attr(dn, 'localSheetId');
    tree.units[`name:${scope ?? ''}:${name}`] = { kind: 'name', label: `Named range ${name}`, data: { ref: X.textOf(X.kids(dn)) } };
  }
  for (const { name, path: p } of await sheetPaths(zip)) {
    if (!p || !zip.file(p)) continue;
    modeled.add(p);
    tree.units['sheet:' + name] = { kind: 'sheet', label: `Sheet "${name}"`, data: {} };
    const nodes = X.parse(await zip.file(p).async('string'));
    const root = nodes.find((n) => X.tagOf(n) === 'worksheet');
    const rest = X.resolveRels(X.kids(root).filter((n) => X.tagOf(n) !== 'sheetData' && X.tagOf(n) !== 'dimension'), await readRels(zip, p));
    for (const sf of X.findAll(rest, 'sheetFormatPr')) { X.setAttr(sf, 'defaultRowHeight', undefined); X.setAttr(sf, 'x14ac:dyDescent', undefined); }
    for (const col of X.findAll(rest, 'col')) { const w = X.attr(col, 'width'); if (w != null) X.setAttr(col, 'width', Number(w).toFixed(1)); }
    for (const sv of X.findAll(rest, 'sheetView')) {
      X.setAttr(sv, 'tabSelected', undefined);
      X.setAttr(sv, 'topLeftCell', undefined); // where the sheet was scrolled to: view state, not content
      const k = X.kids(sv);
      for (let i = k.length - 1; i >= 0; i--) if (X.tagOf(k[i]) === 'selection') k.splice(i, 1);
    }
    tree.units['sheetmeta:' + name] = { kind: 'sheetmeta', label: `Sheet "${name}" layout & formatting`, data: { hash: short(X.serialize(X.normalize(rest))) } };
    const sd = X.kids(root).find((n) => X.tagOf(n) === 'sheetData');
    if (!sd) continue;
    const shared = {}; // shared-formula index -> its master cell
    for (const row of X.kids(sd)) {
      if (X.tagOf(row) !== 'row') continue;
      for (const c of X.kids(row)) {
        if (X.tagOf(c) !== 'c') continue;
        const addr = X.attr(c, 'r');
        const t = X.attr(c, 't');
        let v = null, f = null, is = null;
        for (const ch of X.kids(c)) {
          const ct = X.tagOf(ch);
          if (ct === 'v') v = X.textOf(X.kids(ch));
          else if (ct === 'f') f = sharedFormula(shared, ch, addr);
          else if (ct === 'is') is = siText(X.kids(ch));
        }
        let value = v;
        if (t === 's' && v != null) value = ss[+v] ?? '';
        else if (t === 'inlineStr') value = is ?? '';
        if (value == null && f == null) continue;
        const type = t === 's' || t === 'inlineStr' || t === 'str' ? 'str' : t || 'n';
        tree.units[`cell:${name}!${addr}`] = {
          kind: 'cell', label: `${name}!${addr}`,
          data: { value, formula: f, type, style: styles.sigs[+(X.attr(c, 's') || 0)] || 'default' },
        };
      }
    }
  }
  return modeled;
}

// ---------- Word ----------
function wtext(nodes) {
  let s = '';
  for (const n of nodes) {
    const t = X.tagOf(n);
    if (t === '#text') continue;
    if (t === 'w:t') s += X.textOf(n[t]);
    else if (t === 'w:tab') s += '\t';
    else if (t === 'w:br' || t === 'w:cr') s += '\n';
    else if (t === 'w:del' || t === 'w:instrText' || t === 'w:delText') continue;
    else s += wtext(n[t]);
  }
  return s;
}

function collectParas(nodes, out = []) {
  for (const n of nodes) {
    const t = X.tagOf(n);
    if (t === '#text') continue;
    if (t === 'w:p') out.push({ node: n, parent: nodes });
    else collectParas(n[t], out);
  }
  return out;
}

function paraItem(node, rels = {}) {
  const xml = X.serialize([node]);
  const resolved = /\sr:(id|embed|link)=/.test(xml) ? X.serialize(X.resolveRels(X.parse(xml), rels)) : xml;
  const pPr = X.kids(node).find((c) => X.tagOf(c) === 'w:pPr');
  const ps = pPr && X.kids(pPr).find((c) => X.tagOf(c) === 'w:pStyle');
  return { h: short(resolved), text: wtext(X.kids(node)), style: (ps && X.attr(ps, 'w:val')) || '', xml };
}

const WORD_SEQ = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;
const wordSeqName = (p) => (p === 'word/document.xml' ? 'docx:body' : 'docx:' + path.posix.basename(p, '.xml'));

async function ingestDocx(zip, tree) {
  const modeled = new Set();
  for (const name of Object.keys(zip.files)) {
    if (WORD_SEQ.test(name)) {
      modeled.add(name);
      const nodes = X.normalize(X.parse(await zip.file(name).async('string')));
      const rels = await readRels(zip, name);
      tree.seqs[wordSeqName(name)] = collectParas(nodes).map((p) => paraItem(p.node, rels));
    } else if (name === 'word/comments.xml') {
      modeled.add(name);
      for (const c of X.findAll(X.parse(await zip.file(name).async('string')), 'w:comment')) {
        const id = X.attr(c, 'w:id');
        tree.units['comment:' + id] = {
          kind: 'comment', label: `Comment ${id}`,
          data: { author: X.attr(c, 'w:author') || '', text: wtext(X.kids(c)) },
        };
      }
    } else if (/^word\/(settings|webSettings|fontTable)\.xml$/.test(name)) {
      modeled.add(name); // volatile (rsids, view state)
    }
  }
  return modeled;
}

// ---------- PowerPoint ----------
const aText = (nodes) => {
  const paras = [];
  for (const n of nodes) {
    const t = X.tagOf(n);
    if (t === 'a:p') {
      let s = '';
      (function r(ns) {
        for (const c of ns) {
          const ct = X.tagOf(c);
          if (ct === 'a:t') s += X.textOf(c['a:t']);
          else if (ct === 'a:br') s += '\n';
          else if (ct !== '#text') r(c[ct]);
        }
      })(n[t]);
      paras.push(s);
    } else if (t !== '#text') {
      const inner = aText(n[t]);
      if (inner) paras.push(inner);
    }
  }
  return paras.join('\n');
};

const SHAPES = new Set(['p:sp', 'p:pic', 'p:graphicFrame', 'p:cxnSp', 'p:grpSp']);
function collectShapes(nodes, out = []) {
  for (const n of nodes) {
    const t = X.tagOf(n);
    if (t === '#text') continue;
    if (SHAPES.has(t)) {
      out.push(n);
      if (t === 'p:grpSp') collectShapes(n[t], out);
    } else if (t === 'mc:AlternateContent' || t === 'p:spTree') collectShapes(n[t], out);
  }
  return out;
}
const shapeProps = (n) => {
  const nv = X.kids(n).find((c) => /^p:nv/.test(X.tagOf(c)));
  return nv && X.find(X.kids(nv), 'p:cNvPr');
};
function shapeGeom(n) {
  const x = X.find(X.kids(n), 'a:xfrm') || X.find(X.kids(n), 'p:xfrm');
  if (!x) return '';
  const off = X.find(X.kids(x), 'a:off'), ext = X.find(X.kids(x), 'a:ext');
  return [off && X.attr(off, 'x'), off && X.attr(off, 'y'), ext && X.attr(ext, 'cx'), ext && X.attr(ext, 'cy')].join(',');
}
const shapeText = (n) => {
  const tx = X.kids(n).find((c) => X.tagOf(c) === 'p:txBody') || (X.tagOf(n) === 'p:graphicFrame' ? n : null);
  return tx ? aText(X.kids(tx)) : '';
};

async function pptxSlides(zip) {
  const pres = X.parse(await zip.file('ppt/presentation.xml').async('string'));
  const rels = await readRels(zip, 'ppt/presentation.xml');
  return X.findAll(pres, 'p:sldId').map((n) => ({ id: X.attr(n, 'id'), path: rels[X.attr(n, 'r:id')]?.target })).filter((s) => s.path);
}

async function ingestPptx(zip, tree) {
  const modeled = new Set(['ppt/presentation.xml', 'ppt/presProps.xml', 'ppt/viewProps.xml', 'ppt/tableStyles.xml']);
  const slides = await pptxSlides(zip);
  tree.seqs['pptx:slides'] = [];
  let n = 0;
  for (const s of slides) {
    n++;
    if (!zip.file(s.path)) continue;
    modeled.add(s.path);
    const nodes = X.normalize(X.parse(await zip.file(s.path).async('string')));
    const tree0 = X.find(nodes, 'p:spTree');
    let title = '';
    if (tree0) {
      for (const sh of collectShapes(X.kids(tree0))) {
        const props = shapeProps(sh);
        if (!props) continue;
        const sid = X.attr(props, 'id');
        const ph = X.find(X.kids(sh), 'p:ph');
        const text = shapeText(sh);
        if (ph && /title/i.test(X.attr(ph, 'type') || '') && !title) title = text;
        const spPr = X.kids(sh).find((c) => X.tagOf(c) === 'p:spPr');
        tree.units[`shape:${s.id}/${sid}`] = {
          kind: 'shape', label: `Slide ${n} · ${X.attr(props, 'name') || 'Shape'}`,
          data: {
            type: X.tagOf(sh).slice(2), name: X.attr(props, 'name') || '', text, geom: shapeGeom(sh),
            fmt: spPr ? short(X.serialize([spPr])) : '',
          },
        };
      }
    }
    tree.seqs['pptx:slides'].push({ h: s.id, title });
    tree.units['slide:' + s.id] = { kind: 'slide', label: `Slide ${n}`, data: {} };
    const rels = await readRels(zip, s.path);
    const notes = Object.values(rels).find((r) => /notesSlide$/.test(r.type));
    if (notes && zip.file(notes.target)) {
      modeled.add(notes.target);
      const nn = X.parse(await zip.file(notes.target).async('string'));
      const body = collectShapes(X.kids(X.find(nn, 'p:spTree') || { x: [] }) || []).find((sh) => {
        const ph = X.find(X.kids(sh), 'p:ph');
        return ph && X.attr(ph, 'type') === 'body';
      });
      tree.units['notes:' + s.id] = { kind: 'notes', label: `Slide ${n} notes`, data: { text: body ? shapeText(body) : '' } };
    }
  }
  return modeled;
}

async function ingest(buf) {
  let zip;
  try {
    zip = await JSZip.loadAsync(buf);
  } catch {
    throw new Error('This file is not a valid Office package (could not open as a zip).');
  }
  const format = detectFormat(zip);
  if (!format) throw new Error('Unsupported file: only .xlsx, .docx and .pptx are supported.');
  const tree = { schema: SCHEMA, format, units: {}, seqs: {} };
  let modeled;
  try {
    modeled = await { xlsx: ingestXlsx, docx: ingestDocx, pptx: ingestPptx }[format](zip, tree);
  } catch (e) {
    throw new Error(`Could not parse part of this file (${e.message}).`);
  }
  for (const name of Object.keys(zip.files)) {
    if (zip.files[name].dir || modeled.has(name) || SKIP_COMMON.test(name) || /\.vml$/i.test(name)) continue;
    addOpaque(tree, name, await opaqueHash(zip, name));
  }
  return tree;
}

// ---------- shared formulas ----------
// Excel stores a block of copied formulas once, on a "master" cell, and the other cells only point at it by group
// index (<f t="shared" si="2"/>). Excel renumbers those groups on save, so the index is not content: expand every
// cell to its own formula, exactly as Excel shows it in the formula bar.
const colNum = (s) => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colName = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
const splitCell = (a) => { const m = /^([A-Z]+)(\d+)$/.exec(a); return { col: colNum(m[1]), row: +m[2] }; };

// Move the relative (non-$) references in a formula by dr rows and dc columns; text in quotes is left alone.
function shiftFormula(f, dr, dc) {
  return f.split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/).map((part, i) => (i % 2 ? part : part.replace(
    /(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_(])/g,
    (m, dcol, col, drow, row) => `${dcol}${dcol ? col : colName(colNum(col) + dc)}${drow}${drow ? row : +row + dr}`,
  ))).join('');
}

// The formula of one <f> element; `shared` remembers each group's master as the sheet is read top to bottom.
function sharedFormula(shared, fNode, addr) {
  const text = X.textOf(X.kids(fNode));
  const si = X.attr(fNode, 'si');
  if (text) {
    if (si != null && X.attr(fNode, 't') === 'shared') shared[si] = { addr, text };
    return text;
  }
  if (si == null) return '';
  const m = shared[si];
  if (!m) return `[shared ${si}]`;
  const a = splitCell(m.addr), b = splitCell(addr);
  return shiftFormula(m.text, b.row - a.row, b.col - a.col);
}

// Rewrite a sheet's shared formulas as ordinary per-cell formulas (used before patching cells into it, so changing
// a master cell can never strand the cells that pointed at it). Excel reads both forms identically.
function unshareFormulas(rows) {
  const shared = {};
  for (const row of rows) {
    if (X.tagOf(row) !== 'row') continue;
    for (const c of X.kids(row)) {
      if (X.tagOf(c) !== 'c') continue;
      const fNode = X.kids(c).find((n) => X.tagOf(n) === 'f');
      if (!fNode || X.attr(fNode, 't') !== 'shared') continue;
      const f = sharedFormula(shared, fNode, X.attr(c, 'r'));
      if (f.startsWith('[shared')) continue;
      X.kids(fNode).splice(0, X.kids(fNode).length, { '#text': f });
      for (const k of ['t', 'ref', 'si']) X.setAttr(fNode, k, undefined);
    }
  }
}

const treeHash = (tree) => sha(stable(tree));
const emptyTree = (format) => ({ schema: SCHEMA, format, units: {}, seqs: {} });

module.exports = {
  ingest, treeHash, emptyTree, stable, sha, short, SCHEMA,
  // shared with rebuild
  readRels, sheetPaths, shiftFormula, unshareFormulas, styleIndexBySig, collectParas, paraItem, wtext, WORD_SEQ, wordSeqName,
  collectShapes, shapeProps, pptxSlides, resolveTarget,
};
