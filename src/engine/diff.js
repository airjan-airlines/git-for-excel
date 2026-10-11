'use strict';
// diff(treeA, treeB) -> structured change list [{id, kind, type, label, group, before, after, detail, words}]
const { stable, emptyTree } = require('./ingest');

function lcsOps(A, B) {
  const n0 = A.length, m0 = B.length;
  let s = 0;
  while (s < n0 && s < m0 && A[s] === B[s]) s++;
  let e = 0;
  while (e < n0 - s && e < m0 - s && A[n0 - 1 - e] === B[m0 - 1 - e]) e++;
  const ops = [];
  for (let k = 0; k < s; k++) ops.push({ t: 'eq', i: k, j: k });
  const n = n0 - s - e, m = m0 - s - e;
  if (n && m && n * m <= 3e7 && Math.min(n, m) < 65535) {
    const W = m + 1;
    const L = new Uint16Array((n + 1) * W);
    for (let i = n - 1; i >= 0; i--)
      for (let j = m - 1; j >= 0; j--)
        L[i * W + j] = A[s + i] === B[s + j] ? L[(i + 1) * W + j + 1] + 1 : Math.max(L[(i + 1) * W + j], L[i * W + j + 1]);
    let i = 0, j = 0;
    while (i < n && j < m) {
      if (A[s + i] === B[s + j]) { ops.push({ t: 'eq', i: s + i, j: s + j }); i++; j++; }
      else if (L[(i + 1) * W + j] >= L[i * W + j + 1]) { ops.push({ t: 'del', i: s + i }); i++; }
      else { ops.push({ t: 'ins', j: s + j }); j++; }
    }
    while (i < n) ops.push({ t: 'del', i: s + i++ });
    while (j < m) ops.push({ t: 'ins', j: s + j++ });
  } else {
    for (let i = 0; i < n; i++) ops.push({ t: 'del', i: s + i });
    for (let j = 0; j < m; j++) ops.push({ t: 'ins', j: s + j });
  }
  for (let k = 0; k < e; k++) ops.push({ t: 'eq', i: n0 - e + k, j: m0 - e + k });
  return ops;
}

const tokens = (s) => s.match(/\s+|[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu) || [];

function wordDiff(a, b) {
  const A = tokens(a), B = tokens(b);
  const out = [];
  const push = (t, s) => {
    const last = out[out.length - 1];
    if (last && last.t === t) last.s += s; else out.push({ t, s });
  };
  for (const op of lcsOps(A, B)) {
    if (op.t === 'eq') push('eq', A[op.i]);
    else if (op.t === 'del') push('del', A[op.i]);
    else push('ins', B[op.j]);
  }
  return out;
}

function similarity(a, b) {
  if (a === b) return 1;
  const A = tokens(a).filter((t) => /\S/.test(t)), B = tokens(b).filter((t) => /\S/.test(t));
  if (!A.length || !B.length) return 0;
  const set = new Set(A);
  let common = 0;
  for (const t of B) if (set.has(t)) common++;
  return (2 * common) / (A.length + B.length);
}

// index: base index -> other index or -1
function matchArray(base, other) {
  const m = new Array(base.length).fill(-1);
  for (const op of lcsOps(base.map((x) => x.h), other.map((x) => x.h))) if (op.t === 'eq') m[op.i] = op.j;
  return m;
}

const brief = (it) => it && { text: it.text, style: it.style, title: it.title, id: it.h };

function classifyCell(a, b) {
  if (a.formula !== b.formula) return 'formula';
  if (a.value !== b.value) return a.formula ? 'calculated' : 'input';
  if (a.type !== b.type) return 'type';
  return 'format';
}

function diffSeq(name, A, B, out) {
  const ops = lcsOps(A.map((x) => x.h), B.map((x) => x.h));
  const isSlides = name === 'pptx:slides';
  const noun = isSlides ? 'Slide' : 'Paragraph';
  const group = isSlides ? 'Slide order' : name === 'docx:body' ? 'Document body' : name.replace('docx:', '');
  let n = 0;
  const base = (type, i, j, extra = {}) => ({
    id: `${name}:${type}:${i ?? ''}:${j ?? ''}`, kind: isSlides ? 'slide-order' : 'paragraph', type,
    label: `${noun} ${(j ?? i) + 1}`, group,
    before: i != null ? brief(A[i]) : null, after: j != null ? brief(B[j]) : null, ...extra,
  });
  let k = 0;
  while (k < ops.length) {
    if (ops[k].t === 'eq') { k++; continue; }
    const dels = [], ins = [];
    while (k < ops.length && ops[k].t !== 'eq') { (ops[k].t === 'del' ? dels : ins).push(ops[k]); k++; }
    const used = new Set();
    let jp = 0;
    const res = [];
    for (const d of dels) {
      let found = -1;
      if (!isSlides) {
        for (let q = jp; q < ins.length; q++) {
          if (!used.has(q) && similarity(A[d.i].text, B[ins[q].j].text) >= 0.5) { found = q; break; }
        }
      }
      if (found >= 0) {
        used.add(found); jp = found + 1;
        const a = A[d.i], b = B[ins[found].j];
        res.push(base('modified', d.i, ins[found].j, {
          detail: a.text === b.text ? 'format' : 'text',
          words: a.text === b.text ? null : wordDiff(a.text, b.text),
        }));
      } else res.push(base('removed', d.i, null, { label: `${noun} ${d.i + 1}` }));
    }
    ins.forEach((o, q) => { if (!used.has(q)) res.push(base('added', null, o.j)); });
    out.push(...res);
    n += res.length;
  }
  // moved: same content removed in one place and added in another
  const removed = out.filter((c) => c.id.startsWith(name + ':removed') && c.before);
  for (const r of removed) {
    const a = out.find((c) => c.id.startsWith(name + ':added') && c.after && c.after.id === r.before.id);
    if (a) {
      out.splice(out.indexOf(a), 1);
      Object.assign(r, { type: 'moved', after: a.after, label: `${noun} ${r.before ? A.findIndex((x) => x.h === r.before.id) + 1 : ''} → ${B.findIndex((x) => x.h === r.before.id) + 1}` });
    }
  }
}

function diff(a, b) {
  a = a || emptyTree(b.format);
  b = b || emptyTree(a.format);
  const out = [];
  const ids = [...new Set([...Object.keys(a.units), ...Object.keys(b.units)])].sort();
  for (const id of ids) {
    const ua = a.units[id], ub = b.units[id];
    const u = ub || ua;
    const group = u.kind === 'cell' ? id.slice(5, id.lastIndexOf('!')) : u.kind === 'shape' ? u.label.split(' · ')[0] : u.kind === 'notes' ? u.label : u.kind;
    const c = { id, kind: u.kind, label: u.label, group, before: ua ? ua.data : null, after: ub ? ub.data : null };
    if (!ua) c.type = 'added';
    else if (!ub) c.type = 'removed';
    else if (stable(ua.data) !== stable(ub.data)) {
      c.type = 'modified';
      if (u.kind === 'cell') c.detail = classifyCell(ua.data, ub.data);
      else if (u.kind === 'shape') {
        const d = [];
        if (ua.data.text !== ub.data.text) d.push('text');
        if (ua.data.geom !== ub.data.geom) d.push('position');
        if (ua.data.fmt !== ub.data.fmt) d.push('format');
        if (ua.data.name !== ub.data.name) d.push('name');
        c.detail = d.join(', ');
        if (d.includes('text')) c.words = wordDiff(ua.data.text, ub.data.text);
      } else if (u.kind === 'notes' || u.kind === 'comment') {
        c.words = wordDiff(ua.data.text || '', ub.data.text || '');
        c.detail = 'text';
      }
    } else continue;
    out.push(c);
  }
  for (const name of [...new Set([...Object.keys(a.seqs), ...Object.keys(b.seqs)])].sort())
    diffSeq(name, a.seqs[name] || [], b.seqs[name] || [], out);
  return out;
}

function summarize(changes) {
  if (!changes.length) return 'no changes';
  const noun = { cell: 'cell', paragraph: 'paragraph', shape: 'shape', slide: 'slide', 'slide-order': 'slide', sheet: 'sheet', name: 'named range', opaque: 'embedded part', comment: 'comment', notes: 'notes page', sheetmeta: 'sheet layout' };
  const tally = {};
  for (const c of changes) {
    const k = noun[c.kind] || c.kind;
    const t = c.type === 'moved' ? 'moved' : c.type;
    (tally[k] ||= {})[t] = (tally[k][t] || 0) + 1;
  }
  const parts = [];
  for (const [k, v] of Object.entries(tally))
    for (const [t, n] of Object.entries(v)) parts.push(`${n} ${k}${n > 1 ? 's' : ''} ${t}`);
  return parts.join(', ');
}

module.exports = { diff, summarize, lcsOps, wordDiff, matchArray, similarity };
