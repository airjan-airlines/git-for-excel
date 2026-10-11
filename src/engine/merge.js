'use strict';
// merge(base, ours, theirs, resolutions) -> { tree, conflicts }
// Unit maps merge per unit; sequences (paragraphs, slide order) merge with diff3.
// Anything overlapping becomes a conflict; nothing is silently dropped.
const { stable, short } = require('./ingest');
const X = require('./xml');
const { matchArray } = require('./diff');

// A formula cell's cached value is derived (Excel recalculates it), so only the formula and style count as edits.
const cmp = (u) => (u.kind === 'cell' && u.data.formula != null ? { formula: u.data.formula, style: u.data.style } : u.data);
const same = (a, b) => (a === undefined || b === undefined ? a === b : stable(cmp(a)) === stable(cmp(b)));
const segKey = (s) => s.map((x) => x.h).join('|');

function diff3(base, ours, theirs) {
  const mo = matchArray(base, ours), mt = matchArray(base, theirs);
  const chunks = [];
  let bi = 0, oi = 0, ti = 0;
  for (let k = 0; k <= base.length; k++) {
    const end = k === base.length;
    if (!end && !(mo[k] >= 0 && mt[k] >= 0)) continue;
    const oe = end ? ours.length : mo[k], te = end ? theirs.length : mt[k];
    const c = { base: base.slice(bi, k), ours: ours.slice(oi, oe), theirs: theirs.slice(ti, te) };
    if (c.base.length || c.ours.length || c.theirs.length) chunks.push(c);
    if (!end) { chunks.push({ stable: ours[mo[k]] }); bi = k + 1; oi = mo[k] + 1; ti = mt[k] + 1; }
  }
  return chunks;
}

function makeParaItem(text, template) {
  const node = X.parse(template.xml)[0];
  const kids = X.kids(node);
  const pPr = kids.find((c) => X.tagOf(c) === 'w:pPr');
  const run = X.findAll(kids, 'w:r')[0];
  const rPr = run && X.kids(run).find((c) => X.tagOf(c) === 'w:rPr');
  const r = { 'w:r': [...(rPr ? [rPr] : []), { 'w:t': [{ '#text': text }], ':@': { '@_xml:space': 'preserve' } }] };
  node[X.tagOf(node)] = [...(pPr ? [pPr] : []), r];
  const xml = X.serialize([node]);
  return { h: short(xml), text, style: template.style, xml };
}

const stripXml = (items) => items.map((i) => ({ text: i.text, style: i.style, title: i.title, id: i.h }));

function merge(base, ours, theirs, resolutions = {}) {
  const tree = { schema: ours.schema, format: ours.format, units: {}, seqs: {} };
  const conflicts = [];

  const ids = new Set([...Object.keys(base.units), ...Object.keys(ours.units), ...Object.keys(theirs.units)]);
  for (const id of [...ids].sort()) {
    const b = base.units[id], o = ours.units[id], t = theirs.units[id];
    let pick, conflict = false;
    if (same(o, t)) pick = o;
    else if (same(o, b)) pick = t;
    else if (same(t, b)) pick = o;
    else { conflict = true; pick = o; }
    if (conflict) {
      const r = resolutions[id];
      if (r === 'theirs') pick = t;
      else if (r === 'base') pick = b;
      else if (r === 'ours') pick = o;
      else if (r && r.data) {
        const src = o || t || b;
        pick = { ...src, data: { ...src.data, ...r.data } };
      } else {
        const u = o || t || b;
        conflicts.push({ id, scope: 'unit', kind: u.kind, label: u.label, base: b && b.data, ours: o && o.data, theirs: t && t.data });
      }
    }
    if (pick) tree.units[id] = pick;
  }

  const names = new Set([...Object.keys(base.seqs), ...Object.keys(ours.seqs), ...Object.keys(theirs.seqs)]);
  for (const name of [...names].sort()) {
    const out = [];
    diff3(base.seqs[name] || [], ours.seqs[name] || [], theirs.seqs[name] || []).forEach((c, n) => {
      if (c.stable) { out.push(c.stable); return; }
      const ko = segKey(c.ours), kt = segKey(c.theirs), kb = segKey(c.base);
      if (ko === kt || kt === kb) { out.push(...c.ours); return; }
      if (ko === kb) { out.push(...c.theirs); return; }
      // same number of paragraphs on all sides: edits are in place, so merge position by position
      if (c.base.length === c.ours.length && c.ours.length === c.theirs.length && !resolutions[`${name}#${n}`]) {
        const rows = c.base.map((b, i) => (c.ours[i].h === c.theirs[i].h || c.theirs[i].h === b.h ? c.ours[i] : c.ours[i].h === b.h ? c.theirs[i] : null));
        if (rows.every(Boolean)) { out.push(...rows); return; }
      }
      const id = `${name}#${n}`;
      const r = resolutions[id];
      if (r === 'theirs') out.push(...c.theirs);
      else if (r === 'base') out.push(...c.base);
      else if (r === 'ours') out.push(...c.ours);
      else if (r && typeof r.text === 'string') {
        const tpl = c.ours[0] || c.theirs[0] || c.base[0];
        if (!tpl || !tpl.xml) out.push(...c.ours);
        else out.push(...r.text.split('\n').map((line) => makeParaItem(line, tpl)));
      } else {
        out.push(...c.ours);
        const isSlides = name === 'pptx:slides';
        conflicts.push({
          id, scope: 'seq', kind: isSlides ? 'slide-order' : 'paragraph',
          label: isSlides ? 'Slide order' : 'Paragraphs', base: stripXml(c.base), ours: stripXml(c.ours), theirs: stripXml(c.theirs),
          editable: !isSlides,
        });
      }
    });
    tree.seqs[name] = out;
  }
  return { tree, conflicts };
}

module.exports = { merge };
