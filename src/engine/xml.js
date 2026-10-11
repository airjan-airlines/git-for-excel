'use strict';
// Thin helpers over fast-xml-parser's order-preserving tree.
const { XMLParser, XMLBuilder } = require('fast-xml-parser');

const opts = {
  preserveOrder: true,
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  trimValues: false,
  parseTagValue: false,
  parseAttributeValue: false,
};
const parser = new XMLParser(opts);
const builder = new XMLBuilder({ ...opts, suppressEmptyNode: true, format: false });

const parse = (s) => parser.parse(s);
const serialize = (nodes) => builder.build(nodes);

const tagOf = (n) => Object.keys(n).find((k) => k !== ':@');
const kids = (n) => n[tagOf(n)];
const attr = (n, a) => (n[':@'] ? n[':@']['@_' + a] : undefined);
const setAttr = (n, a, v) => {
  if (!n[':@']) n[':@'] = {};
  if (v === undefined) delete n[':@']['@_' + a];
  else n[':@']['@_' + a] = v;
};

function textOf(nodes) {
  let s = '';
  for (const n of nodes) {
    const t = tagOf(n);
    if (t === '#text') s += n['#text'];
    else s += textOf(n[t]);
  }
  return s;
}

function findAll(nodes, tag, out = []) {
  for (const n of nodes) {
    const t = tagOf(n);
    if (t === '#text') continue;
    if (t === tag) out.push(n);
    findAll(n[t], tag, out);
  }
  return out;
}
const find = (nodes, tag) => findAll(nodes, tag)[0];
const child = (n, tag) => kids(n).find((c) => tagOf(c) === tag);
const children = (n, tag) => kids(n).filter((c) => tagOf(c) === tag);

// Volatile data that changes on a no-op save: rsids, paraIds, spell-check marks.
const VOLATILE_ATTR = /^@_(w:rsid\w*|w14:paraId|w14:textId|xr:uid|xr\d+:uid|dirty|err)$/;
const VOLATILE_TAG = new Set(['w:proofErr', 'w:rsids']);
function normalize(nodes) {
  for (let i = nodes.length - 1; i >= 0; i--) {
    const n = nodes[i];
    const t = tagOf(n);
    if (t === '#text') continue;
    if (VOLATILE_TAG.has(t)) {
      nodes.splice(i, 1);
      continue;
    }
    if (n[':@']) for (const k of Object.keys(n[':@'])) if (VOLATILE_ATTR.test(k)) delete n[':@'][k];
    normalize(n[t]);
  }
  return nodes;
}

// Word/Excel/PowerPoint renumber relationship ids (rId10 -> rId6) on save; compare what they point to instead.
const REL_ATTR = /^@_r:(id|embed|link|pict|href|dm|lo|qs|cs)$/;
function resolveRels(nodes, rels) {
  for (const n of nodes) {
    const t = tagOf(n);
    if (t === '#text') continue;
    if (n[':@']) for (const k of Object.keys(n[':@'])) if (REL_ATTR.test(k)) { const r = rels[n[':@'][k]]; n[':@'][k] = 'rel:' + (r ? r.target : n[':@'][k]); }
    resolveRels(n[t], rels);
  }
  return nodes;
}

module.exports = { resolveRels, parse, serialize, tagOf, kids, attr, setAttr, textOf, findAll, find, child, children, normalize };
