'use strict';
// Processing engine: ingest / diff / merge / render / rebuild.
const I = require('./ingest');
const { diff, summarize } = require('./diff');
const { merge } = require('./merge');
const { rebuild, Unsupported } = require('./rebuild');

// Three-way merge of whole files. Returns { conflicts } when unresolved, { buffer } when done,
// or { unsupported } when the engine can't safely rebuild the merged package.
async function mergeFiles(baseBuf, oursBuf, theirsBuf, resolutions = {}) {
  const [base, ours, theirs] = await Promise.all([I.ingest(baseBuf), I.ingest(oursBuf), I.ingest(theirsBuf)]);
  if (base.format !== ours.format || ours.format !== theirs.format) return { unsupported: 'The file type differs between versions.' };
  const { tree, conflicts } = merge(base, ours, theirs, resolutions);
  if (conflicts.length) return { conflicts };
  try {
    const buffer = await rebuild({ merged: tree, oursTree: ours, theirsTree: theirs, oursBuf, theirsBuf });
    await I.ingest(buffer); // verify the result parses
    return { conflicts: [], buffer };
  } catch (e) {
    if (e instanceof Unsupported) return { unsupported: e.message };
    throw e;
  }
}

async function render() {
  throw new Error('Visual diff needs LibreOffice, which was not found on this computer.');
}

module.exports = { ingest: I.ingest, diff, summarize, merge, mergeFiles, rebuild, render, treeHash: I.treeHash, emptyTree: I.emptyTree, sha: I.sha };
