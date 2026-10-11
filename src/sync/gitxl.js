'use strict';
// Pushes a local git repository into the GitXL Base44 app.
//
// GitXL's entities (discovered from the live app):
//   Repository  name, description, owner_name, visibility, default_branch, file_type, color
//   Branch      repository_id, repository_name, name, head_commit_id, protected, created_by
//   Commit      repository_id, hash, parent_hash, branch, author, author_email, message,
//               ai_summary, files[{path,change_type,summary}], changes[], stats{additions,deletions}
//   Member      repository_id, repository_name, user_email, user_name, role
//
// A GitXL `change` is one semantic unit, and its unit_id is the same id the engine produces
// ("cell:P&L!D18"), so the web app's per-cell comments line up with the desktop diff.
const path = require('path');
const engine = require('../engine');

const SHORT = 7; // GitXL stores 7-character hashes
const short = (sha) => (sha || '').slice(0, SHORT);

// engine change.detail -> which property of the cell actually moved
const PROPERTY = { formula: 'formula', input: 'value', calculated: 'value', type: 'value', format: 'format', text: 'text' };

const asText = (side, property) => {
  if (side == null) return null;
  if (property === 'formula') return side.formula != null ? `=${side.formula}` : (side.value != null ? String(side.value) : null);
  if (property === 'format') return side.style != null ? String(side.style) : null;
  if (side.text != null) return side.text;
  if (side.value != null) return String(side.value);
  if (side.ref != null) return String(side.ref);
  if (side.hash != null) return 'binary content';
  return null;
};

// One engine change -> one GitXL change row.
function mapChange(filePath, c) {
  const format = (filePath.split('.').pop() || '').toLowerCase();
  const property = PROPERTY[c.detail] || (c.kind === 'cell' ? 'value' : 'content');
  return {
    file_path: filePath,
    file_type: format,
    unit_id: c.id,
    unit_label: c.label,
    unit_type: c.kind,
    change_type: c.type === 'moved' ? 'modified' : c.type,
    before: asText(c.before, property),
    after: asText(c.after, property),
    property,
  };
}

// A modified unit counts as one line out and one line in, the way the web app's sample data does.
const countStats = (changes) => ({
  additions: changes.filter((c) => c.change_type === 'added' || c.change_type === 'modified').length,
  deletions: changes.filter((c) => c.change_type === 'removed' || c.change_type === 'deleted' || c.change_type === 'modified').length,
});

// "Merge branch 'x' into y", "Merge x into y", "Merge remote-tracking branch 'base44/x'" -> x
const mergeSource = (message) => {
  const m = /^Merge (?:remote-tracking )?(?:branch )?'?([^'\s]+)'?/.exec(message || '');
  return m ? m[1].replace(/^(base44|origin)\//, '') : null;
};

// Build the Commit payload for one local commit, including the semantic diff of every file it touched.
async function buildCommit(repo, commit, repositoryId, { maxChangesPerCommit = 500 } = {}) {
  const files = [];
  let changes = [];
  for (const f of commit.changed) {
    let summary = f.status;
    if (f.status !== 'deleted') {
      try {
        const d = await repo.diffCommit(commit.id, f.path);
        summary = engine.summarize(d.changes);
        changes.push(...d.changes.map((c) => mapChange(f.path, c)));
      } catch (e) { summary = `could not be compared (${e.message})`; }
    }
    files.push({ path: f.path, change_type: f.status === 'renamed' ? 'modified' : f.status, summary });
  }
  const truncated = changes.length > maxChangesPerCommit;
  if (truncated) changes = changes.slice(0, maxChangesPerCommit);
  return {
    repository_id: repositoryId,
    hash: short(commit.id),
    parent_hash: short(commit.parents[0]) || null,
    second_parent_hash: short(commit.parents[1]) || null,
    is_merge: commit.parents.length > 1,
    merge_source_branch: commit.parents.length > 1 ? mergeSource(commit.message) : null,
    branch: commit.branch,
    author: commit.author,
    author_email: commit.email || `${String(commit.author).replace(/\s+/g, '.').toLowerCase()}@officegit.local`,
    message: commit.message + (truncated ? ` (showing first ${maxChangesPerCommit} changes)` : ''),
    ai_summary: null,
    files,
    changes,
    stats: countStats(changes),
  };
}

// Find an existing record by a natural key without relying on the schema endpoint.
async function findBy(client, entity, key, limit = 200) {
  const rows = await client.list(entity, { q: key, limit });
  return (Array.isArray(rows) ? rows : []).find(Boolean) || null;
}

/**
 * Push every branch and commit of `repo` into GitXL.
 * Existing commits (matched on repository_id + hash) are left alone, so this is safe to re-run.
 */
async function push(repo, client, { name, owner, description, repositoryId: knownId, only, onProgress = () => {} } = {}) {
  let branches = (await repo.branches()).list.filter((b) => b.tip !== null); // an unborn branch has nothing to push
  if (only) branches = branches.filter((b) => only.includes(b.name));
  const defaultBranch = branches.some((b) => b.name === 'main') ? 'main' : (branches[0] || {}).name || 'main';

  // 1. the Repository record: the one we were given, else found by name, else a new one
  let repository = knownId ? await client.get('Repository', knownId) : await findBy(client, 'Repository', { name: name || path.basename(repo.root) });
  const repoName = repository ? repository.name : name || path.basename(repo.root);
  if (!repository) {
    repository = await client.create('Repository', {
      name: repoName,
      description: description || `Synced from ${repo.root}`,
      owner_name: owner || 'OfficeGit user',
      visibility: 'private',
      default_branch: defaultBranch,
      file_type: 'xlsx',
    });
    onProgress({ step: 'repository', created: true, name: repoName });
  } else {
    onProgress({ step: 'repository', created: false, name: repoName });
  }
  const repositoryId = repository.id;

  // 2. commits, oldest first so parents exist before children
  const existing = await client.list('Commit', { q: { repository_id: repositoryId }, limit: 5000 });
  const known = new Set((Array.isArray(existing) ? existing : []).map((c) => c.hash));
  const recordIdByHash = new Map((Array.isArray(existing) ? existing : []).map((c) => [c.hash, c.id]));

  let pushed = 0;
  for (const branch of branches) {
    const log = (await repo.log(branch.name)).reverse();
    for (const c of log) {
      if (known.has(short(c.id))) continue;
      const payload = await buildCommit(repo, { ...c, branch: branch.name }, repositoryId);
      const created = await client.create('Commit', payload);
      known.add(payload.hash);
      recordIdByHash.set(payload.hash, created.id);
      pushed++;
      onProgress({ step: 'commit', hash: payload.hash, message: c.message, changes: payload.changes.length, pushed });
    }
  }

  // 3. branches, pointing at the GitXL record id of their tip commit
  for (const b of branches) {
    const tip = await repo.g.try('rev-parse', b.name);
    const headRecord = tip.ok ? recordIdByHash.get(short(tip.out)) : null;
    const data = {
      repository_id: repositoryId,
      repository_name: repoName,
      name: b.name,
      head_commit_id: headRecord || null,
      protected: b.name === (repository.default_branch || defaultBranch),
      created_by: owner || 'OfficeGit user',
    };
    const found = await findBy(client, 'Branch', { repository_id: repositoryId, name: b.name });
    if (found) await client.update('Branch', found.id, data);
    else await client.create('Branch', data);
    onProgress({ step: 'branch', name: b.name });
  }

  return { repositoryId, repositoryName: repoName, branches: branches.length, commitsPushed: pushed, commitsTotal: known.size };
}

module.exports = { push, buildCommit, mapChange, countStats, mergeSource, short };
