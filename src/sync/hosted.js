'use strict';
// Base44 as the git host.
//
// Base44 is not a git server, so the repository travels as a git bundle (one file holding every branch, commit and
// workbook version), stored in Base44 file storage and pointed at by Repository.bundle_url. Locally, a bare mirror at
// .git/officegit/base44.git stands in for the server and is registered as the git remote "base44", so push, pull,
// ahead/behind and the Office merge driver all work exactly as they do with any other remote:
//
//   fetch   download the bundle  -> git fetch it into the mirror -> git fetch base44
//   push    fetch (as above)     -> git push base44 <branch>  (git refuses if you are behind)
//           -> bundle the mirror -> upload -> Repository.bundle_url -> Commit/Branch records for the web app
//   pull    fetch (as above)     -> git pull base44 <branch>   (merge driver handles overlapping workbooks)
const fs = require('fs');
const fsp = fs.promises;
const os = require('os');
const path = require('path');
const { run } = require('../main/git');
const gitxl = require('./gitxl');

const REMOTE = 'base44';
const KEY = 'officegit.base44';

const bare = (mirror, ...args) => run(mirror, args);

async function mirrorPath(repo) {
  const gitDir = path.resolve(repo.root, await repo.g.git('rev-parse', '--git-dir'));
  return path.join(gitDir, 'officegit', 'base44.git');
}

async function ensureMirror(repo) {
  const mirror = await mirrorPath(repo);
  if (!fs.existsSync(path.join(mirror, 'HEAD'))) {
    await fsp.mkdir(mirror, { recursive: true });
    await bare(mirror, 'init', '--bare', '-q');
  }
  // The mirror sits deep inside .git, so Windows' 260-character path limit is easy to hit.
  await bare(mirror, 'config', 'core.longpaths', 'true');
  await repo.g.git('config', '--local', 'core.longpaths', 'true');
  const remotes = await repo.remotes();
  if (!remotes.some((r) => r.name === REMOTE)) await repo.g.git('remote', 'add', REMOTE, mirror);
  return mirror;
}

/** Which Base44 repository this folder is linked to, or null. */
async function linked(repo) {
  const r = await repo.g.try('config', '--local', `${KEY}.repositoryId`);
  if (!r.ok || !r.out) return null;
  const name = await repo.g.try('config', '--local', `${KEY}.name`);
  return { repositoryId: r.out, name: name.ok ? name.out : null, remote: REMOTE };
}

/** Repositories hosted in the Base44 app, for the "connect" picker. */
async function listRemote(client) {
  const rows = await client.list('Repository', { sort: '-updated_date', limit: 200 });
  return (Array.isArray(rows) ? rows : []).map((r) => ({ id: r.id, name: r.name, owner: r.owner_name, hasData: !!r.bundle_url, defaultBranch: r.default_branch || 'main' }));
}

/**
 * Link this folder to a Base44 repository: an existing one (repositoryId) or a new one (name).
 * Linking to an existing repository downloads it; if this folder has no history yet, its default branch is checked out.
 */
async function link(repo, client, { repositoryId, name, owner }) {
  let record;
  if (repositoryId) record = await client.get('Repository', repositoryId);
  else {
    if (!name || !name.trim()) throw new Error('Give the repository a name.');
    record = await client.create('Repository', {
      name: name.trim(), description: '', owner_name: owner || 'OfficeGit user',
      visibility: 'private', default_branch: (await repo.g.branch()) || 'main', file_type: 'xlsx',
    });
  }
  await repo.g.git('config', '--local', `${KEY}.repositoryId`, record.id);
  await repo.g.git('config', '--local', `${KEY}.name`, record.name);
  await ensureMirror(repo);
  const got = await fetch(repo, client);

  // A fresh folder (a "clone"): check out the hosted default branch so the workbooks appear.
  if (got.branches.length && !(await repo.g.hasCommits())) {
    const branch = got.branches.includes(record.default_branch) ? record.default_branch : got.branches[0];
    await repo.g.git('checkout', '-B', branch, '--track', `${REMOTE}/${branch}`);
  }
  return { repositoryId: record.id, name: record.name, branches: got.branches, cloned: got.branches.length > 0 };
}

async function requireLink(repo) {
  const l = await linked(repo);
  if (!l) throw new Error('This repository is not connected to Base44 yet.');
  return l;
}

/** Download the hosted bundle into the mirror, then update the base44/* remote-tracking branches. */
async function fetch(repo, client) {
  const { repositoryId } = await requireLink(repo);
  const mirror = await ensureMirror(repo);
  const record = await client.get('Repository', repositoryId);
  const url = record.bundle_url || null;
  if (url) {
    const tmp = path.join(os.tmpdir(), `officegit-${process.pid}-${Date.now()}.bundle`);
    await fsp.writeFile(tmp, await client.download(url));
    try {
      // The server copy is the truth: force-update and prune, so a rejected push of ours is rolled back too.
      await bare(mirror, 'fetch', '--prune', '-q', tmp, '+refs/heads/*:refs/heads/*');
    } finally { await fsp.rm(tmp, { force: true }); }
  }
  await repo.g.git('config', '--local', `${KEY}.bundle`, url || '');
  await repo.g.git('fetch', '--prune', REMOTE);
  const heads = await bare(mirror, 'for-each-ref', '--format=%(refname:short)', 'refs/heads').catch(() => '');
  return { url, branches: heads.split('\n').filter(Boolean) };
}

/** Push the current branch to Base44. */
async function push(repo, client, { owner, onProgress } = {}) {
  const { repositoryId } = await requireLink(repo);
  const before = await fetch(repo, client);
  const branch = await repo.g.branch();
  if (!(await repo.g.hasCommits())) throw new Error('Commit something first, then push.');

  const r = await repo.g.try('push', '--set-upstream', REMOTE, branch);
  if (!r.ok) {
    if (/non-fast-forward|fetch first|\(rejected\)/i.test(r.error)) throw new Error('Base44 has versions you don’t have yet. Pull first, then push again.');
    throw new Error(r.error);
  }

  // Pack the whole hosted repository into one file and publish it.
  const mirror = await ensureMirror(repo);
  const tmp = path.join(os.tmpdir(), `officegit-${process.pid}-${Date.now()}.bundle`);
  await bare(mirror, 'bundle', 'create', '-q', tmp, '--branches');
  let url;
  try { url = await client.upload(await fsp.readFile(tmp), `${repositoryId}.bundle`); } finally { await fsp.rm(tmp, { force: true }); }

  // Someone else may have pushed while we were packing; never overwrite their work.
  const latest = await client.get('Repository', repositoryId);
  if ((latest.bundle_url || null) !== before.url) {
    await fetch(repo, client); // roll the mirror back to what is really hosted
    throw new Error('Someone else pushed to Base44 at the same moment. Pull, then push again.');
  }
  await client.update('Repository', repositoryId, { bundle_url: url });
  await repo.g.git('config', '--local', `${KEY}.bundle`, url);

  // The web app's view: commit records with cell-level diffs, and the branch pointer.
  const meta = await gitxl.push(repo, client, { repositoryId, only: [branch], owner, onProgress });
  return { branch, remote: REMOTE, url, commitsPushed: meta.commitsPushed };
}

module.exports = { REMOTE, linked, listRemote, link, fetch, push, mirrorPath };
