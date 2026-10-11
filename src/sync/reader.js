'use strict';
// Reads an OfficeGit repository and produces the JSON the Base44 web app stores and renders.
//
// A "source" is anything that can answer four questions about a git repository. Two are provided:
//   localSource(path)     - a clone on this machine (used by the desktop app and the tests)
//   githubSource({...})   - the GitHub REST API (used by the Base44 backend function)
// Everything above the source is identical, so the diff a reviewer sees on the web is the same one the
// desktop app shows.
const engine = require('../engine');

const OFFICE = /\.(xlsx|docx|pptx)$/i;
const PAYLOAD_VERSION = 1;

// ---------------------------------------------------------------- local git
function localSource(repoPath) {
  const { Git } = require('../main/git');
  const g = new Git(repoPath);
  const S = '\u001e';
  return {
    kind: 'local',
    id: repoPath,
    async branches() {
      const out = await g.git('for-each-ref', '--format=%(refname:short)' + S + '%(objectname)' + S + '%(committerdate:iso-strict)', 'refs/heads');
      const head = await g.branch();
      return out.split('\n').filter(Boolean).map((l) => {
        const [name, sha, date] = l.split(S);
        return { name, head: sha, updatedAt: date, isDefault: name === head };
      });
    },
    async commits(branch, limit = 100) {
      const fmt = ['%H', '%P', '%an', '%ae', '%aI', '%s', '%b'].join(S);
      const out = await g.git('log', `--max-count=${limit}`, `--format=${S}${fmt}`, '--name-status', '-M', '-z', branch);
      return parseLog(out, S);
    },
    async readFile(ref, path) { return g.raw('show', `${ref}:${path}`); },
  };
}

function parseLog(out, S) {
  const parts = out.split('\0');
  const commits = [];
  let cur = null;
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i].replace(/^\n+/, '');
    if (!p) continue;
    if (p.startsWith(S)) {
      const [, sha, parents, author, email, at, subject, body] = p.split(S);
      cur = { sha, parents: parents ? parents.split(' ') : [], author, email, createdAt: at, message: subject, description: (body || '').trim(), files: [] };
      commits.push(cur);
      continue;
    }
    if (!cur || !/^[AMDRCT]\d*$/.test(p.split('\n')[0])) continue;
    const letter = p[0];
    if (letter === 'R' || letter === 'C') {
      const from = parts[++i], to = parts[++i];
      if (OFFICE.test(to)) cur.files.push({ path: to, status: 'renamed', previousPath: from });
    } else {
      const f = parts[++i];
      if (f && OFFICE.test(f)) cur.files.push({ path: f, status: { A: 'added', M: 'modified', D: 'deleted', T: 'modified' }[letter] || 'modified' });
    }
  }
  return commits;
}

// ---------------------------------------------------------------- GitHub REST
// token: a fine-grained PAT or GitHub App installation token with Contents: read on the repository.
function githubSource({ owner, repo, token, fetch: f = globalThis.fetch, api = 'https://api.github.com' }) {
  const call = async (url, accept = 'application/vnd.github+json') => {
    const res = await f(url.startsWith('http') ? url : `${api}${url}`, {
      headers: { accept, authorization: `Bearer ${token}`, 'x-github-api-version': '2022-11-28', 'user-agent': 'officegit' },
    });
    if (!res.ok) {
      const body = await res.text().catch(() => '');
      const e = new Error(`GitHub ${res.status}: ${body.slice(0, 200)}`);
      e.status = res.status;
      throw e;
    }
    return res;
  };
  const json = async (url) => (await call(url)).json();
  const slug = `/repos/${owner}/${repo}`;

  return {
    kind: 'github',
    id: `${owner}/${repo}`,
    async branches() {
      const [list, meta] = await Promise.all([json(`${slug}/branches?per_page=100`), json(slug)]);
      return list.map((b) => ({ name: b.name, head: b.commit.sha, updatedAt: null, isDefault: b.name === meta.default_branch }));
    },
    async commits(branch, limit = 100) {
      const list = await json(`${slug}/commits?sha=${encodeURIComponent(branch)}&per_page=${Math.min(limit, 100)}`);
      // the list endpoint omits per-commit file lists, so fetch each commit for them
      return Promise.all(list.map(async (c) => {
        const full = await json(`${slug}/commits/${c.sha}`);
        return {
          sha: c.sha,
          parents: (c.parents || []).map((p) => p.sha),
          author: c.commit.author.name,
          email: c.commit.author.email,
          createdAt: c.commit.author.date,
          message: c.commit.message.split('\n')[0],
          description: c.commit.message.split('\n').slice(1).join('\n').trim(),
          files: (full.files || []).filter((x) => OFFICE.test(x.filename)).map((x) => ({
            path: x.filename,
            status: { added: 'added', removed: 'deleted', modified: 'modified', renamed: 'renamed' }[x.status] || 'modified',
            ...(x.previous_filename ? { previousPath: x.previous_filename } : {}),
          })),
        };
      }));
    },
    // Office files are far past GitHub's 1 MB contents-API limit, so go through the blob API.
    async readFile(ref, filePath) {
      const node = await json(`${slug}/contents/${filePath.split('/').map(encodeURIComponent).join('/')}?ref=${encodeURIComponent(ref)}`);
      const blob = await json(`${slug}/git/blobs/${node.sha}`);
      if (blob.encoding !== 'base64') throw new Error(`Unexpected blob encoding ${blob.encoding}`);
      return Buffer.from(blob.content, 'base64');
    },
  };
}

// ---------------------------------------------------------------- payloads
// One commit as the Base44 `Commit` entity expects it.
const toCommitEntity = (repositoryId, branch, c) => ({
  payloadVersion: PAYLOAD_VERSION,
  repositoryId,
  sha: c.sha,
  parentShas: c.parents,
  branch,
  author: c.author,
  authorEmail: c.email,
  message: c.message,
  description: c.description,
  createdAt: c.createdAt,
  isMerge: c.parents.length > 1,
  files: c.files,
});

async function listBranches(source, repositoryId) {
  return (await source.branches()).map((b) => ({ payloadVersion: PAYLOAD_VERSION, repositoryId, ...b }));
}

async function listCommits(source, repositoryId, branch, limit) {
  return (await source.commits(branch, limit)).map((c) => toCommitEntity(repositoryId, branch, c));
}

// The semantic diff for one file in one commit: what the reviewer actually reads.
async function getDiff(source, { sha, path: filePath, parentSha }) {
  const commits = await source.commits(sha, 1);
  const commit = commits[0];
  if (!commit) throw new Error(`Unknown commit ${sha}`);
  const entry = commit.files.find((f) => f.path === filePath);
  const parent = parentSha || commit.parents[0] || null;

  const read = async (ref, p) => { try { return await source.readFile(ref, p); } catch { return null; } };
  const after = entry && entry.status === 'deleted' ? null : await read(sha, filePath);
  const before = parent ? await read(parent, (entry && entry.previousPath) || filePath) : null;
  if (!before && !after) throw new Error(`${filePath} is not in commit ${sha}`);

  const [a, b] = await Promise.all([before && engine.ingest(before), after && engine.ingest(after)]);
  const changes = engine.diff(a || null, b || null);
  return {
    payloadVersion: PAYLOAD_VERSION,
    sha,
    parentSha: parent,
    path: filePath,
    format: (b || a).format,
    status: entry ? entry.status : 'modified',
    summary: engine.summarize(changes),
    changeCount: changes.length,
    changes,
  };
}

module.exports = { localSource, githubSource, listBranches, listCommits, getDiff, toCommitEntity, PAYLOAD_VERSION, OFFICE };
