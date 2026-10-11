// Base44 backend function: reads an OfficeGit repository from GitHub and serves it to the web app.
//
// Runtime: Deno. Deploy with `npx base44 functions deploy` from the project root.
// Secrets (set in the app dashboard or `npx base44 secrets set`):
//   GITHUB_TOKEN  - fine-grained PAT (or GitHub App installation token) with Contents: read
//
// Actions (POST JSON  { action, ... }):
//   { action: 'syncBranches',  repositoryId }                  -> upserts Branch entities
//   { action: 'syncCommits',   repositoryId, branch, limit }   -> upserts Commit entities
//   { action: 'getDiff',       repositoryId, sha, path }       -> semantic diff (cached in CommitDiff)
//
// The heavy lifting lives in ../../../src/sync/reader.js and ../../../src/engine, copied into this folder
// at deploy time (see base44/README.md) so the exact same diff runs on the desktop and on the web.
import { createClientFromRequest } from 'npm:@base44/sdk';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const reader = require('./lib/sync/reader.js');

const json = (body, status = 200) => Response.json(body, { status });
const fail = (status, code, message) => json({ error: { code, message } }, status);

async function sourceFor(base44, repositoryId) {
  const repo = await base44.asServiceRole.entities.Repository.get(repositoryId);
  if (!repo) throw Object.assign(new Error(`No repository ${repositoryId}`), { status: 404, code: 'not-found' });
  const token = await Deno.env.get('GITHUB_TOKEN');
  if (!token) throw Object.assign(new Error('GITHUB_TOKEN is not configured'), { status: 500, code: 'no-token' });
  if (!repo.githubOwner || !repo.githubRepo) {
    throw Object.assign(new Error('Repository is missing githubOwner / githubRepo'), { status: 400, code: 'not-connected' });
  }
  return { repo, source: reader.githubSource({ owner: repo.githubOwner, repo: repo.githubRepo, token }) };
}

// Entities are keyed by natural id so re-running a sync updates rather than duplicates.
async function upsert(entity, filter, data) {
  const [existing] = await entity.filter(filter, undefined, 1);
  return existing ? entity.update(existing.id, data) : entity.create(data);
}

async function syncBranches(base44, { repositoryId }) {
  const { source } = await sourceFor(base44, repositoryId);
  const branches = await reader.listBranches(source, repositoryId);
  const E = base44.asServiceRole.entities;
  for (const b of branches) await upsert(E.Branch, { repositoryId, name: b.name }, b);
  return { branches };
}

async function syncCommits(base44, { repositoryId, branch, limit = 100 }) {
  const { source } = await sourceFor(base44, repositoryId);
  const commits = await reader.listCommits(source, repositoryId, branch, limit);
  const E = base44.asServiceRole.entities;
  for (const c of commits) await upsert(E.Commit, { repositoryId, sha: c.sha }, c);
  return { count: commits.length, commits };
}

// Diffs are deterministic for a (sha, path) pair, so they are computed once and cached.
async function getDiff(base44, { repositoryId, sha, path }) {
  const E = base44.asServiceRole.entities;
  const [cached] = await E.CommitDiff.filter({ repositoryId, sha, path }, undefined, 1);
  if (cached && cached.payloadVersion === reader.PAYLOAD_VERSION) return cached;
  const { source } = await sourceFor(base44, repositoryId);
  const diff = await reader.getDiff(source, { sha, path });
  const record = { repositoryId, ...diff };
  return cached ? E.CommitDiff.update(cached.id, record) : E.CommitDiff.create(record);
}

const ACTIONS = { syncBranches, syncCommits, getDiff };

Deno.serve(async (req) => {
  if (req.method !== 'POST') return fail(405, 'method-not-allowed', 'POST a JSON body');
  let body;
  try { body = await req.json(); } catch { return fail(400, 'bad-request', 'Body is not valid JSON'); }
  const handler = ACTIONS[body.action];
  if (!handler) return fail(400, 'bad-request', `Unknown action "${body.action}". Expected one of: ${Object.keys(ACTIONS).join(', ')}`);
  if (!body.repositoryId) return fail(400, 'bad-request', 'repositoryId is required');

  const base44 = createClientFromRequest(req);
  // Membership is checked server-side: the desktop client is never trusted to enforce permissions.
  const user = await base44.auth.me().catch(() => null);
  if (!user) return fail(401, 'unauthorized', 'Sign in first');
  const [member] = await base44.asServiceRole.entities.Member.filter({ repositoryId: body.repositoryId, userId: user.id }, undefined, 1);
  if (!member) return fail(403, 'forbidden', 'You do not have access to this repository');

  try {
    return json(await handler(base44, body));
  } catch (e) {
    return fail(e.status || 500, e.code || 'error', e.message);
  }
});
