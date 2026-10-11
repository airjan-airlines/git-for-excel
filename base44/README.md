# Connecting the desktop app to Base44

Base44 is **not** a git host — its GitHub integration syncs a Base44 app's own source code, not arbitrary
repositories. So the wiring is:

```
Desktop app  ──git push──▶  GitHub (private repo)  ◀──REST reads──  Base44 backend function
                                                                            │
                                                                            ▼
                                                                 Base44 entities ──▶ web app
```

GitHub holds the workbooks and the history (plain git). Base44 holds the *metadata and diffs* reviewers
read, and never has to understand `.xlsx` itself — the same engine the desktop app uses runs server-side.

---

## 1. Create the GitHub repository

A **private** repo, empty (no README — the first push should be the workbooks).

## 2. Point the desktop app at it

In the app: **Sync → Connect a remote**, paste `https://github.com/<owner>/<repo>.git`, then
**Sync → Publish**. Git handles sign-in; on GitHub use a **personal access token as the password**.
Equivalent from a terminal, since it is a normal git repo:

```bash
git remote add origin https://github.com/<owner>/<repo>.git
git push -u origin main
```

## 3. Give Base44 read access

Create a **fine-grained personal access token** (or a GitHub App installation token) scoped to that one
repository with **Contents: read**, then store it as a secret:

```bash
npx base44 secrets set GITHUB_TOKEN=github_pat_...
```

## 4. Deploy the function

```bash
node scripts/build-base44.js      # copies src/engine + src/sync/reader.js into the function
npx base44 functions deploy
```

`build-base44.js` must be re-run whenever the engine changes, so the web diff stays identical to the
desktop diff.

## 5. Entities the function expects

| Entity | Fields it reads / writes |
| --- | --- |
| `Repository` | `githubOwner`, `githubRepo` (you set these), plus your own name/visibility fields |
| `Branch` | `repositoryId`, `name`, `head`, `isDefault`, `updatedAt` |
| `Commit` | `repositoryId`, `sha`, `parentShas`, `branch`, `author`, `authorEmail`, `message`, `description`, `createdAt`, `isMerge`, `files[]` |
| `CommitDiff` | `repositoryId`, `sha`, `path`, `parentSha`, `format`, `status`, `summary`, `changeCount`, `changes[]`, `payloadVersion` |
| `Member` | `repositoryId`, `userId` — membership is checked on **every** call |

## 6. Calling it from the web app

```js
await base44.functions.invoke('officegit-sync', { action: 'syncBranches', repositoryId });
await base44.functions.invoke('officegit-sync', { action: 'syncCommits', repositoryId, branch: 'main' });
const diff = await base44.functions.invoke('officegit-sync', { action: 'getDiff', repositoryId, sha, path: 'Budget.xlsx' });
```

Call `syncCommits` from a GitHub webhook (push event) or on a schedule; `getDiff` is cached per
`(sha, path)` in `CommitDiff`, so it only computes once.

### Diff payload

`changes[]` is exactly what the desktop app renders:

```json
{
  "format": "xlsx",
  "summary": "1 cell modified",
  "changeCount": 1,
  "changes": [{
    "id": "cell:Sheet1!B3",
    "kind": "cell",
    "type": "modified",
    "detail": "input",
    "label": "Sheet1!B3",
    "group": "Sheet1",
    "before": { "value": "5", "formula": null, "type": "n", "style": "…" },
    "after":  { "value": "4242", "formula": null, "type": "n", "style": "…" }
  }]
}
```

`detail` tells a reviewer *what kind* of edit it was: `input` (a typed value), `formula`,
`calculated` (a formula result that moved because something upstream changed), or `format`.

---

## Tested vs. not

- **Tested** (`npm test`): everything up to and including the payloads — `src/sync/reader.js` against a real
  git repo of real workbooks, and the GitHub REST adapter against a stubbed API (URLs, auth headers,
  base64 blob decoding, error handling).
- **Not tested**: `entry.js` itself. I have no Base44 app to deploy to, so the Deno handler, the SDK calls
  and the entity shapes are written from the docs and unverified.
- **Known risk**: the engine is CommonJS and is loaded with `createRequire`. Deno resolves `npm:` specifiers
  natively but CommonJS `require('jszip')` inside a copied tree may not resolve. If the deploy fails on that,
  the fix is to vendor the two dependencies into `lib/` or convert the engine to ESM — tell me and I'll do it.
