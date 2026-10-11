'use strict';
// A repository is a plain git repository. Git does the version control (objects, branches, history, merge);
// this layer translates it into Office-file terms and plugs in the semantic diff + merge driver.
const fs = require('fs');
const fsp = fs.promises;
const os = require('os');
const path = require('path');
const engine = require('../engine');
const { Git } = require('./git');

const OFFICE = /\.(xlsx|docx|pptx)$/i;
const ATTR_LINE = '*.xlsx -text -diff merge=officegit';
const ATTRS = ['# Office files are zip archives: never touch line endings, and merge them with the OfficeGit engine.',
  ATTR_LINE, '*.docx -text -diff merge=officegit', '*.pptx -text -diff merge=officegit', ''].join('\n');
const posix = (p) => p.split(path.sep).join('/');
const STATUS = { A: 'added', M: 'modified', D: 'deleted', R: 'renamed', C: 'added', T: 'modified' };
const ident = (who) => ['-c', `user.name=${who}`, '-c', `user.email=${who.replace(/\s+/g, '.').toLowerCase()}@officegit.local`];

class Repo {
  constructor(root) {
    this.root = root;
    this.g = new Git(root);
    this.treeCache = new Map(); // blob sha -> semantic tree
    this.merge = null;
  }

  abs(rel) { return path.join(this.root, ...rel.split('/')); }
  lock(fn) {
    const run = (this._lock || Promise.resolve()).then(fn, fn);
    this._lock = run.catch(() => {});
    return run;
  }

  async open() {
    const top = await Git.topLevel(this.root);
    if (top && top.toLowerCase() === path.resolve(this.root).toLowerCase()) this.existing = true;
    else if (top) throw new Error(`That folder is inside the git repository at ${top}. Add that folder instead.`);
    else await this.g.init();
    await this.installMergeDriver();
    return this;
  }

  // Teach git how to merge Office files. Everything lives inside .git/ (local config + info/attributes), so the
  // user's working folder stays clean and their global git setup is untouched.
  async installMergeDriver() {
    const driver = path.join(__dirname, 'merge-driver.js');
    // Inside Electron, process.execPath is electron.exe, which only runs a plain script with this env var set.
    // git runs driver commands through a shell, so the prefix works on Windows and POSIX alike.
    const prefix = process.versions.electron ? 'ELECTRON_RUN_AS_NODE=1 ' : '';
    await this.g.git('config', 'merge.officegit.name', 'OfficeGit semantic merge');
    await this.g.git('config', 'merge.officegit.driver', `${prefix}"${process.execPath}" "${driver}" %O %A %B %P`);
    const gitDir = path.resolve(this.root, await this.g.git('rev-parse', '--git-dir'));
    const attrs = path.join(gitDir, 'info', 'attributes');
    await fsp.mkdir(path.dirname(attrs), { recursive: true });
    const cur = fs.existsSync(attrs) ? await fsp.readFile(attrs, 'utf8') : '';
    if (!cur.includes(ATTR_LINE)) await fsp.writeFile(attrs, cur ? cur.replace(/\s*$/, '\n') + ATTRS : ATTRS);
    // Office's "~$Book.xlsx" lock files appear while a file is open; they are never content.
    const exclude = path.join(gitDir, 'info', 'exclude');
    const ex = fs.existsSync(exclude) ? await fsp.readFile(exclude, 'utf8') : '';
    if (!ex.includes('~$*')) await fsp.writeFile(exclude, ex.replace(/\s*$/, '\n') + '~$*\n');
  }

  async repoId() {
    const r = await this.g.try('config', '--local', 'officegit.repoid');
    if (r.ok && r.out) return r.out;
    const id = require('crypto').randomUUID();
    await this.g.git('config', '--local', 'officegit.repoid', id);
    return id;
  }

  // ---- content ----
  async readAt(rev, rel) { return this.g.raw('show', `${rev}:${rel}`); }
  async treeAt(rev, rel) {
    const sha = await this.g.git('rev-parse', `${rev}:${rel}`);
    if (!this.treeCache.has(sha)) this.treeCache.set(sha, await engine.ingest(await this.readAt(rev, rel)));
    return this.treeCache.get(sha);
  }
  async treeOfWorking(rel) {
    const buf = await fsp.readFile(this.abs(rel));
    const sha = engine.sha(buf);
    if (!this.treeCache.has(sha)) this.treeCache.set(sha, await engine.ingest(buf));
    return this.treeCache.get(sha);
  }
  async safeTree(rev, rel) { try { return await this.treeAt(rev, rel); } catch { return null; } }

  async trackedFiles(rev = 'HEAD') {
    if (!(await this.g.hasCommits())) return [];
    return (await this.g.zList('ls-tree', '-r', '--name-only', '-z', rev)).filter((p) => OFFICE.test(p));
  }

  // ---- status ----
  async status() {
    await this.g.git('add', '-A', '--', '.');
    const base = await this.g.baseTree();
    const recs = await this.g.zList('diff', '--cached', '--name-status', '-M', '-z', base);
    const locks = [];
    for (const f of await this.g.zList('ls-files', '-o', '-c', '-z', '--exclude-standard')) {
      const b = path.basename(f);
      if (b.startsWith('~$')) locks.push(posix(path.join(path.dirname(f), b.slice(2))).toLowerCase());
    }
    const changes = [], resaved = [];
    for (let i = 0; i < recs.length; i++) {
      const letter = recs[i][0];
      let oldPath = null, p;
      if (letter === 'R' || letter === 'C') { oldPath = recs[++i]; p = recs[++i]; } else p = recs[++i];
      if (!OFFICE.test(p)) continue;
      const status = STATUS[letter] || 'modified';
      const c = { path: p, status, oldPath, open: locks.some((l) => p.toLowerCase().endsWith(l)) };
      try {
        c.size = status === 'deleted' ? 0 : fs.statSync(this.abs(p)).size;
        if (status === 'modified' || status === 'renamed') {
          const before = await this.treeAt(base, oldPath || p);
          const after = await this.treeOfWorking(p);
          // A re-save with no real edits changes the bytes but not the content: not a change.
          if (engine.treeHash(before) === engine.treeHash(after)) {
            if (status === 'modified') { resaved.push(p); continue; }
            c.detail = 'renamed only';
          }
        }
      } catch (e) { c.error = e.message; }
      changes.push(c);
    }
    changes.sort((a, b) => a.path.localeCompare(b.path));
    return { changes, resaved };
  }

  // ---- diff ----
  diffTrees(a, b) { return { changes: engine.diff(a, b), format: (b || a || {}).format }; }
  async diffWorking(p, oldPath) {
    const base = await this.g.baseTree();
    const before = await this.safeTree(base, oldPath || p);
    const after = fs.existsSync(this.abs(p)) ? await this.treeOfWorking(p) : null;
    return this.diffTrees(before, after);
  }
  async diffCommit(id, p) {
    const after = await this.safeTree(id, p);
    const parent = await this.g.try('rev-parse', `${id}^`);
    const before = parent.ok ? await this.safeTree(parent.out, p) : null;
    return this.diffTrees(before, after);
  }
  async compareWithWorking(id, p) {
    const before = await this.safeTree(id, p);
    const after = fs.existsSync(this.abs(p)) ? await this.treeOfWorking(p) : null;
    return this.diffTrees(before, after);
  }

  async suggestMessage(paths) {
    const { changes } = await this.status();
    const parts = [];
    for (const c of changes.filter((x) => paths.includes(x.path))) {
      const name = path.basename(c.path);
      if (c.status === 'added') parts.push(`Add ${name}`);
      else if (c.status === 'deleted') parts.push(`Remove ${name}`);
      else if (c.status === 'renamed') parts.push(`Rename ${path.basename(c.oldPath)} to ${name}`);
      else parts.push(`Update ${name}: ${engine.summarize((await this.diffWorking(c.path)).changes)}`);
    }
    return parts.join('; ');
  }

  // ---- commit ----
  commit(args) { return this.lock(() => this._commit(args)); }
  async _commit({ message, description, paths, author, force }) {
    if (!message || !message.trim()) throw new Error('Write a short summary of what changed.');
    const { changes } = await this.status();
    const sel = changes.filter((c) => paths.includes(c.path));
    if (!sel.length) throw new Error('Select at least one changed file.');
    const locked = sel.filter((c) => c.open);
    if (locked.length && !force) {
      const e = new Error(`These files are open in Office and may have unsaved edits: ${locked.map((c) => c.path).join(', ')}`);
      e.locked = locked.map((c) => c.path);
      throw e;
    }
    const bad = sel.find((c) => c.error);
    if (bad) throw new Error(`${bad.path} can't be read: ${bad.error}`);
    const files = [];
    for (const c of sel) { files.push(c.path); if (c.oldPath) files.push(c.oldPath); }
    const args = [...ident(author || 'OfficeGit user'), 'commit', '-m', message.trim()];
    if (description && description.trim()) args.push('-m', description.trim());
    await this.g.git(...args, '--', ...files);
    return { id: await this.g.head(), message: message.trim() };
  }

  // ---- history ----
  async log(branch, file, limit = 300) {
    if (!(await this.g.hasCommits())) return [];
    const S = '\u001e';
    const args = ['log', `--max-count=${limit}`, `--format=${S}%H${S}%P${S}%an${S}%ae${S}%at${S}%s${S}%b${S}`, '--name-status', '-M', '-z'];
    args.push(branch || 'HEAD');
    if (file) args.push('--follow', '--', file);
    const parts = (await this.g.git(...args)).split('\0');
    const commits = [];
    let cur = null;
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i].replace(/^\n+/, '');
      if (!p) continue;
      if (p.startsWith(S)) {
        const [, id, parents, author, email, at, subject, body] = p.split(S);
        cur = { id, parents: parents ? parents.split(' ') : [], author, email, time: +at * 1000, message: subject, description: (body || '').trim(), changed: [] };
        commits.push(cur);
        continue;
      }
      if (!cur || !/^[AMDRCT]\d*$/.test(p.split('\n')[0])) continue;
      const letter = p[0];
      if (letter === 'R' || letter === 'C') { const from = parts[++i], to = parts[++i]; if (OFFICE.test(to)) cur.changed.push({ path: to, status: 'renamed', oldPath: from }); }
      else { const f = parts[++i]; if (f && OFFICE.test(f)) cur.changed.push({ path: f, status: STATUS[letter] || 'modified' }); }
    }
    return commits;
  }

  restoreFile(id, rel) {
    return this.lock(async () => {
      await this.g.git('checkout', id, '--', rel);
      await this.g.try('reset', '--', rel);
    });
  }

  revertCommit(id, author) {
    return this.lock(async () => {
      await this.g.git(...ident(author || 'OfficeGit user'), 'revert', '--no-edit', id);
      return { commit: { id: await this.g.head() }, skipped: [] };
    });
  }

  // Discarding goes through git stash, so nothing is ever truly lost.
  discard(paths) {
    return this.lock(() => this.g.git('stash', 'push', '--include-untracked', '-m', `Discarded ${new Date().toLocaleString()}`, '--', ...paths));
  }

  async openVersion(id, rel) {
    const dir = path.join(os.tmpdir(), 'OfficeGit', path.basename(this.root), String(id).slice(0, 8));
    await fsp.mkdir(dir, { recursive: true });
    const dest = path.join(dir, path.basename(rel));
    await fsp.rm(dest, { force: true });
    await fsp.writeFile(dest, await this.readAt(id, rel));
    fs.chmodSync(dest, 0o444);
    return dest;
  }

  // ---- branches ----
  async branches() {
    const current = await this.g.branch();
    const r = await this.g.try('for-each-ref', '--format=%(refname:short)\t%(objectname)\t%(committerdate:unix)', 'refs/heads');
    const list = (r.ok && r.out ? r.out.split('\n').filter(Boolean) : []).map((l) => {
      const [name, tip, time] = l.split('\t');
      return { name, tip, time: +time * 1000, current: name === current };
    });
    if (!list.some((b) => b.current)) list.push({ name: current, tip: null, time: Date.now(), current: true });
    return { current, list: list.sort((a, b) => a.name.localeCompare(b.name)), deleted: [] };
  }

  // Scenarios = your branches plus the ones teammates shared (base44/*), with who last changed each and when.
  async scenarios() {
    const current = await this.g.branch();
    const r = await this.g.try('for-each-ref', '--format=%(refname)%09%(authorname)%09%(committerdate:unix)%09%(subject)', 'refs/heads', 'refs/remotes/base44');
    const byName = new Map();
    for (const line of r.ok && r.out ? r.out.split('\n').filter(Boolean) : []) {
      const [ref, author, at, subject] = line.split('\t');
      const name = ref.replace(/^refs\/(heads|remotes\/base44)\//, '');
      if (name === 'HEAD') continue;
      const e = byName.get(name) || { name, local: false, shared: false, time: 0 };
      if (ref.startsWith('refs/heads/')) e.local = true; else e.shared = true;
      if (+at * 1000 > e.time) Object.assign(e, { author, time: +at * 1000, subject });
      byName.set(name, e);
    }
    return { current, list: [...byName.values()].sort((a, b) => b.time - a.time) };
  }

  checkName(name) {
    if (!name || /[\s~^:?*[\\]/.test(name) || name.includes('..')) throw new Error('Branch names can’t contain spaces or special characters.');
  }
  createAndSwitch(name) {
    this.checkName(name);
    return this.lock(() => this.g.git('checkout', '-b', name));
  }
  renameBranch(oldName, name) { this.checkName(name); return this.g.git('branch', '-m', oldName, name); }
  async unmergedCommits(name) {
    // --exclude only applies to the --branches that follows it, so the order here matters
    const r = await this.g.try('log', '--format=%H', name, '--not', `--exclude=${name}`, '--branches');
    return r.ok && r.out ? r.out.split('\n').filter(Boolean).length : 0;
  }
  async deleteBranch(name, force) {
    const r = await this.g.try('branch', force ? '-D' : '-d', name);
    if (!r.ok) {
      if (/not fully merged/i.test(r.error)) { const e = new Error('That branch has versions that exist nowhere else.'); e.unmerged = true; throw e; }
      throw new Error(r.error);
    }
  }

  switchBranch(name, opts) { return this.lock(() => this._switchBranch(name, opts)); }
  async _switchBranch(name, { stash, discard } = {}) {
    const from = await this.g.branch();
    if (name === from) return { switched: false };
    const { changes } = await this.status();
    if (changes.length) {
      if (stash || discard) await this.g.git('stash', 'push', '--include-untracked', '-m', `${discard ? 'Discarded' : 'Work'} on ${from}`);
      else return { needsAction: true, changes };
    }
    await this.g.git('checkout', name);
    return { switched: true };
  }

  // ---- stash ----
  async stashSave(message) {
    const { changes } = await this.status();
    if (!changes.length) throw new Error('There are no changes to stash.');
    await this.g.git('stash', 'push', '--include-untracked', '-m', message || `Stashed ${new Date().toLocaleString()}`);
  }
  async stashList() {
    const r = await this.g.try('stash', 'list', '--format=%gd\t%gs\t%at');
    if (!r.ok || !r.out) return [];
    return r.out.split('\n').filter(Boolean).map((l) => {
      const [id, subject, at] = l.split('\t');
      return { id, message: subject.replace(/^(On|WIP on) [^:]*:\s*/, ''), branch: (/^(?:On|WIP on) ([^:]*)/.exec(subject) || [, ''])[1], at: +at * 1000 };
    });
  }
  stashApply(id) { return this.lock(() => this.g.git('stash', 'pop', id)); }
  stashDrop(id) { return this.g.git('stash', 'drop', id); }

  // ---- merge: git runs it, our driver handles the Office files ----
  mergeStart(branch) { return this.lock(() => this._mergeStart(branch)); }
  async _mergeStart(branch) {
    const into = await this.g.branch();
    if (branch === into) throw new Error('Choose a different branch to merge.');
    if ((await this.status()).changes.length) throw new Error('Commit or stash your changes before merging.');
    if ((await this.g.try('merge-base', '--is-ancestor', branch, 'HEAD')).ok) return { status: 'uptodate' };
    const ff = (await this.g.try('merge-base', '--is-ancestor', 'HEAD', branch)).ok;
    const r = await this.g.try('merge', '--no-commit', ...(ff ? [] : ['--no-ff']), branch);
    const conflicted = await this.g.zList('diff', '--name-only', '--diff-filter=U', '-z').catch(() => []);
    if (!r.ok && !conflicted.length) { await this.g.try('merge', '--abort'); throw new Error(r.error); }
    if (ff && !conflicted.length) { this.merge = null; return { status: 'fastforward', branch, into }; }
    const staged = await this.g.zList('diff', '--cached', '--name-only', '-z', 'HEAD').catch(() => []);
    const files = staged.filter((p) => !conflicted.includes(p)).map((p) => ({ path: p, state: 'clean' }));
    const stillConflicted = [];
    for (const p of conflicted) {
      const d = await this.conflictDetail(p);
      // Safety net: if git flagged a conflict but the engine can merge the file (e.g. the driver did not
      // run because this clone was never configured), merge it here instead of bothering the user.
      if (d.merged) {
        await fsp.writeFile(this.abs(p), d.merged);
        await this.g.git('add', '--', p);
        files.push({ path: p, state: 'clean' });
        continue;
      }
      stillConflicted.push(p);
      files.push({ path: p, state: 'conflict', summary: d.summary, items: d.items });
    }
    this.merge = { branch, into };
    return { status: stillConflicted.length ? 'conflicts' : 'ready', branch, into, files };
  }

  // What actually clashes, in Office terms, for the conflict screen.
  async conflictDetail(p) {
    try {
      const [base, ours, theirs] = await Promise.all([this.g.raw('show', `:1:${p}`), this.g.raw('show', `:2:${p}`), this.g.raw('show', `:3:${p}`)]);
      const r = await engine.mergeFiles(base, ours, theirs);
      if (r.buffer) return { merged: r.buffer, summary: 'Merged automatically', items: [] };
      if (r.unsupported) return { summary: r.unsupported, items: [] };
      return {
        summary: `${r.conflicts.length} overlapping edit${r.conflicts.length === 1 ? '' : 's'}`,
        items: r.conflicts.slice(0, 25).map((c) => ({ label: c.label, ours: brief(c.ours), theirs: brief(c.theirs) })),
      };
    } catch (e) { return { summary: e.message, items: [] }; }
  }

  mergeComplete(resolutions = {}, author) { return this.lock(() => this._mergeComplete(resolutions, author)); }
  async _mergeComplete(resolutions, author) {
    if (!this.merge) throw new Error('No merge in progress.');
    const conflicted = await this.g.zList('diff', '--name-only', '--diff-filter=U', '-z').catch(() => []);
    const unresolved = conflicted.filter((p) => !resolutions[p]);
    if (unresolved.length) throw new Error(`Choose a version for: ${unresolved.join(', ')}`);
    for (const p of conflicted) {
      await this.g.git('checkout', resolutions[p] === 'theirs' ? '--theirs' : '--ours', '--', p);
      await this.g.git('add', '--', p);
    }
    await this.g.git(...ident(author || 'OfficeGit user'), 'commit', '--no-edit');
    const m = this.merge;
    this.merge = null;
    return { id: await this.g.head(), message: `Merge ${m.branch} into ${m.into}` };
  }
  async mergeAbort() { this.merge = null; await this.g.try('merge', '--abort'); }

  // ---- remotes (any git host: GitHub, GitLab, Gitea, a network share...) ----
  async remotes() {
    const r = await this.g.try('remote', '-v');
    if (!r.ok || !r.out) return [];
    const seen = new Map();
    for (const line of r.out.split('\n').filter(Boolean)) {
      const [name, rest] = line.split('\t');
      if (!seen.has(name)) seen.set(name, { name, url: (rest || '').replace(/\s+\(\w+\)$/, '') });
    }
    return [...seen.values()];
  }
  async setRemote(name, url) {
    const exists = (await this.remotes()).some((r) => r.name === name);
    await this.g.git('remote', exists ? 'set-url' : 'add', name, url);
  }
  async removeRemote(name) { await this.g.git('remote', 'remove', name); }

  // How far the current branch is ahead of / behind its upstream. null when there is no remote yet.
  async syncState() {
    const remotes = await this.remotes();
    const branch = await this.g.branch();
    if (!remotes.length) return { remotes, branch, upstream: null };
    const up = await this.g.try('rev-parse', '--abbrev-ref', '--symbolic-full-name', '@{upstream}');
    if (!up.ok) return { remotes, branch, upstream: null, publishable: true };
    const counts = await this.g.try('rev-list', '--left-right', '--count', `${up.out}...HEAD`);
    const [behind, ahead] = counts.ok ? counts.out.split(/\s+/).map(Number) : [0, 0];
    return { remotes, branch, upstream: up.out, ahead, behind };
  }

  fetch(remote = 'origin') { return this.lock(() => this.g.git('fetch', '--prune', remote)); }

  // Publish the current branch and set it to track the remote.
  push(remote = 'origin', { setUpstream } = {}) {
    return this.lock(async () => {
      const branch = await this.g.branch();
      const args = ['push'];
      if (setUpstream) args.push('--set-upstream');
      args.push(remote, branch);
      const r = await this.g.try(...args);
      if (!r.ok) throw new Error(friendlyRemoteError(r.error));
      return { branch, remote };
    });
  }

  // Pull is fetch + merge, so the Office merge driver handles any overlapping edits exactly as a local merge does.
  pull(remote = 'origin') {
    return this.lock(async () => {
      if ((await this.status()).changes.length) throw new Error('Commit or stash your changes before pulling.');
      const branch = await this.g.branch();
      const r = await this.g.try('pull', '--no-rebase', '--no-edit', remote, branch);
      // git names the merge after the remote's URL; give it a readable message instead
      if (r.ok && (await this.g.try('rev-parse', '--verify', '-q', 'HEAD^2')).ok && /^Merge branch '.*' of /.test(await this.g.git('log', '-1', '--format=%s'))) {
        await this.g.git('commit', '--amend', '--no-edit', '-m', `Merge ${remote}/${branch} into ${branch}`);
      }
      if (r.ok) return { conflicts: [] };
      const conflicted = await this.g.zList('diff', '--name-only', '--diff-filter=U', '-z').catch(() => []);
      if (!conflicted.length) throw new Error(friendlyRemoteError(r.error));
      // Same conflict screen as a local merge.
      const files = [];
      for (const p of conflicted) {
        const d = await this.conflictDetail(p);
        if (d.merged) { await fsp.writeFile(this.abs(p), d.merged); await this.g.git('add', '--', p); continue; }
        files.push({ path: p, state: 'conflict', summary: d.summary, items: d.items });
      }
      this.merge = { branch: `${remote}/${await this.g.branch()}`, into: await this.g.branch() };
      if (!files.length) return { conflicts: [], autoMerged: conflicted };
      return { conflicts: files, into: this.merge.into, branch: this.merge.branch };
    });
  }

  async fsck() {
    const r = await this.g.try('fsck', '--no-progress', '--connectivity-only');
    return { ok: r.ok, problems: r.ok ? [] : [r.error] };
  }
}

// git's remote errors are aimed at developers; say what the user can actually do about it.
function friendlyRemoteError(msg) {
  if (/could not read Username|Authentication failed|Invalid username or password|terminal prompts disabled/i.test(msg)) {
    return 'The remote rejected your sign-in. Check your account access, or use a personal access token as the password.';
  }
  if (/Permission.*denied|403/i.test(msg)) return 'Your account does not have permission to write to that remote.';
  if (/Could not resolve host|unable to access|timed out/i.test(msg)) return 'Could not reach the remote. Check the address and your connection.';
  if (/non-fast-forward|fetch first|rejected/i.test(msg)) return 'The remote has versions you don’t have yet. Pull first, then push again.';
  if (/does not appear to be a git repository/i.test(msg)) return 'That address is not a git repository.';
  return msg.split('\n').filter(Boolean).slice(-2).join(' ');
}

function brief(d) {
  if (d == null) return '(removed)';
  if (Array.isArray(d)) return d.map((x) => x.text ?? x.title ?? '…').join(' / ');
  if (d.formula != null) return '=' + d.formula;
  if (d.text != null) return d.text;
  if (d.value != null) return String(d.value);
  return JSON.stringify(d);
}

module.exports = { Repo, OFFICE };
