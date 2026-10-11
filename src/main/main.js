'use strict';
const { app, BrowserWindow, dialog, ipcMain, shell, Menu, nativeTheme } = require('electron');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Repo } = require('./repo');
const hosted = require('../sync/hosted');
const { Patterns } = require('./patterns');
const { fromEnv } = require('../sync/base44');

let win, watcher, watchTimer;
const repos = new Map(); // path -> Repo
let config = { repos: [], active: null, author: os.userInfo().username };
const cfgPath = () => path.join(app.getPath('userData'), 'config.json');
const saveConfig = () => fs.writeFileSync(cfgPath(), JSON.stringify(config));
const loadConfig = () => {
  try { config = { ...config, ...JSON.parse(fs.readFileSync(cfgPath(), 'utf8')) }; } catch { /* first run */ }
  // folders can be moved or deleted between sessions; forget those rather than failing on startup
  config.repos = (config.repos || []).filter((p) => fs.existsSync(p));
  if (!config.repos.includes(config.active)) config.active = config.repos[0] || null;
};

async function getRepo(p = config.active) {
  if (!p) throw new Error('Add a repository first.');
  if (!repos.has(p)) repos.set(p, await new Repo(p).open());
  return repos.get(p);
}
let b44;
const base44 = () => (b44 ||= fromEnv()); // throws a readable error if base44.local.json / env vars are missing
const isHosted = (remote) => remote === hosted.REMOTE;

const repoList = () => config.repos.filter((p) => fs.existsSync(p)).map((p) => ({ path: p, name: path.basename(p), active: p === config.active }));

function watch() {
  if (watcher) watcher.close();
  watcher = null;
  if (!config.active || !fs.existsSync(config.active)) return;
  try {
    watcher = fs.watch(config.active, { recursive: true }, (_e, f) => {
      if (f && (f.startsWith('.git') || f.includes('~$'))) return;
      clearTimeout(watchTimer);
      watchTimer = setTimeout(() => win && !win.isDestroyed() && win.webContents.send('fs-changed'), 400);
    });
  } catch { /* watching is best-effort */ }
}

async function addRepo(p) {
  if (!config.repos.includes(p)) config.repos.push(p);
  config.active = p;
  saveConfig();
  await getRepo(p);
  watch();
}

const api = {
  async 'repos:list'() { return { repos: repoList(), author: config.author }; },
  async 'repos:add'() {
    const r = await dialog.showOpenDialog(win, { title: 'Add a folder of Office files', properties: ['openDirectory'] });
    if (r.canceled) return null;
    await addRepo(r.filePaths[0]);
    return repoList();
  },
  async 'repos:select'(p) { config.active = p; saveConfig(); await getRepo(p); watch(); return repoList(); },
  async 'repos:remove'(p) {
    config.repos = config.repos.filter((x) => x !== p);
    repos.delete(p);
    if (config.active === p) config.active = config.repos[0] || null;
    saveConfig(); watch();
    return repoList();
  },
  async 'app:quit'() { app.quit(); },
  async 'app:devtools'() { win.webContents.toggleDevTools(); },
  async 'repos:reveal'() { shell.showItemInFolder(config.active); },
  async 'settings:author'(name) { config.author = name.trim() || config.author; saveConfig(); return config.author; },

  async 'status'() { return (await getRepo()).status(); },
  async 'diff:working'(p, oldPath) { return (await getRepo()).diffWorking(p, oldPath); },
  async 'diff:commit'(id, p) { return (await getRepo()).diffCommit(id, p); },
  async 'diff:compare'(id, p) { return (await getRepo()).compareWithWorking(id, p); },
  async 'patterns'(paths) {
    const repo = await getRepo();
    if (!repo.patterns) repo.patterns = new Patterns(repo);
    return repo.patterns.suggest(paths);
  },
  async 'suggest'(paths) { return (await getRepo()).suggestMessage(paths); },
  async 'commit'(args) { return (await getRepo()).commit({ ...args, author: config.author }); },
  async 'discard'(paths) { return (await getRepo()).discard(paths); },
  async 'log'(branch, file) { return (await getRepo()).log(branch, file); },
  async 'tracked'() { return (await getRepo()).trackedFiles(); },
  async 'restore'(id, p) { return (await getRepo()).restoreFile(id, p); },
  async 'revert'(id) { return (await getRepo()).revertCommit(id, config.author); },
  async 'open:version'(id, p) {
    const f = await (await getRepo()).openVersion(id, p);
    const err = await shell.openPath(f);
    if (err) throw new Error(`Couldn’t open ${path.basename(f)}: ${err}`);
  },
  async 'open:working'(p) {
    const err = await shell.openPath(path.join(config.active, ...p.split('/')));
    if (err) throw new Error(err);
  },

  async 'branches'() { return (await getRepo()).branches(); },
  async 'branch:create'(name) { return (await getRepo()).createAndSwitch(name); },
  async 'branch:switch'(name, opts) { return (await getRepo()).switchBranch(name, opts); },
  async 'branch:rename'(a, b) { return (await getRepo()).renameBranch(a, b); },
  async 'branch:unmerged'(name) { return (await getRepo()).unmergedCommits(name); },
  async 'branch:delete'(name, force) { return (await getRepo()).deleteBranch(name, force); },

  async 'stash:save'(msg) { return (await getRepo()).stashSave(msg); },
  async 'stash:list'() { return (await getRepo()).stashList(); },
  async 'stash:apply'(id) { return (await getRepo()).stashApply(id); },
  async 'stash:drop'(id) { return (await getRepo()).stashDrop(id); },

  async 'sync:state'() {
    const repo = await getRepo();
    const s = await repo.syncState();
    s.hosted = await hosted.linked(repo);
    // When linked to Base44, that is the remote every Sync action uses.
    if (s.hosted) s.remotes.sort((a, b) => (isHosted(b.name) ? 1 : 0) - (isHosted(a.name) ? 1 : 0));
    return s;
  },
  async 'sync:setRemote'(name, url) { return (await getRepo()).setRemote(name, url); },
  async 'sync:fetch'(r) {
    const repo = await getRepo();
    return isHosted(r) ? repo.lock(() => hosted.fetch(repo, base44())) : repo.fetch(r);
  },
  async 'sync:push'(r, opts) {
    const repo = await getRepo();
    return isHosted(r) ? repo.lock(() => hosted.push(repo, base44(), { owner: config.author })) : repo.push(r, opts);
  },
  async 'sync:pull'(r) {
    const repo = await getRepo();
    if (isHosted(r)) await repo.lock(() => hosted.fetch(repo, base44()));
    return repo.pull(r);
  },
  async 'hosted:list'() { return hosted.listRemote(base44()); },
  async 'hosted:link'(opts) {
    const repo = await getRepo();
    return repo.lock(() => hosted.link(repo, base44(), { ...opts, owner: config.author }));
  },

  async 'merge:start'(branch) { return (await getRepo()).mergeStart(branch); },
  async 'merge:complete'(res) { return (await getRepo()).mergeComplete(res, config.author); },
  async 'merge:abort'() { return (await getRepo()).mergeAbort(); },
};

ipcMain.handle('api', async (_e, name, ...args) => {
  try {
    if (!api[name]) throw new Error('Unknown action ' + name);
    return { ok: true, value: await api[name](...args) };
  } catch (e) {
    return { ok: false, error: e.message, extra: { locked: e.locked, unmerged: e.unmerged } };
  }
});

app.whenReady().then(async () => {
  loadConfig();
  Menu.setApplicationMenu(null);
  win = new BrowserWindow({
    width: 1280, height: 820, minWidth: 900, minHeight: 600, title: 'GitXL', backgroundColor: '#ffffff',
    titleBarStyle: 'hidden', titleBarOverlay: { color: '#ffffff', symbolColor: '#111827', height: 64 },
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false, sandbox: true },
  });
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
  win.webContents.on('will-navigate', (e) => e.preventDefault());
  if (process.env.OG_REPO) await addRepo(process.env.OG_REPO);
  await win.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));
  watch();
  if (process.env.OG_SHOT) { // dev utility: OG_SHOT=out.png captures the window after a short delay
    setTimeout(async () => { if (process.env.OG_JS) { console.log('JS:', await win.webContents.executeJavaScript(process.env.OG_JS)); await new Promise((r) => setTimeout(r, 1200)); } fs.writeFileSync(process.env.OG_SHOT, (await win.webContents.capturePage()).toPNG()); app.quit(); }, +process.env.OG_DELAY || 2500);
  }
});
app.on('window-all-closed', () => app.quit());
