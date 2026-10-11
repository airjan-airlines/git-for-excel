'use strict';
// skip null/false children instead of rendering them as text
const _rc = Element.prototype.replaceChildren;
Element.prototype.replaceChildren = function (...k) { return _rc.apply(this, k.flat(Infinity).filter((x) => x != null && x !== false)); };
const call = (...a) => window.og.call(...a);
const $ = (id) => document.getElementById(id);
const h = (tag, props = {}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(props || {})) {
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (v === true) e.setAttribute(k, '');
    else if (v !== false && v != null) e.setAttribute(k, v);
  }
  for (const k of kids.flat()) if (k != null && k !== false) e.append(k.nodeType ? k : document.createTextNode(k));
  return e;
};

const state = {
  repos: [], author: '', tab: 'changes', changes: [], resaved: [], selected: new Set(), current: null,
  commits: [], commit: null, commitFile: null, branches: null, filter: '',
};

const ICON = { added: '＋', removed: '−', deleted: '−', modified: '～', moved: '↔', renamed: '→' };
const LABEL = { added: 'Added', removed: 'Removed', deleted: 'Deleted', modified: 'Modified', moved: 'Moved', renamed: 'Renamed' };

function toast(msg, error) {
  const t = $('toast');
  t.textContent = msg;
  t.className = 'show' + (error ? ' error' : '');
  clearTimeout(toast.t);
  toast.t = setTimeout(() => (t.className = ''), error ? 5000 : 2500);
}
const guard = (fn) => async (...a) => { try { return await fn(...a); } catch (e) { toast(e.message, true); } };
const ago = (t) => {
  const s = (Date.now() - t) / 1000;
  if (s < 60) return 'just now';
  if (s < 3600) return `${Math.floor(s / 60)} min ago`;
  if (s < 86400) return `${Math.floor(s / 3600)} h ago`;
  return new Date(t).toLocaleDateString();
};

// ---------- popovers & modals ----------
function closeMenu() { $('menu').hidden = true; }
function openMenu(anchor, content) {
  const m = $('menu');
  m.replaceChildren(...content);
  m.hidden = false;
  const r = anchor.getBoundingClientRect ? anchor.getBoundingClientRect() : { left: anchor.x, bottom: anchor.y };
  m.style.left = Math.min(r.left, window.innerWidth - 290) + 'px';
  m.style.top = Math.min(r.bottom, window.innerHeight - 100) + 'px';
}
document.addEventListener('mousedown', (e) => { if (!$('menu').contains(e.target)) closeMenu(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { closeMenu(); closeModal(); } });
const item = (label, fn, cls = '') => h('button', { class: 'item ' + cls, onclick: () => { closeMenu(); fn(); } }, label);

function closeModal() { $('modalRoot').replaceChildren(); }
function modal({ title, body, actions = [], wide }) {
  const foot = h('footer', {}, ...actions.map((a) => h('button', {
    class: a.kind || 'secondary', onclick: async () => { try { if ((await a.run?.()) !== false) closeModal(); } catch (e) { toast(e.message, true); } },
  }, a.label)));
  const m = h('div', { class: 'modal' + (wide ? ' wide' : ''), role: 'dialog', 'aria-modal': 'true', 'aria-label': title },
    h('header', {}, title), h('div', { class: 'body' }, body), foot);
  $('modalRoot').replaceChildren(h('div', { class: 'overlay' }, m));
  m.querySelector('input,textarea,button')?.focus();
}
const confirmBox = (title, body, label, run, kind = 'danger') =>
  modal({ title, body, actions: [{ label: 'Cancel' }, { label, kind, run }] });

// ---------- data refresh ----------
async function loadRepos() {
  const r = await call('repos:list');
  state.repos = r.repos; state.author = r.author;
  const active = r.repos.find((x) => x.active);
  $('repoName').textContent = active ? active.name : 'Open a project';
  const initials = (t) => (t || '').split(/[\s_\-]+|(?=[A-Z][a-z])/).filter(Boolean).map((w) => w[0]).join('').slice(0, 2).toUpperCase() || '—';
  $('projTile').textContent = active ? initials(active.name) : '＋';
  $('projMeta').textContent = `Excel · Owner: ${state.author || 'you'}`;
  $('meAvatar').textContent = initials(state.author || '?').slice(0, 1);
  return active;
}

async function refresh() {
  const active = await loadRepos();
  if (!active) { renderEmpty(); renderList(); $('branchName').textContent = '—'; return; }
  const [st, br] = await Promise.all([call('status'), call('branches')]);
  state.changes = st.changes; state.resaved = st.resaved; state.branches = br;
  const paths = new Set(st.changes.map((c) => c.path));
  for (const p of [...state.selected]) if (!paths.has(p)) state.selected.delete(p);
  if (!state.initSel) { st.changes.forEach((c) => state.selected.add(c.path)); state.initSel = true; }
  $('branchName').textContent = br.current;
  renderScenario(br);
  renderSync().catch(() => {});
  $('changeCount').textContent = st.changes.length;
  renderList();
  if (state.tab === 'changes') loadSuggestions();
  if (state.tab === 'history') await loadHistory();
  else if (state.current && !paths.has(state.current)) { state.current = null; showWorkingDiff(); }
  else if (state.current) showWorkingDiff(true);
  else if (st.changes.length) selectFile(st.changes[0].path);
  else showWorkingDiff();
}

// While on a scenario, say so in plain words and offer the two things you can do with it.
function renderScenario(br) {
  const b = $('scenarioBanner');
  const main = br.list.some((x) => x.name === 'main') ? 'main' : null;
  if (!main || br.current === main) { b.hidden = true; return; }
  const name = br.current.replace(/-/g, ' ');
  b.replaceChildren(h('span', { class: 'dot' }), h('span', { class: 'grow' }, 'You’re working on the scenario ', h('b', {}, name), '. The main model is unchanged until you bring it in.'),
    h('button', { class: 'secondary', onclick: guard(() => switchTo(main)) }, 'Back to main model'),
    h('button', { class: 'primary', onclick: guard(() => bringIn(br.current, main)) }, 'Bring into main model'));
  b.hidden = false;
}
async function bringIn(scenario, main) {
  if (state.changes.length) throw new Error('Save your changes in this scenario first.');
  const r = await call('branch:switch', main, {});
  if (r.needsAction) throw new Error('Save your changes in this scenario first.');
  await refresh();
  await startMerge(scenario);
}
$('whatIfBtn').addEventListener('click', () => (hasRepo() ? newBranchDialog() : toast('Open a project first', true)));
$('moreBtn').addEventListener('click', () => {
  const scen = (state.branches && state.branches.list.filter((b) => b.name !== 'main')) || [];
  openMenu($('moreBtn'), [
    ...scen.map((b) => item(`Open scenario “${b.name.replace(/-/g, ' ')}”`, () => switchTo(b.name))),
    state.branches && state.branches.current !== 'main' && item('Back to main model', () => switchTo('main')),
    scen.length && h('hr'),
    item('Open a project folder…', guard(addRepo)),
    item('Share with your team…', guard(connectBase44Dialog)),
    item('Check for team updates', guard(async () => { state.checked = 0; await refresh(); })),
    h('hr'),
    item('Rename this scenario…', withBranches(renameDialog)),
    item('Delete a scenario…', withBranches(deleteChooser)),
    item('Stashed & discarded work…', needRepo(() => openTrash())),
    item('Show in Explorer', needRepo(() => call('repos:reveal'))),
    item('Settings…', () => openSettings()),
    item('How GitXL works', helpDialog),
  ].filter(Boolean));
});

async function addRepo() {
  if (!(await call('repos:add'))) return;
  state.initSel = false; state.selected.clear(); state.current = null; state.checked = 0;
  await refresh();
  // A brand-new folder with no history and no remote: offer to bring a team repository down from Base44.
  const s = await call('sync:state');
  if (!s.remotes.length && !state.branches.list.some((b) => b.tip)) connectBase44Dialog().catch((e) => toast(e.message, true));
}

// ---------- sync with a remote ----------
async function renderSync() {
  const btn = $('syncBtn');
  const label = btn.querySelector('.tb-value');
  const sub = btn.querySelector('.tb-label');
  btn.disabled = false;
  try {
    state.sync = await call('sync:state');
  } catch { btn.disabled = true; return; }
  const s = state.sync;
  btn.className = 'tb';
  if (!s.remotes.length) { sub.textContent = 'Team'; label.textContent = 'Not shared yet · Share with team'; btn.classList.add('todo'); btn.title = 'Share this project with your team'; return; }
  // Base44 can't notify us, so look for teammates' pushes in the background (at most once a minute).
  if (s.hosted && !(Date.now() - (state.checked || 0) < 60000)) {
    state.checked = Date.now();
    call('sync:fetch', s.remotes[0].name).then(() => renderSync()).catch(() => {});
  }
  if (!s.upstream) { sub.textContent = 'Team'; label.textContent = s.branch === 'main' ? 'Not shared yet · Share now' : 'Private scenario'; btn.classList.add(s.branch === 'main' ? 'todo' : 'busy'); btn.title = 'Click to share'; return; }
  // Teammates saved something and you have nothing unsaved: bring it in without asking.
  if (s.behind && !state.changes.length && !state.pulling) {
    state.pulling = true;
    call('sync:pull', s.remotes[0].name).then((r) => {
      if (r.conflicts && r.conflicts.length) return showConflicts({ branch: r.branch, into: r.into, files: r.conflicts }, 'Combine');
      toast(`Updated with ${s.behind} new version${s.behind === 1 ? '' : 's'} from your team`);
      return refresh();
    }).catch((e) => toast(e.message, true)).finally(() => { state.pulling = false; });
    sub.textContent = 'Team'; label.textContent = 'Getting your team’s latest…'; btn.classList.add('busy');
    return;
  }
  const bits = [];
  if (s.ahead) bits.push(`↑ ${s.ahead}`);
  if (s.behind) bits.push(`↓ ${s.behind}`);
  sub.textContent = 'Team';
  label.textContent = s.behind ? `${s.behind} update${s.behind === 1 ? '' : 's'} from your team · Get` : s.ahead ? `${s.ahead} version${s.ahead === 1 ? '' : 's'} not shared · Share` : '✓ Shared with team';
  btn.classList.add(s.behind || s.ahead ? 'todo' : 'ok');
  btn.title = bits.length ? `${s.ahead} version(s) to share, ${s.behind} to get` : 'Your team has everything you have';
}

$('syncBtn').addEventListener('click', guard(async () => {
  if (!hasRepo()) return;
  const s = state.sync || (await call('sync:state'));
  if (!s.remotes.length) return connectBase44Dialog();
  if (s.behind) return doPull();
  if (!s.upstream || s.ahead) { await shareNow(); return; }
  const items = [];
  if (s.hosted) items.push(h('div', { class: 'note pad' }, `Shared project · ${s.hosted.name || ''}`));
  if (!s.upstream) items.push(item(`Share “${s.branch}” with the team`, guard(async () => {
    await call('sync:push', s.remotes[0].name, { setUpstream: true });
    toast('Shared with your team'); await refresh();
  })));
  else {
    items.push(item(`Share ${s.ahead || 0} new version(s)`, guard(async () => { await call('sync:push', s.remotes[0].name); toast('Shared with your team'); await refresh(); })));
    items.push(item(`Get latest (${s.behind || 0} new version${s.behind === 1 ? '' : 's'})`, guard(doPull)));
  }
  items.push(item('Check for changes', guard(async () => { await call('sync:fetch', s.remotes[0].name); toast('Checked'); await refresh(); })));
  if (!s.hosted) items.push(h('hr'), item('Change remote address…', connectRemoteDialog));
  if (!s.hosted) items.push(item('Share with your team…', connectBase44Dialog));
  openMenu($('syncBtn'), items);
}));

async function doPull() {
  const s = state.sync || (await call('sync:state'));
  const r = await call('sync:pull', s.remotes[0].name);
  if (!r.conflicts || !r.conflicts.length) { toast('Up to date'); return refresh(); }
  showConflicts({ branch: r.branch, into: r.into, files: r.conflicts }, 'Combine');
}

// Base44 hosts the repository: either start a new one from this folder, or bring an existing one down (a clone).
async function connectBase44Dialog() {
  const remote = await call('hosted:list');
  const NEW = '__new__';
  const pick = h('select', { 'aria-label': 'Project' },
    h('option', { value: NEW }, '＋ New project'),
    ...remote.map((r) => h('option', { value: r.id }, `${r.name}${r.owner ? ` — ${r.owner}` : ''}${r.hasData ? '' : ' (empty)'}`)));
  const name = h('input', { type: 'text', value: state.repos.find((r) => r.active)?.name || '', placeholder: 'Project name', 'aria-label': 'Project name' });
  const hint = h('p', { class: 'note' });
  const sync = () => {
    const isNew = pick.value === NEW;
    name.hidden = !isNew;
    hint.textContent = isNew
      ? 'Creates the project for your team. Then use Share to send your versions.'
      : 'Links this folder to the project and downloads its full history. In an empty folder, the workbooks appear right away.';
  };
  pick.addEventListener('change', sync); sync();
  modal({
    title: 'Share with your team',
    body: [h('p', {}, 'Your team’s projects live in GitXL. Share sends your versions there; Get latest brings everyone else’s down and combines them cell by cell.'), pick, name, hint,
      h('p', { class: 'note' }, h('a', { href: '#', onclick: (e) => { e.preventDefault(); closeModal(); connectRemoteDialog(); } }, 'Use a git address (GitHub, GitLab…) instead'))],
    actions: [{ label: 'Cancel' }, { label: 'Connect', kind: 'primary', run: async () => {
      const r = await call('hosted:link', pick.value === NEW ? { name: name.value } : { repositoryId: pick.value });
      toast(r.cloned ? `Opened “${r.name}” with its full history` : `“${r.name}” is ready to share`);
      state.initSel = false; state.checked = Date.now();
      await refresh();
    } }],
  });
}

function connectRemoteDialog() {
  const cur = (state.sync && state.sync.remotes[0]) || null;
  const input = h('input', { type: 'text', value: cur ? cur.url : '', placeholder: 'https://github.com/your-team/budgets.git', 'aria-label': 'Remote address' });
  modal({
    title: cur ? 'Change remote address' : 'Connect a remote',
    body: [
      h('p', {}, 'Paste the address of an empty git repository your team can reach — GitHub, GitLab, Azure DevOps, a self-hosted server, or a shared network folder.'),
      input,
      h('p', { class: 'note' }, 'Your files are pushed there as normal git commits. Sign-in uses the git credentials already on this computer; for GitHub, use a personal access token as the password when asked.'),
    ],
    actions: [{ label: 'Cancel' }, { label: 'Save', kind: 'primary', run: guard(async () => {
      const url = input.value.trim();
      if (!url) throw new Error('Enter an address.');
      await call('sync:setRemote', 'origin', url);
      toast('Remote connected'); await refresh();
    }) }],
  });
}

// ---------- changes tab ----------
const OFFICE_LETTER = { xlsx: 'X', docx: 'W', pptx: 'P' };
const SICON = { added: '＋', deleted: '−', modified: '•', renamed: '→' };
const nameParts = (p) => { const i = p.lastIndexOf('/'); return [i < 0 ? '' : p.slice(0, i + 1).replace(/\//g, '\\'), p.slice(i + 1)]; };
function fileLabel(p, status) {
  const [dir, base] = nameParts(p);
  const ext = p.split('.').pop().toLowerCase();
  return [
    h('span', { class: 'office ' + ext, 'aria-hidden': 'true' }, OFFICE_LETTER[ext] || '?'),
    h('span', { class: 'fname', title: p }, dir && h('span', { class: 'dir' }, dir), base),
  ];
}
const sicon = (status) => h('span', { class: 'sicon ' + status, title: LABEL[status], 'aria-label': LABEL[status] }, SICON[status]);

function renderList() {
  const q = ($('fileFilterInput').value || '').toLowerCase();
  const list = $('fileList');
  list.replaceChildren(...state.changes.filter((c) => c.path.toLowerCase().includes(q)).map((c) => {
    const cb = h('input', { type: 'checkbox', 'aria-label': 'Include ' + c.path, onclick: (e) => e.stopPropagation(), onchange: () => { cb.checked ? state.selected.add(c.path) : state.selected.delete(c.path); updateCommitBox(); } });
    cb.checked = state.selected.has(c.path);
    return h('li', { role: 'option', tabindex: '0', 'aria-selected': String(state.current === c.path), onclick: () => selectFile(c.path), oncontextmenu: (e) => { e.preventDefault(); fileMenu(e, c); }, onkeydown: (e) => { if (e.key === 'Enter') selectFile(c.path); } },
      cb, ...fileLabel(c.path, c.status),
      c.open ? h('span', { class: 'warnicon', title: 'Open in Office — may have unsaved edits' }, '⚠ Open') : null,
      sicon(c.status));
  }));
  updateCommitBox();
}
$('fileFilterInput').addEventListener('input', renderList);
function updateCommitBox() {
  const n = state.selected.size, total = state.changes.length;
  $('selectAllLabel').textContent = `${total} changed file${total === 1 ? '' : 's'}`;
  $('selectAll').checked = total > 0 && n === total;
  $('selectAll').indeterminate = n > 0 && n < total;
  const b = state.branches ? state.branches.current : 'main';
  $('commitBtn').innerHTML = '';
  $('commitBtn').append(b !== 'main' ? 'Save to scenario' : state.sync && state.sync.hosted ? 'Save & share with team' : 'Save version');
  $('commitBtn').disabled = !n || !$('summary').value.trim();
  $('suggestBtn').disabled = !n;
  $('avatar').textContent = (state.author || '?').split(/\s+/).map((w) => w[0]).join('').slice(0, 2).toUpperCase();
}
function fileMenu(e, c) {
  openMenu({ x: e.clientX, y: e.clientY }, [
    item('Open in Office', guard(() => call('open:working', c.path))),
    item('Discard changes…', () => discardDialog([c])),
  ]);
}
async function selectFile(p) { state.current = p; renderList(); await showWorkingDiff(); }

// Patterns from the model's own history, shown next to Save version.
async function loadSuggestions() {
  const box = $('suggestions');
  const paths = [...state.selected];
  const ticket = (loadSuggestions.n = (loadSuggestions.n || 0) + 1);
  if (!paths.length) { box.hidden = true; return; }
  let found = [];
  try { found = await call('patterns', paths); } catch { found = []; }
  if (ticket !== loadSuggestions.n) return;
  box.replaceChildren(...found.map((s) => h('div', { class: 'insight pattern' }, h('span', { class: 'ic' }, '↻'), h('div', {},
    h('b', {}, `You usually also update ${s.missing.slice(0, 3).map((m) => m.cell).join(', ')}${s.missing.length > 3 ? '…' : ''}`),
    h('div', { class: 'why' }, `In ${s.missing[0].times} of the last ${s.of} versions that changed ${s.changed.join(', ')}, ${s.missing.length === 1 ? 'that cell' : 'these cells'} changed too. Did you mean to update ${s.missing.length === 1 ? 'it' : 'them'}?`)))));
  box.hidden = !found.length;
}

function emptyState(title, text, actions) {
  $('detail').hidden = true; $('empty').hidden = false;
  $('empty').replaceChildren(...[h('h2', {}, title), h('p', {}, text),
    actions.length ? h('div', { class: 'actions' }, ...actions.map(([t, d, fn]) => h('button', { class: 'action', onclick: guard(fn) }, h('span', {}, h('b', {}, t), h('span', { class: 'note' }, d))))) : null].filter(Boolean));
}
function renderEmpty() {
  emptyState('Welcome to GitXL', 'Every cell, formula and sheet in your models gets a full, reviewable history. Your workbooks stay exactly where they are.',
    [['Open a project folder', 'Pick the folder where your models live, or an empty one to download a shared project', addRepo]]);
}

async function showWorkingDiff(silent) {
  $('empty').hidden = true;
  $('commitMeta').hidden = true; $('historyFiles').hidden = true;
  const c = state.changes.find((x) => x.path === state.current);
  if (!c) {
    // A folder with nothing tracked and nothing changed has no Office files at all — say so, rather
    // than leaving the user staring at an empty list.
    const tracked = await call('tracked').catch(() => []);
    if (!tracked.length && !state.resaved.length) {
      emptyState('No Office files in this folder',
        'OfficeGit tracks Excel, Word and PowerPoint files (.xlsx, .docx, .pptx). This folder doesn’t contain any, so there is nothing to version here.',
        [['Add a different folder', 'Pick a folder that holds your spreadsheets or documents', addRepo],
          ['Show this folder in Explorer', 'Check what is actually in it', () => call('repos:reveal')]]);
      return;
    }
    emptyState('No unsaved changes', state.resaved.length ? `${state.resaved.length} file(s) were re-saved without any real edits, so there is nothing new to save.` : 'Edit your model in Excel and save it. Every changed cell shows up here, ready to save as a version.',
      [['Show folder in Explorer', 'Open the project folder', () => call('repos:reveal')], ['View history', 'See every version and restore any of them', () => setTab('history')]]);
    return;
  }
  $('detail').hidden = false;
  $('detailTitle').textContent = c.path;
  $('detailActions').replaceChildren(h('button', { class: 'secondary', onclick: guard(() => call('open:working', c.path)) }, 'Open in Office'));
  if (!silent) $('diff').replaceChildren(h('div', { class: 'note pad' }, 'Loading diff…'));
  try {
    if (c.error) throw new Error(c.error);
    renderDiff(await call('diff:working', c.path, c.oldPath), c.status === 'renamed' ? `Renamed from ${c.oldPath}` : '');
  } catch (e) { $('diff').replaceChildren(h('div', { class: 'note pad' }, 'This file can’t be compared: ' + e.message)); }
}

// ---------- diff rendering ----------
const tag = (type, detail) => h('span', { class: 'tag ' + type }, `${ICON[type]} ${LABEL[type]}`, detail ? ` · ${detail}` : '');
const DETAIL = { input: 'Input changed', formula: 'Formula changed', calculated: 'Calculated value changed', type: 'Type changed', format: 'Formatting changed', text: 'Text changed' };
const num = (v) => (typeof v === 'string' && /^-?\d+(\.\d+)?(E[-+]?\d+)?$/i.test(v) ? Number(v).toLocaleString(undefined, { maximumFractionDigits: 4 }) : v);
const show = (v) => (v == null ? '—' : v === '' ? '(empty)' : num(v));
const rowEl = (cls, g, ...content) => h('div', { class: 'row ' + cls }, h('span', { class: 'g', 'aria-hidden': 'true' }, g), h('span', { class: 't' }, ...content));
const wordsOf = (ops, keep) => ops.filter((o) => o.t === 'eq' || o.t === keep).map((o) => (o.t === 'eq' ? o.s : h(keep === 'ins' ? 'ins' : 'del', {}, o.s)));
const textRows = (ops) => [rowEl('del', '−', ...wordsOf(ops, 'del')), rowEl('add', '+', ...wordsOf(ops, 'ins'))];

function renderDiff(d, note) {
  renderInsights(d.changes);
  const out = [];
  if (note) out.push(h('div', { class: 'banner' }, note));
  if (!d.changes.length) out.push(h('div', { class: 'note pad' }, 'No meaningful differences. (A file re-saved without edits is not counted as a change.)'));
  const groups = new Map();
  for (const c of d.changes) { const FRIENDLY = { sheet: 'Sheets', sheetmeta: 'Sheet formatting', opaque: 'Other content (styles, masters, media)', name: 'Named ranges', comment: 'Comments' };
    const k = FRIENDLY[c.group] || (c.kind === 'cell' ? 'Sheet ' : '') + (c.group || ''); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(c); }
  for (const [name, list] of groups) {
    const sec = h('div', { class: 'group' }, h('h3', {}, `${name} — ${list.length} change${list.length === 1 ? '' : 's'}`));
    const cells = list.filter((c) => c.kind === 'cell');
    if (cells.length) {
      const f = (x) => (x ? (x.formula != null ? `=${x.formula}` + (x.value != null ? `  → ${num(x.value)}` : '') : show(x.value)) : '—');
      sec.append(h('table', { class: 'cells' }, h('thead', {}, h('tr', {}, ...['Cell', 'Change', 'Before', 'After'].map((x) => h('th', {}, x)))),
        h('tbody', {}, ...cells.map((c) => h('tr', { class: hardcoded(c) ? 'risk' : '' }, h('td', {}, c.label.split('!')[1]), h('td', {}, hardcoded(c) ? h('span', { class: 'tag risk' }, '⚠ Formula replaced by a typed number') : tag(c.type, DETAIL[c.detail])),
          h('td', { class: 'mono before' }, c.before ? '− ' + f(c.before) : ''), h('td', { class: 'mono after' }, c.after ? '+ ' + f(c.after) : ''))))));
    }
    for (const c of list.filter((x) => x.kind !== 'cell')) sec.append(renderOther(c));
    out.push(sec);
  }
  $('diff').replaceChildren(...out);
}

// A typed number where a formula used to be: the most common way a financial model silently goes wrong.
const hardcoded = (c) => c.kind === 'cell' && c.before && c.after && c.before.formula != null && c.after.formula == null;
const cellName = (c) => c.label.replace(/^.*!/, '') + ' (' + c.group + ')';
function renderInsights(changes) {
  const box = $('insights');
  const cells = changes.filter((c) => c.kind === 'cell');
  const hard = cells.filter(hardcoded);
  const inputs = cells.filter((c) => !hardcoded(c) && (c.detail === 'input' || (c.type !== 'modified' && !(c.after || c.before || {}).formula)));
  const formulas = cells.filter((c) => !hardcoded(c) && (c.detail === 'formula' || (c.type !== 'modified' && (c.after || c.before || {}).formula)));
  const calc = cells.filter((c) => c.detail === 'calculated');
  const out = [];
  if (hard.length) out.push(h('div', { class: 'insight risk' }, h('span', { class: 'ic' }, '!'), h('div', {},
    h('b', {}, `${hard.length} formula${hard.length === 1 ? ' was' : 's were'} replaced by a typed number`),
    h('div', { class: 'why' }, `${hard.slice(0, 4).map(cellName).join(', ')}${hard.length > 4 ? '…' : ''}. These cells no longer update when their inputs change. Confirm this is intended before it reaches the approved model.`))));
  if (inputs.length || formulas.length || calc.length) {
    const parts = [];
    if (inputs.length) parts.push(`${inputs.length} input${inputs.length === 1 ? '' : 's'}`);
    if (formulas.length) parts.push(`${formulas.length} formula${formulas.length === 1 ? '' : 's'}`);
    const cause = parts.join(' and ') || 'No inputs or formulas';
    out.push(h('div', { class: 'insight info' }, h('span', { class: 'ic' }, 'Σ'), h('div', {},
      h('b', {}, `${cause} changed${calc.length ? `, moving ${calc.length} calculated cell${calc.length === 1 ? '' : 's'}` : ''}`),
      h('div', { class: 'why' }, calc.length ? 'Review the inputs and formulas; the calculated cells are their knock-on effects.' : 'No calculated results moved.'))));
  }
  box.replaceChildren(...out);
  box.hidden = !out.length;
}

function renderOther(c) {
  const head = h('div', { class: 'rhead' }, tag(c.type, DETAIL[c.detail] || c.detail), h('strong', {}, c.label));
  const rows = [];
  const side = (d) => d && (d.text ?? (d.title ? `Slide “${d.title}”` : d.name ? `${d.type || 'Shape'} “${d.name}”` : 'Slide'));
  if (c.kind === 'paragraph' || c.kind === 'slide-order') {
    if (c.words) rows.push(...textRows(c.words));
    else if (c.type === 'modified') rows.push(rowEl('mod', '~', side(c.after), '   (formatting changed)'));
    else if (c.type === 'moved') rows.push(rowEl('mod', '↔', side(c.after || c.before)));
    else rows.push(rowEl(c.type === 'removed' ? 'del' : 'add', c.type === 'removed' ? '−' : '+', side(c.type === 'removed' ? c.before : c.after)));
  } else if (c.kind === 'shape' || c.kind === 'notes' || c.kind === 'comment') {
    if (c.type === 'modified') {
      if (c.words) rows.push(...textRows(c.words));
      if (c.before.geom !== c.after.geom && c.before.geom != null) rows.push(rowEl('mod', '~', `Position/size: ${c.before.geom} → ${c.after.geom}`));
      if (c.before.name !== c.after.name && c.before.name != null) rows.push(rowEl('mod', '~', `Renamed: ${c.before.name} → ${c.after.name}`));
      if (c.before.fmt !== c.after.fmt && c.before.fmt != null) rows.push(rowEl('mod', '~', 'Shape formatting changed'));
    } else {
      const d = c.type === 'removed' ? c.before : c.after;
      rows.push(rowEl(c.type === 'removed' ? 'del' : 'add', c.type === 'removed' ? '−' : '+', d.text || `${d.type || ''} ${d.name || ''}`.trim()));
    }
  } else {
    const what = { opaque: 'Supporting content (styles, masters, charts, media) changed.', sheetmeta: 'Column widths, merged cells or other sheet formatting changed.', sheet: c.type === 'added' ? 'Sheet added' : 'Sheet removed', name: c.after ? `Refers to ${c.after.ref}` : 'Removed' }[c.kind] || c.label;
    rows.push(rowEl(c.type === 'added' ? 'add' : c.type === 'removed' ? 'del' : 'mod', c.type === 'added' ? '+' : c.type === 'removed' ? '−' : '~', what));
  }
  return h('div', { class: 'blk' }, head, ...rows);
}


// ---------- sharing happens on its own ----------
// Share the current draft if the project is connected; if teammates saved first, bring their work in and retry.
// Scenarios (anything but main) stay private until they are brought into the main model.
async function shareNow(quiet) {
  const s = await call('sync:state');
  if (!s.hosted || (quiet && s.branch !== 'main')) return false;
  const btn = $('syncBtn'); btn.className = 'tb busy'; btn.querySelector('.tb-value').textContent = 'Sharing…';
  try {
    await call('sync:push', s.remotes[0].name, { setUpstream: true });
  } catch (e) {
    if (!/Pull first/i.test(e.message)) { toast(e.message, true); await refresh(); return false; }
    const r = await call('sync:pull', s.remotes[0].name);
    if (r.conflicts && r.conflicts.length) { showConflicts({ branch: r.branch, into: r.into, files: r.conflicts }, 'Combine'); return false; }
    await call('sync:push', s.remotes[0].name, { setUpstream: true });
  }
  if (!quiet) toast('Shared with your team');
  await refresh();
  return true;
}

// ---------- commit actions ----------
async function doCommit(force) {
  try {
    await call('commit', { message: $('summary').value, description: $('description').value, paths: [...state.selected], force });
    $('summary').value = ''; $('description').value = '';
    state.current = null;
    await refresh();
    if (await shareNow(true)) toast('Version saved and shared with your team'); else toast('Version saved');
  } catch (e) {
    if (e.extra && e.extra.locked) {
      confirmBox('Files are open in Office', [h('p', {}, 'These files look open and may have unsaved edits, so the saved version could be out of date:'), h('ul', {}, ...e.extra.locked.map((p) => h('li', {}, p)))], 'Commit anyway', () => doCommit(true), 'primary');
    } else toast(e.message, true);
  }
}
$('commitBtn').addEventListener('click', () => doCommit(false));
$('summary').addEventListener('input', updateCommitBox);
$('summary').addEventListener('keydown', (e) => { if (e.key === 'Enter' && !$('commitBtn').disabled) doCommit(false); });
$('suggestBtn').addEventListener('click', guard(async () => {
  $('suggestBtn').textContent = 'Working…';
  try { $('summary').value = (await call('suggest', [...state.selected])).slice(0, 120); updateCommitBox(); }
  finally { $('suggestBtn').textContent = '✨ Suggest message'; }
}));
$('selectAll').addEventListener('change', () => {
  state.selected = new Set($('selectAll').checked ? state.changes.map((c) => c.path) : []);
  renderList();
});

function discardDialog(changes) {
  const what = { added: 'will be moved to the trash (new file never committed)', deleted: 'will be restored from the last version', modified: 'edits since the last version will be lost', renamed: 'will be reverted to its original name' };
  confirmBox('Discard changes?', [
    h('p', {}, 'This will discard the following. A copy is kept in the Trash for 30 days.'),
    h('ul', {}, ...changes.map((c) => h('li', {}, h('strong', {}, c.path), ' — ', what[c.status]))),
  ], 'Discard', guard(async () => { await call('discard', changes.map((c) => c.path)); await refresh(); toast('Changes discarded (recoverable from Trash)'); }));
}

// ---------- history tab ----------
async function loadHistory() {
  const tracked = await call('tracked');
  const sel = $('fileFilter');
  sel.replaceChildren(h('option', { value: '' }, 'All files'), ...tracked.map((p) => h('option', { value: p }, p)));
  sel.value = state.filter;
  state.commits = await call('log', null, state.filter || null);
  const list = $('commitList');
  list.replaceChildren(...state.commits.map((c) => h('li', {
    class: 'commit', role: 'option', tabindex: '0', 'aria-selected': String(state.commit === c.id), onclick: () => selectCommit(c.id),
    oncontextmenu: (e) => { e.preventDefault(); selectCommit(c.id); commitMenu(e, c); },
    onkeydown: (e) => { if (e.key === 'Enter') selectCommit(c.id); },
  }, h('span', { class: 'ring' + (c.parents.length > 1 ? ' merge' : '') }), h('div', { class: 'body' }, h('div', { class: 'msg' }, c.message),
    h('div', { class: 'meta' }, h('span', { class: 'who' }, c.author), h('span', { class: 'chip' }, c.parents.length > 1 ? 'combined' : (state.branches ? state.branches.current : 'main')), h('span', {}, '· ' + new Date(c.time).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })))))));
  if (!state.commits.length) list.replaceChildren(h('li', { class: 'sub' }, 'No versions yet. Save your first version to start the history.'));
  const keep = state.commits.find((c) => c.id === state.commit);
  if (keep) selectCommit(keep.id, true);
  else if (state.commits[0]) selectCommit(state.commits[0].id, true);
  else emptyState('No history yet', 'Save a version to start the history of your model.', []);
}
$('fileFilter').addEventListener('change', guard(async () => { state.filter = $('fileFilter').value; await loadHistory(); }));

function commitMenu(e, c) {
  openMenu({ x: e.clientX, y: e.clientY }, [
    item('Undo this version…', () => confirmBox('Undo this version?', [h('p', {}, `Creates a new version that undoes “${c.message}”. Nothing is deleted from the history.`)], 'Undo version', guard(async () => {
      const r = await call('revert', c.id); toast(r.skipped.length ? `Undone; skipped ${r.skipped.length} file(s) changed since` : 'Version undone'); await refresh();
    }), 'primary')),
  ]);
}

async function selectCommit(id, keep) {
  state.commit = id;
  const c = state.commits.find((x) => x.id === id);
  [...$('commitList').children].forEach((li, i) => li.setAttribute('aria-selected', String(state.commits[i] && state.commits[i].id === id)));
  $('empty').hidden = true; $('detail').hidden = false;
  $('detailTitle').textContent = c.message;
  $('detailActions').replaceChildren();
  $('commitMeta').hidden = false;
  const initials = (c.author || '?').split(/\s+/).map((w) => w[0]).join('').slice(0, 2).toUpperCase();
  $('commitMeta').replaceChildren(h('div', { class: 'who' }, h('span', { class: 'avatar', 'aria-hidden': 'true' }, initials), `${c.author} saved this ${ago(c.time)} · ${new Date(c.time).toLocaleString()}`), c.description ? h('div', {}, c.description) : null);
  $('historyFiles').hidden = false;
  if (!keep || !c.changed.some((f) => f.path === state.commitFile)) state.commitFile = (state.filter && c.changed.find((f) => f.path === state.filter) ? state.filter : c.changed[0]?.path) || null;
  renderHistFiles(c);
  await showCommitDiff(c);
}
function renderHistFiles(c) {
  const ul = $('historyFiles');
  ul.replaceChildren(...c.changed.map((f) => h('li', {
    role: 'option', tabindex: '0', 'aria-selected': String(f.path === state.commitFile),
    onclick: () => { state.commitFile = f.path; renderHistFiles(c); showCommitDiff(c); },
    onkeydown: (e) => { if (e.key === 'Enter') { state.commitFile = f.path; renderHistFiles(c); showCommitDiff(c); } },
    oncontextmenu: (e) => { e.preventDefault(); versionMenu(e, c, f); },
  }, ...fileLabel(f.path), sicon(f.status))));
  if (!c.changed.length) ul.replaceChildren(h('li', { class: 'sub' }, 'No file changes in this version.'));
}
function versionMenu(e, c, f) {
  const items = [];
  if (f.status !== 'deleted') items.push(
    item('Open in Office (read-only)', guard(() => call('open:version', c.id, f.path))),
    item('Compare with working copy', guard(async () => { $('diff').replaceChildren(h('div', { class: 'note pad' }, 'Loading…')); renderDiff(await call('diff:compare', c.id, f.path), 'Comparing this version → your current working copy'); })),
    item('Restore this version…', () => confirmBox('Restore this version?', [h('p', {}, `“${f.path}” in your folder will be replaced with the version from “${c.message}”. Your current file is kept in the Trash for 30 days.`)], 'Restore', guard(async () => { await call('restore', c.id, f.path); toast('Version restored'); await refresh(); }), 'primary')),
  );
  items.push(item('Undo this version…', () => commitMenu(e, c)));
  openMenu({ x: e.clientX, y: e.clientY }, items);
}
async function showCommitDiff(c) {
  const f = c.changed.find((x) => x.path === state.commitFile);
  if (!f) { $('diff').replaceChildren(); return; }
  $('detailActions').replaceChildren(
    f.status !== 'deleted' && h('button', { class: 'secondary', onclick: (e) => versionMenu({ clientX: e.clientX, clientY: e.clientY }, c, f) }, 'Version actions ▾'));
  $('diff').replaceChildren(h('div', { class: 'note pad' }, 'Loading diff…'));
  try { renderDiff(await call('diff:commit', c.id, f.path)); } catch (e) { $('diff').replaceChildren(h('div', { class: 'note pad' }, 'Can’t compare this file: ' + e.message)); }
}


// ---------- tabs ----------
async function setTab(t) {
  state.tab = t;
  $('tabChanges').setAttribute('aria-selected', String(t === 'changes'));
  $('tabHistory').setAttribute('aria-selected', String(t === 'history'));
  $('changesPane').hidden = t !== 'changes'; $('historyPane').hidden = t !== 'history';
  if (!state.repos.some((r) => r.active)) return;
  if (t === 'history') await loadHistory(); else { $('commitMeta').hidden = true; $('historyFiles').hidden = true; await refresh(); }
}
$('tabChanges').addEventListener('click', guard(() => setTab('changes')));
$('tabHistory').addEventListener('click', guard(() => setTab('history')));

// ---------- toolbar ----------
$('repoBtn').addEventListener('click', (e) => openMenu($('repoBtn'), [
  ...state.repos.map((r) => h('button', { class: 'item' + (r.active ? ' current' : ''), onclick: guard(async () => { closeMenu(); await call('repos:select', r.path); state.initSel = false; state.checked = 0; state.selected.clear(); state.current = null; state.commit = null; await refresh(); }) },
    h('span', { class: 'grow' }, r.name), h('span', { class: 'link', title: 'Remove from list (files are not touched)', onclick: (ev) => { ev.stopPropagation(); closeMenu(); confirmBox('Remove project?', [h('p', {}, `“${r.name}” will be removed from the list. Your files and their history stay on disk.`)], 'Remove', guard(async () => { await call('repos:remove', r.path); await refresh(); })); } }, '✕'))),
  h('hr'), item('Open a project folder…', guard(addRepo)),
  state.repos.some((r) => r.active) && item('Show in Explorer', guard(() => call('repos:reveal'))),
]));

$('branchBtn').addEventListener('click', guard(async () => {
  if (!state.repos.some((r) => r.active)) return;
  const br = await call('branches');
  state.branches = br;
  const filter = h('input', { type: 'text', placeholder: 'Find a draft', 'aria-label': 'Find a draft' });
  const listEl = h('div');
  const draw = () => listEl.replaceChildren(...br.list.filter((b) => b.name.includes(filter.value)).map((b) => h('div', {},
    h('button', { class: 'item' + (b.current ? ' current' : ''), onclick: () => { closeMenu(); switchTo(b.name); } }, h('span', { class: 'grow' }, (b.current ? '✓ ' : '') + b.name)),
  )));
  filter.addEventListener('input', draw); draw();
  openMenu($('branchBtn'), [
    filter, listEl, h('hr'),
    item('New draft…', newBranchDialog),
    item('Rename this draft…', renameDialog),
    item('Delete a draft…', deleteChooser),
    item(`Combine a draft into ${br.current}…`, mergeChooser),
  ]);
  filter.focus();
}));

async function switchTo(name) {
  try {
    const r = await call('branch:switch', name, {});
    if (r.needsAction) return dirtyDialog(name, r.changes);
    state.current = null; state.commit = null; await refresh();
  } catch (e) { toast(e.message, true); }
}
function dirtyDialog(name, changes) {
  modal({
    title: 'You have unsaved changes',
    body: [h('p', {}, `Switching to “${name}” would rewrite these files:`), h('ul', {}, ...changes.map((c) => h('li', {}, `${LABEL[c.status]}: ${c.path}`))),
      h('p', {}, 'Stash keeps them safe so you can bring them back later. Discard throws them away (recoverable from Trash for 30 days).')],
    actions: [{ label: 'Cancel' },
      { label: 'Discard & switch', kind: 'danger', run: guard(async () => { await call('branch:switch', name, { discard: true }); await refresh(); }) },
      { label: 'Stash & switch', kind: 'primary', run: guard(async () => { await call('branch:switch', name, { stash: true }); toast('Changes stashed'); await refresh(); }) }],
  });
}
function newBranchDialog() {
  const input = h('input', { type: 'text', placeholder: 'e.g. Downside case', 'aria-label': 'Scenario name' });
  modal({ title: 'Try a what-if', body: [h('p', {}, 'Make changes in a private copy of the model. The main model is untouched until you bring the scenario in.'), input],
    actions: [{ label: 'Cancel' }, { label: 'Start scenario', kind: 'primary', run: async () => {
      await call('branch:create', input.value.trim().replace(/\s+/g, '-'));
      await refresh();
    } }] });
}
function renameDialog() {
  const input = h('input', { type: 'text', value: state.branches.current, 'aria-label': 'New name' });
  modal({ title: 'Rename draft', body: [input], actions: [{ label: 'Cancel' }, { label: 'Rename', kind: 'primary', run: async () => { await call('branch:rename', state.branches.current, input.value.trim()); await refresh(); } }] });
}
function mergeChooser() {
  const others = state.branches.list.filter((b) => !b.current);
  if (!others.length) return toast('There is no other draft to combine.', true);
  openMenu($('moreBtn'), others.map((b) => item(`Combine “${b.name}” into ${state.branches.current}`, guard(() => startMerge(b.name)))));
}

// Merge runs in git. Our merge driver merges the Office files; anything it can't merge comes back as a
// conflict and the user picks a side.
function showConflicts(m, verb) {
  const picks = {};
  const conflicts = m.files.filter((f) => f.state === 'conflict');
  const clean = m.files.filter((f) => f.state !== 'conflict');
  const update = () => { const b = document.querySelector('.modal footer .primary'); if (b) b.disabled = conflicts.some((f) => !picks[f.path]); };
  const body = [];
  if (clean.length) body.push(h('div', { class: 'insight info' }, h('span', { class: 'ic' }, '✓'), h('div', {}, h('b', {}, `Combined cell by cell: ${clean.map((f) => f.path).join(', ')}`), h('div', { class: 'why' }, 'Edits to different cells were kept from both sides. Formulas recalculate when the file opens.'))));
  if (!conflicts.length) body.push(h('p', {}, 'No cell was changed on both sides, so nothing needs a decision.'));
  for (const f of conflicts) {
    const box = h('div', { class: 'conflict' },
      h('div', { class: 'chead' }, h('span', { class: 'grow' }, f.path), h('span', { class: 'tag modified' }, '⚠ ' + f.summary)));
    if (f.items && f.items.length) {
      box.append(h('table', { class: 'cells' },
        h('thead', {}, h('tr', {}, h('th', {}, 'Cell'), h('th', {}, `Yours (${m.into})`), h('th', {}, `Theirs (${m.branch})`))),
        h('tbody', {}, ...f.items.map((it) => h('tr', {}, h('td', {}, it.label), h('td', { class: 'mono before' }, it.ours), h('td', { class: 'mono after' }, it.theirs))))));
    }
    box.append(h('div', { class: 'custom' },
      h('label', {}, h('input', { type: 'radio', name: 'c-' + f.path, onchange: () => { picks[f.path] = 'ours'; update(); } }), ` Keep my whole file (${m.into})`),
      h('label', {}, h('input', { type: 'radio', name: 'c-' + f.path, onchange: () => { picks[f.path] = 'theirs'; update(); } }), ` Use their whole file (${m.branch})`)));
    body.push(box);
  }
  modal({
    wide: conflicts.length > 0,
    title: `${verb} “${m.branch}” into “${m.into}”`,
    body,
    actions: [
      { label: 'Cancel', run: async () => { await call('merge:abort'); await refresh(); } },
      { label: 'Complete ' + verb.toLowerCase(), kind: 'primary', run: async () => { await call('merge:complete', picks); toast(verb + ' complete'); await refresh(); await shareNow(true); } },
    ],
  });
  update();
}

async function startMerge(branch) {
  const m = await call('merge:start', branch);
  if (m.status === 'uptodate') return toast(`${state.branches.current} already has everything from ${branch}`);
  if (m.status === 'fastforward') { toast(`Brought “${branch}” into the main model`); await refresh(); return shareNow(true); }
  showConflicts(m, 'Combine');
}

function deleteChooser() {
  const others = state.branches.list.filter((b) => !b.current);
  if (!others.length) return toast('There are no other drafts to delete.', true);
  openMenu($('moreBtn'), others.map((b) => item(`Delete “${b.name}”…`, guard(async () => {
    const n = await call('branch:unmerged', b.name);
    confirmBox('Delete draft?', [h('p', {}, `“${b.name}” will be deleted.`), n ? h('p', { class: 'warnicon' }, `⚠ It has ${n} version${n === 1 ? '' : 's'} that exist in no other draft. They will be lost.`) : h('p', {}, 'All its versions exist in other drafts.')],
      'Delete draft', guard(async () => { await call('branch:delete', b.name, n > 0); await refresh(); }));
  }))));
}
// ---------- stash / trash / settings ----------
const openStash = guard(async () => {
  if (!state.repos.some((r) => r.active)) return;
  const list = await call('stash:list');
  const body = [h('p', {}, 'Stash puts uncommitted changes aside so you can switch branches mid-edit.'),
    ...(list.length ? list.map((s) => h('div', { class: 'rowitem' }, h('div', { class: 'grow' }, h('strong', {}, s.message), h('div', { class: 'note' }, `${s.branch} · ${ago(s.at)} · ${Object.keys(s.files).length} file(s)`)),
      h('button', { class: 'secondary', onclick: guard(async () => { await call('stash:apply', s.id); closeModal(); toast('Stash applied'); await refresh(); }) }, 'Apply'),
      h('button', { class: 'secondary', onclick: guard(async () => { await call('stash:drop', s.id); openStash(); }) }, 'Drop'))) : [h('p', { class: 'note' }, 'No stashed changes.')])];
  modal({ title: 'Stashed changes', body, actions: [{ label: 'Close' }, { label: 'Stash current changes', kind: 'primary', run: guard(async () => { await call('stash:save', ''); toast('Changes stashed'); await refresh(); }) }] });
});
// Discarded work goes to git's stash, so the "trash" is just the stash list.
// Discarded work goes to git's stash, so recovering it means popping a stash entry.
const openTrash = guard(async () => {
  if (!hasRepo()) return;
  const list = await call('stash:list');
  modal({
    title: 'Discarded & stashed work',
    body: list.length
      ? list.map((t) => h('div', { class: 'rowitem' },
          h('div', { class: 'grow' }, h('strong', {}, t.message), h('div', { class: 'note' }, `${t.branch} · ${ago(t.at)}`)),
          h('button', { class: 'secondary', onclick: guard(async () => { await call('stash:apply', t.id); closeModal(); toast('Restored'); await refresh(); }) }, 'Restore')))
      : [h('p', { class: 'note' }, 'Nothing here. Anything you discard is kept here so you can get it back.')],
    actions: [{ label: 'Close' }],
  });
});

function openSettings() {
  const input = h('input', { type: 'text', value: state.author, 'aria-label': 'Your name' });
  modal({ title: 'Settings', body: [h('label', {}, 'Your name (shown on versions you save)'), input,
    h('p', { class: 'note' }, 'Local use needs no account. Sync, sharing and AI summaries need the OfficeGit platform, which is not connected in this build.')],
    actions: [{ label: 'Cancel' }, { label: 'Save', kind: 'primary', run: guard(async () => { state.author = await call('settings:author', input.value); }) }] });
}

// ---------- menu bar & shortcuts ----------
const hasRepo = () => state.repos.some((r) => r.active);
const needRepo = (fn) => () => (hasRepo() ? fn() : toast('Add a repository first', true));
const withBranches = (fn) => needRepo(async () => { state.branches = await call('branches'); fn(); });
const exec = (c) => () => document.execCommand(c);
function helpDialog() {
  modal({ title: 'How versions work', body: [
    h('p', {}, h('b', {}, 'Save version'), ' = a checkpoint of your model with a note on what changed and why. You can return to any version later.'),
    h('p', {}, h('b', {}, 'Draft'), ' = a scenario or fix you work on without touching the approved model (main).'),
    h('p', {}, h('b', {}, 'Combine'), ' = bring a draft back into the model. Edits to different cells combine automatically; if both sides changed the same cell, you decide which wins.'),
    h('p', {}, h('b', {}, 'Share / Get latest'), ' = send your versions to the team, or bring theirs down.'),
    h('p', {}, h('b', {}, 'Restore'), ' = put any older version of a file back. The file it replaces goes to the Trash for 30 days.'),
  ], actions: [{ label: 'Close' }] });
}
const MENUS = {
  File: [['Open a project folder…', guard(addRepo), 'Ctrl+O'], ['Remove project…', needRepo(() => $('repoBtn').click()), ''], '-', ['Settings…', () => openSettings(), 'Ctrl+,'], '-', ['Exit', () => call('app:quit'), 'Alt+F4']],
  Edit: [['Undo', exec('undo'), 'Ctrl+Z'], ['Redo', exec('redo'), 'Ctrl+Y'], '-', ['Cut', exec('cut'), 'Ctrl+X'], ['Copy', exec('copy'), 'Ctrl+C'], ['Paste', exec('paste'), 'Ctrl+V'], ['Select all', exec('selectAll'), 'Ctrl+A']],
  View: [['Changes', needRepo(() => setTab('changes')), 'Ctrl+1'], ['History', needRepo(() => setTab('history')), 'Ctrl+2'], '-', ['Reload', () => location.reload(), 'Ctrl+R'], ['Toggle developer tools', () => call('app:devtools'), 'Ctrl+Shift+I']],
  Project: [['Show in Explorer', needRepo(() => call('repos:reveal')), ''], ['Stash changes…', needRepo(() => openStash()), ''], ['Discarded & stashed work…', needRepo(() => openTrash()), '']],
  Draft: [['New draft…', needRepo(newBranchDialog), 'Ctrl+Shift+N'], ['Rename draft…', withBranches(renameDialog), 'Ctrl+Shift+R'], ['Delete draft…', withBranches(deleteChooser), ''], ['Combine a draft…', withBranches(mergeChooser), 'Ctrl+Shift+M']],
  Help: [['How versions work', helpDialog, ''], ['About GitXL', () => modal({ title: 'GitXL', body: [h('p', {}, 'Version control for Excel: every cell, formula and sheet has a reviewable history, drafts combine cell by cell, and your team shares one source of truth.')], actions: [{ label: 'Close' }] }), '']],
};
let menuFromBar = false;
function openBarMenu(name, btn) {
  menuFromBar = true;
  [...$('menubar').children].forEach((b) => b.setAttribute('aria-expanded', String(b === btn)));
  openMenu(btn, MENUS[name].map((m) => (m === '-' ? h('hr') : h('button', { class: 'item', onclick: () => { closeMenu(); m[1](); } }, m[0], m[2] ? h('span', { class: 'kbd' }, m[2]) : null))));
}
for (const name of Object.keys(MENUS)) {
  const b = h('button', { 'aria-haspopup': 'true', 'aria-expanded': 'false', onmousedown: (e) => e.stopPropagation(), onclick: () => (b.getAttribute('aria-expanded') === 'true' && !$('menu').hidden ? closeMenu() : openBarMenu(name, b)),
    onmouseenter: () => { if (!$('menu').hidden && menuFromBar) openBarMenu(name, b); } }, name);
  $('menubar').append(b);
}
const baseClose = closeMenu;
document.addEventListener('mousedown', () => { if ($('menu').hidden) [...$('menubar').children].forEach((b) => b.setAttribute('aria-expanded', 'false')); });
new MutationObserver(() => { if ($('menu').hidden) { menuFromBar = false; [...$('menubar').children].forEach((b) => b.setAttribute('aria-expanded', 'false')); } }).observe($('menu'), { attributes: true, attributeFilter: ['hidden'] });
void baseClose;
document.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey)) return;
  const k = e.key.toLowerCase();
  if (e.shiftKey && k === 'n') { e.preventDefault(); needRepo(newBranchDialog)(); }
  else if (e.shiftKey && k === 'm') { e.preventDefault(); withBranches(mergeChooser)(); }
  else if (k === 'o' && !e.shiftKey) { e.preventDefault(); guard(addRepo)(); }
  else if (k === '1') { e.preventDefault(); needRepo(() => setTab('changes'))(); }
  else if (k === '2') { e.preventDefault(); needRepo(() => setTab('history'))(); }
  else if (k === ',') { e.preventDefault(); openSettings(); }
});

window.og.onFsChanged(() => { if (state.tab === 'changes') refresh().catch(() => {}); });
window.addEventListener('focus', () => { if (state.tab === 'changes' && state.repos.some((r) => r.active)) refresh().catch(() => {}); });
refresh().catch((e) => toast(e.message, true));
