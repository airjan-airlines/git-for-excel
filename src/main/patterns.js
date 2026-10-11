'use strict';
// Patterns from a model's own history: "when this cell changes, these others change with it".
// Each saved version is reduced to the set of cells someone actually edited (inputs and formulas, not values that
// merely recalculated). If the cells being edited now usually travel with others that haven't been touched, the
// analyst may have forgotten them - the classic missed-update error in a financial model.
const OFFICE_CELL = /^cell:/;

// Cells a person edited in one diff: typed inputs, formulas, added/removed cells. Recalculated results are effects.
const edited = (changes) => changes.filter((c) => OFFICE_CELL.test(c.id) && c.detail !== 'calculated' && c.detail !== 'format').map((c) => c.id);

// Row label for a cell: the left-most text in the same row ("Revenue"), so suggestions read like the model.
function rowLabel(tree, id) {
  const m = /^cell:(.*)!([A-Z]+)(\d+)$/.exec(id);
  if (!m || !tree) return null;
  const [, sheet, , row] = m;
  let best = null;
  for (const [k, u] of Object.entries(tree.units)) {
    const r = /^cell:(.*)!([A-Z]+)(\d+)$/.exec(k);
    if (!r || r[1] !== sheet || r[3] !== row || u.data.type !== 'str' || !u.data.value) continue;
    if (!best || r[2].length < best.col.length || (r[2].length === best.col.length && r[2] < best.col)) best = { col: r[2], text: String(u.data.value) };
  }
  return best && best.text;
}
const pretty = (tree, id) => {
  const addr = id.slice(id.lastIndexOf('!') + 1);
  const label = rowLabel(tree, id);
  return label ? `${label} (${addr})` : id.slice(5);
};

class Patterns {
  constructor(repo) { this.repo = repo; this.byCommit = new Map(); }

  // [{ id, file, cells:Set }] for recent single-parent versions (merges only repeat their parents' edits)
  async history(limit = 60) {
    const out = [];
    for (const c of await this.repo.log(null, null, limit)) {
      if (c.parents.length !== 1) continue;
      for (const f of c.changed.filter((x) => /\.xlsx$/i.test(x.path) && x.status === 'modified')) {
        const key = `${c.id}:${f.path}`;
        if (!this.byCommit.has(key)) {
          try { this.byCommit.set(key, new Set(edited((await this.repo.diffCommit(c.id, f.path)).changes))); } catch { this.byCommit.set(key, new Set()); }
        }
        out.push({ id: c.id, file: f.path, cells: this.byCommit.get(key) });
      }
    }
    return out;
  }

  // Suggestions for the files being edited right now.
  async suggest(paths, { minTimes = 2, minShare = 0.6 } = {}) {
    const hist = await this.history();
    const out = [];
    for (const p of paths.filter((x) => /\.xlsx$/i.test(x))) {
      let now;
      try { now = new Set(edited((await this.repo.diffWorking(p)).changes)); } catch { continue; }
      if (!now.size) continue;
      const related = hist.filter((h) => h.file === p && [...now].some((id) => h.cells.has(id)));
      if (related.length < minTimes) continue;
      const counts = new Map();
      for (const h of related) for (const id of h.cells) if (!now.has(id)) counts.set(id, (counts.get(id) || 0) + 1);
      const missing = [...counts].filter(([, n]) => n >= minTimes && n / related.length >= minShare).sort((a, b) => b[1] - a[1]).slice(0, 6);
      if (!missing.length) continue;
      const tree = await this.repo.treeOfWorking(p).catch(() => null);
      out.push({
        file: p,
        changed: [...now].slice(0, 3).map((id) => pretty(tree, id)),
        missing: missing.map(([id, n]) => ({ cell: pretty(tree, id), times: n })),
        of: related.length,
      });
    }
    return out;
  }
}

module.exports = { Patterns, edited, rowLabel };
