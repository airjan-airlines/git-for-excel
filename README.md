# OfficeGit

Version control for Excel files, with a GitHub Desktop-style UI.

**It is git.** A repository is a real git repository; git does the commits, branches, history and merging.
The part git can't do on its own is Office files — it sees `.xlsx` as an opaque blob — so this project adds:

1. **A semantic diff** — unpacks the OOXML, so you see *"B3: 5 → 4242, input changed"* instead of *"binary files differ"*.
2. **A git merge driver** — so two people editing *different cells* of the same workbook merge automatically.

```bash
npm install
npm start                                   # run the app
npm test                                    # test suite (runs against real .xlsx files in ~/Downloads)
npm run demo -- ~/Downloads /tmp/demo       # build a demo repo with two branches ready to merge
npm run build:base44                        # bundle the engine into the Base44 backend function
```

## The merge driver

Registered per-repository in `.git/config` + `.git/info/attributes` (nothing is written into your folder,
and your global git setup is untouched):

```
*.xlsx -text -diff merge=officegit
merge.officegit.driver = node src/main/merge-driver.js %O %A %B %P
```

Because it lives in git, **plain `git merge` works too** — no app required:

```
$ git merge travel-budget-revision
Auto-merging Budget.xlsx
Merge made by the 'ort' strategy.
```

Edits to different cells merge. Edits to the *same* cell become a normal git conflict, and the app's
conflict screen shows the clashing cells side by side so you can pick a side.

## Layout
- `src/engine` — OOXML → semantic model, diff, three-way merge, rebuild. No native dependencies.
- `src/main/git.js` — thin promise wrapper over the git CLI.
- `src/main/repo.js` — translates git into Office terms (status, diff, branch, merge).
- `src/main/merge-driver.js` — the git merge driver.
- `src/renderer` — the UI.

## What works
Add any folder as a repository, see changed files, semantic cell-level diffs, commit with a suggested
message, browse history, restore any version (byte-for-byte the original file), revert, branch, switch
branches, stash, and merge with conflict resolution. A file re-saved in Excel with no real edits shows
as *no change*. Discarded work goes to `git stash`, so it is recoverable.

## Limits
- **Excel is the tested format.** `.docx` / `.pptx` go through the same pipeline and the merge driver is
  registered for them, but only `.xlsx` is covered by the test suite.
- **Conflict resolution is whole-file** (keep mine / use theirs). Per-cell resolution is not wired up.
- **Visual diff** needs LibreOffice; that tab is disabled.
- **Remote sync** works with any git host (GitHub, GitLab, Gitea, a network share). See [base44/README.md](base44/README.md)
  for connecting the Base44 web app on top of a GitHub remote.
- Windows is the only platform tested.

## Dev
`OG_REPO=<folder> OG_SHOT=out.png OG_JS="<js>" npx electron .` opens the app on a folder, optionally runs a
snippet in the page, saves a screenshot and quits.
