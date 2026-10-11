'use strict';
// git merge driver for Office files.  Registered as:
//   .gitattributes:  *.xlsx -text -diff merge=officegit
//   config:          merge.officegit.driver = node <this file> %O %A %B %P
// git calls it with the common ancestor (%O), our version (%A, also the output path) and theirs (%B).
// Exit 0 = merged cleanly (result written to %A); exit 1 = conflict, git marks the file conflicted.
const fs = require('fs');
const engine = require('../engine');

(async () => {
  const [base, ours, theirs, label] = process.argv.slice(2);
  const name = label || 'file';
  try {
    const r = await engine.mergeFiles(fs.readFileSync(base), fs.readFileSync(ours), fs.readFileSync(theirs));
    if (r.buffer) {
      fs.writeFileSync(ours, r.buffer);
      process.exit(0);
    }
    if (r.conflicts && r.conflicts.length) {
      console.error(`${name}: ${r.conflicts.length} overlapping edit(s): ${r.conflicts.slice(0, 5).map((c) => c.label).join(', ')}`);
    } else {
      console.error(`${name}: ${r.unsupported || 'cannot be merged automatically'}`);
    }
    process.exit(1);
  } catch (e) {
    console.error(`${name}: ${e.message}`);
    process.exit(1);
  }
})();
