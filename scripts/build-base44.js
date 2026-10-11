'use strict';
// Copies the engine + reader into the Base44 function folder so the same diff code runs on the server.
//   node scripts/build-base44.js
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const lib = path.join(root, 'base44', 'functions', 'officegit-sync', 'lib');
const copy = (from, to) => {
  fs.mkdirSync(path.dirname(to), { recursive: true });
  fs.copyFileSync(from, to);
};

fs.rmSync(lib, { recursive: true, force: true });
for (const f of fs.readdirSync(path.join(root, 'src', 'engine'))) copy(path.join(root, 'src', 'engine', f), path.join(lib, 'engine', f));
copy(path.join(root, 'src', 'sync', 'reader.js'), path.join(lib, 'sync', 'reader.js'));

// Declare the engine's two dependencies for the function's own install.
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
fs.writeFileSync(path.join(lib, 'package.json'), JSON.stringify({
  name: 'officegit-engine',
  private: true,
  type: 'commonjs',
  dependencies: { jszip: pkg.dependencies.jszip, 'fast-xml-parser': pkg.dependencies['fast-xml-parser'] },
}, null, 2) + '\n');

// Sanity check: the copied tree must load on its own.
delete require.cache[require.resolve(path.join(lib, 'sync', 'reader.js'))];
const reader = require(path.join(lib, 'sync', 'reader.js'));
if (typeof reader.getDiff !== 'function') throw new Error('copied reader did not load');
console.log(`base44 function lib built at ${path.relative(root, lib)} (payload v${reader.PAYLOAD_VERSION})`);
