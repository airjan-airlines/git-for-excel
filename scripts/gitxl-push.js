'use strict';
// Push a local repository into the GitXL Base44 app.
//   node scripts/gitxl-push.js <repo-folder> [--name "Display name"] [--owner "Your Name"]
// Requires BASE44_APP_ID and BASE44_API_KEY in the environment.
const path = require('path');
const { Repo } = require('../src/main/repo');
const { fromEnv } = require('../src/sync/base44');
const gitxl = require('../src/sync/gitxl');

const arg = (flag) => { const i = process.argv.indexOf(flag); return i > 0 ? process.argv[i + 1] : undefined; };

(async () => {
  const folder = process.argv[2];
  if (!folder || folder.startsWith('--')) throw new Error('usage: gitxl-push.js <repo-folder> [--name N] [--owner O]');
  const repo = await new Repo(path.resolve(folder)).open();
  const client = fromEnv();
  const who = client.describe();
  console.log(`pushing ${path.resolve(folder)} -> app ${who.appId} (key ${who.token})\n`);

  const result = await gitxl.push(repo, client, {
    name: arg('--name'),
    owner: arg('--owner'),
    onProgress: (p) => {
      if (p.step === 'repository') console.log(`repository "${p.name}" ${p.created ? 'created' : 'already existed'}`);
      if (p.step === 'commit') console.log(`  commit ${p.hash}  ${p.changes.toString().padStart(4)} change(s)  ${p.message}`);
      if (p.step === 'branch') console.log(`branch  ${p.name}`);
    },
  });
  console.log(`\ndone: ${result.commitsPushed} new commit(s), ${result.commitsTotal} total, ${result.branches} branch(es)`);
  console.log(`repository id ${result.repositoryId}`);
})().catch((e) => { console.error('\n' + e.message); process.exit(1); });
