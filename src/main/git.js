'use strict';
// Thin promise wrapper around the git CLI. Git is the version control system; this app is a friendly
// front end for it plus an Office-aware diff/merge driver.
const { execFile } = require('child_process');
const path = require('path');

const EMPTY_TREE = '4b825dc642cb6eb9a060e54bf8d69288fbee4904';

function run(cwd, args, { buffer = false, input, maxBuffer = 256 * 1024 * 1024 } = {}) {
  return new Promise((resolve, reject) => {
    const child = execFile('git', args, { cwd, maxBuffer, encoding: buffer ? 'buffer' : 'utf8', windowsHide: true }, (err, stdout, stderr) => {
      if (err) {
        const e = new Error(((buffer ? stderr.toString() : stderr) || err.message).trim());
        e.code = err.code;
        e.stderr = buffer ? stderr.toString() : stderr;
        e.stdout = buffer ? stdout.toString() : stdout;
        return reject(e);
      }
      resolve(buffer ? stdout : stdout.replace(/\r?\n$/, ''));
    });
    if (input !== undefined) { child.stdin.end(input); }
  });
}

class Git {
  constructor(root) { this.root = root; }
  git(...args) { return run(this.root, args); }
  raw(...args) { return run(this.root, args, { buffer: true }); }
  async try(...args) { try { return { ok: true, out: await this.git(...args) }; } catch (e) { return { ok: false, error: e.message }; } }

  static async topLevel(dir) {
    try { return path.resolve(await run(dir, ['rev-parse', '--show-toplevel'])); } catch { return null; }
  }
  static async version() { return run(process.cwd(), ['--version']); }

  async init(defaultBranch = 'main') {
    await this.git('init', '-b', defaultBranch).catch(async () => {
      await this.git('init');
      await this.git('checkout', '-B', defaultBranch);
    });
  }

  async head() { try { return await this.git('rev-parse', 'HEAD'); } catch { return null; } }
  // --show-current works before the first commit (where rev-parse HEAD fails) and prints
  // nothing on a detached HEAD, where the short sha is the useful answer.
  async branch() {
    const r = await this.try('branch', '--show-current');
    if (r.ok && r.out) return r.out;
    const d = await this.try('rev-parse', '--short', 'HEAD');
    return d.ok && d.out ? d.out : 'main';
  }
  async hasCommits() { return (await this.head()) !== null; }
  async baseTree() { return (await this.hasCommits()) ? 'HEAD' : EMPTY_TREE; }

  // null-separated list -> array of records
  async zList(...args) {
    const out = await this.git(...args);
    return out.split('\0').filter((s) => s.length);
  }
}

module.exports = { Git, EMPTY_TREE, run };
