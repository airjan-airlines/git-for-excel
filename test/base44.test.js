'use strict';
// The Base44 REST client, against a stubbed API: no network, no credentials.
const test = require('node:test');
const assert = require('node:assert/strict');
const { createClient, Base44Error } = require('../src/sync/base44');

function stub(handler) {
  const calls = [];
  const fetch = async (url, init) => {
    calls.push({ url: new URL(url), method: init.method, headers: init.headers, body: init.body ? JSON.parse(init.body) : undefined });
    const r = handler(new URL(url), init) || {};
    return {
      ok: r.status === undefined || (r.status >= 200 && r.status < 300),
      status: r.status || 200,
      text: async () => (r.body === undefined ? '' : JSON.stringify(r.body)),
    };
  };
  return { calls, client: createClient({ appId: 'app_123', token: 'tok_secret_abcd', fetch }) };
}

test('sends the documented auth header and URL shape', async () => {
  const { client, calls } = stub(() => ({ body: { schemas: [], total: 0 } }));
  await client.schemas();
  assert.equal(calls[0].url.toString(), 'https://app.base44.com/api/apps/app_123/entity-schemas');
  assert.equal(calls[0].method, 'GET');
  assert.equal(calls[0].headers.authorization, 'Bearer tok_secret_abcd');
});

test('list encodes filters, paging and field selection as query parameters', async () => {
  const { client, calls } = stub(() => ({ body: [] }));
  await client.list('Commit', { q: { projectId: 'p1', size: { $gt: 10 } }, limit: 25, skip: 50, sort: '-created_date', fields: ['id', 'sha'] });
  const u = calls[0].url;
  assert.equal(u.pathname, '/api/apps/app_123/entities/Commit');
  assert.deepEqual(JSON.parse(u.searchParams.get('q')), { projectId: 'p1', size: { $gt: 10 } });
  assert.equal(u.searchParams.get('limit'), '25');
  assert.equal(u.searchParams.get('skip'), '50');
  assert.equal(u.searchParams.get('sort'), '-created_date');
  assert.equal(u.searchParams.get('fields'), 'id,sha');
});

test('create posts a flat JSON body', async () => {
  const { client, calls } = stub(() => ({ body: { id: 'rec_1' } }));
  const out = await client.create('Commit', { sha: 'abc', message: 'Update budget' });
  assert.equal(calls[0].method, 'POST');
  assert.equal(calls[0].headers['content-type'], 'application/json');
  assert.deepEqual(calls[0].body, { sha: 'abc', message: 'Update budget' });
  assert.equal(out.id, 'rec_1');
});

test('upsert creates when absent and updates when present', async () => {
  let existing = [];
  const { client, calls } = stub((url, init) => {
    if (init.method === 'GET') return { body: existing };
    return { body: { id: 'rec_1' } };
  });
  await client.upsert('Commit', { sha: 'abc' }, { message: 'first' });
  assert.equal(calls[1].method, 'POST');
  assert.deepEqual(calls[1].body, { sha: 'abc', message: 'first' }, 'key is merged into the new record');

  existing = [{ id: 'rec_1', sha: 'abc' }];
  await client.upsert('Commit', { sha: 'abc' }, { message: 'second' });
  assert.equal(calls[3].method, 'PUT');
  assert.equal(calls[3].url.pathname, '/api/apps/app_123/entities/Commit/rec_1');
  assert.deepEqual(calls[3].body, { message: 'second' });
});

test('update falls back to PATCH when PUT is not supported', async () => {
  const { client, calls } = stub((url, init) => (init.method === 'PUT' ? { status: 405, body: { detail: 'nope' } } : { body: { id: 'rec_1' } }));
  await client.update('Commit', 'rec_1', { message: 'x' });
  assert.deepEqual(calls.map((c) => c.method), ['PUT', 'PATCH']);
});

test('turns Base44 status codes into advice, and never leaks the key', async () => {
  const cases = [
    [401, /personal access token/i],
    [403, /read-only keys are refused|row-level security/i],
    [404, /BASE44_APP_ID|could not find/i],
    [422, /required field is missing/i],
    [429, /Too many requests/i],
  ];
  for (const [status, re] of cases) {
    const { client } = stub(() => ({ status, body: { detail: 'server said so' } }));
    await assert.rejects(client.list('Commit'), (e) => {
      assert.ok(e instanceof Base44Error);
      assert.equal(e.status, status);
      assert.match(e.message, re);
      assert.ok(!e.message.includes('tok_secret_abcd'), 'error text must not contain the key');
      assert.match(e.message, /server said so/, 'keeps the server detail');
      return true;
    });
  }
});

test('falls back to the legacy api_key header when Bearer is rejected', async () => {
  const { client, calls } = stub((url, init) => (init.headers.authorization ? { status: 401, body: { detail: 'no' } } : { body: [{ id: 'r1' }] }));
  const rows = await client.list('Commit');
  assert.deepEqual(calls.map((c) => (c.headers.authorization ? 'bearer' : 'api_key')), ['bearer', 'api_key']);
  assert.deepEqual(rows, [{ id: 'r1' }]);
  await client.list('Commit');
  assert.equal(calls.length, 3, 'remembers the working style instead of retrying every call');
});

test('describe() redacts the key so it is safe to print', () => {
  const { client } = stub(() => ({}));
  assert.deepEqual(client.describe(), { appId: 'app_123', host: 'https://app.base44.com', token: '…abcd' });
});

test('refuses to start without credentials, naming the variable to set', () => {
  assert.throws(() => createClient({ token: 'x' }), /BASE44_APP_ID/);
  assert.throws(() => createClient({ appId: 'a' }), /BASE44_API_KEY/);
});
