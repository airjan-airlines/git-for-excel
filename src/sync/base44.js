'use strict';
// Minimal client for the Base44 Entities REST API (beta), used to push local versions into a Base44 app.
//
//   base   https://app.base44.com/api/apps/{appId}
//   auth   Authorization: Bearer <personal access token>   (read-only and workspace keys are refused)
//
// The token is read from the environment and never logged: errors quote the response body, which Base44
// does not echo credentials into, and `describe()` redacts it.
const DEFAULT_HOST = 'https://app.base44.com';

class Base44Error extends Error {
  constructor(status, message, body) {
    super(message);
    this.status = status;
    this.body = body;
  }
}

// Base44's status codes, in words a user can act on.
const EXPLAIN = {
  401: 'Base44 rejected the API key. Check BASE44_API_KEY is a personal access token (not a workspace key).',
  403: 'The key is valid but not allowed to do this. Read-only keys are refused, and the entity may have a row-level security rule blocking it.',
  404: 'Base44 could not find that app or entity. Check BASE44_APP_ID and the entity name.',
  422: 'Base44 rejected the record: a required field is missing or a value has the wrong type.',
  429: 'Too many requests to Base44 — it allows about 80 writes and 100 reads a minute per app.',
};

function createClient({ appId, token, host = DEFAULT_HOST, fetch: f = globalThis.fetch } = {}) {
  if (!appId) throw new Error('Base44 app id is missing. Set BASE44_APP_ID (the id in https://app.base44.com/apps/<id>/editor).');
  if (!token) throw new Error('Base44 API key is missing. Set BASE44_API_KEY to a personal access token.');
  const base = `${host}/api/apps/${encodeURIComponent(appId)}`;

  // Personal access tokens go in `Authorization: Bearer`. Legacy account API keys used an `api_key`
  // header and stop working on 2026-10-15; we retry with it once so an older key still reports something
  // more useful than a bare 401.
  let authStyle = 'bearer';
  const authHeader = (style) => (style === 'bearer' ? { authorization: `Bearer ${token}` } : { api_key: token });

  async function send(method, url, body, style) {
    return f(url, {
      method,
      headers: { ...authHeader(style), ...(body ? { 'content-type': 'application/json' } : {}) },
      body: body ? JSON.stringify(body) : undefined,
    });
  }

  async function call(method, pathname, { body, query } = {}) {
    const url = new URL(base + pathname);
    for (const [k, v] of Object.entries(query || {})) if (v !== undefined && v !== null) url.searchParams.set(k, v);
    let res = await send(method, url.toString(), body, authStyle);
    if (res.status === 401 && authStyle === 'bearer') {
      const retry = await send(method, url.toString(), body, 'api_key');
      if (retry.ok) { authStyle = 'api_key'; res = retry; } // a legacy account key: remember and carry on
    }
    const text = await res.text().catch(() => '');
    let parsed = null;
    try { parsed = text ? JSON.parse(text) : null; } catch { /* non-JSON error body */ }
    if (!res.ok) {
      const detail = (parsed && (parsed.detail || parsed.message)) || text.slice(0, 300);
      throw new Base44Error(res.status, `${EXPLAIN[res.status] || `Base44 returned ${res.status}.`}${detail ? ` (${detail})` : ''}`, parsed);
    }
    return parsed;
  }

  return {
    appId,
    describe: () => ({ appId, host, token: token ? `…${String(token).slice(-4)}` : null }),

    // What entities this app actually has, and their fields. Personal access tokens can read entities but
    // are refused by /entity-schemas, so fall back to the app's public metadata, which carries the same schemas.
    async schemas() {
      try { return await call('GET', '/entity-schemas'); } catch (e) {
        if (e.status !== 401 && e.status !== 403) throw e;
        const res = await f(`${host}/api/apps/public/prod/by-id/${encodeURIComponent(appId)}`);
        if (!res.ok) throw e;
        const { entities = {} } = await res.json();
        const schemas = Object.entries(entities).map(([entity_name, entity_schema]) => ({ entity_name, entity_schema }));
        return { schemas, total: schemas.length };
      }
    },

    // q is a filter object, e.g. { status: 'open' } or { amount: { $gt: 100 } }
    list: (entity, { q, limit, skip, sort, fields } = {}) =>
      call('GET', `/entities/${encodeURIComponent(entity)}`, {
        query: { q: q ? JSON.stringify(q) : undefined, limit, skip, sort, fields: Array.isArray(fields) ? fields.join(',') : fields },
      }),

    get: (entity, id) => call('GET', `/entities/${encodeURIComponent(entity)}/${encodeURIComponent(id)}`),
    create: (entity, data) => call('POST', `/entities/${encodeURIComponent(entity)}`, { body: data }),

    // The reference documents create and list; update is PUT, with a PATCH fallback for older apps.
    async update(entity, id, data) {
      const p = `/entities/${encodeURIComponent(entity)}/${encodeURIComponent(id)}`;
      try { return await call('PUT', p, { body: data }); } catch (e) {
        if (e.status === 404 || e.status === 405) return call('PATCH', p, { body: data });
        throw e;
      }
    },

    // Base44 file storage: upload returns a public URL; download fetches it back as a Buffer.
    async upload(buffer, filename) {
      const form = new FormData();
      form.append('file', new Blob([buffer], { type: 'application/octet-stream' }), filename);
      const res = await f(`${base}/integration-endpoints/Core/UploadFile`, { method: 'POST', headers: authHeader(authStyle), body: form });
      const text = await res.text().catch(() => '');
      if (!res.ok) throw new Base44Error(res.status, `${EXPLAIN[res.status] || `Base44 returned ${res.status} on upload.`} ${text.slice(0, 200)}`.trim());
      const { file_url: url } = JSON.parse(text);
      if (!url) throw new Error('Base44 accepted the upload but returned no file URL.');
      return url;
    },
    async download(url) {
      const res = await f(url, { method: 'GET' });
      if (!res.ok) throw new Base44Error(res.status, `Could not download ${url} from Base44 (${res.status}).`);
      return Buffer.from(await res.arrayBuffer());
    },

    // Create or update by a natural key, so re-running a sync never duplicates rows.
    async upsert(entity, key, data) {
      const found = await this.list(entity, { q: key, limit: 1 });
      const rows = Array.isArray(found) ? found : (found && found.items) || [];
      return rows.length ? this.update(entity, rows[0].id, data) : this.create(entity, { ...key, ...data });
    },

    call,
  };
}

// Credentials: the environment wins, else base44.local.json at the project root (gitignored), so the desktop app
// works out of the box without the key ever being committed.
function localConfig() {
  try { return JSON.parse(require('fs').readFileSync(require('path').join(__dirname, '..', '..', 'base44.local.json'), 'utf8')); } catch { return {}; }
}

function fromEnv(env = process.env, overrides = {}) {
  const local = localConfig();
  return createClient({
    appId: env.BASE44_APP_ID || local.appId,
    token: env.BASE44_API_KEY || local.apiKey,
    host: env.BASE44_HOST || local.host || DEFAULT_HOST,
    ...overrides,
  });
}

module.exports = { createClient, fromEnv, Base44Error, DEFAULT_HOST };
