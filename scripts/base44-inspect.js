'use strict';
// Read-only: prints the entities a Base44 app defines, so the sync can be written against the real schema.
//
//   setx BASE44_APP_ID  <id from https://app.base44.com/apps/<id>/editor>
//   setx BASE44_API_KEY <personal access token>     (open a NEW terminal after setx)
//   node scripts/base44-inspect.js [--records]
//
// --records also prints up to 3 example rows per entity, with long values truncated.
const { fromEnv, Base44Error } = require('../src/sync/base44');

const short = (v) => {
  const s = typeof v === 'string' ? v : JSON.stringify(v);
  if (s === undefined) return 'undefined';
  return s.length > 80 ? s.slice(0, 77) + '…' : s;
};

(async () => {
  const client = fromEnv();
  const who = client.describe();
  console.log(`app ${who.appId} @ ${who.host} (key ${who.token})\n`);

  const { schemas, total } = await client.schemas();
  console.log(`${total ?? schemas.length} entities\n`);
  for (const { entity_name: name, entity_schema: schema } of schemas) {
    const props = (schema && schema.properties) || {};
    const required = new Set((schema && schema.required) || []);
    console.log(`── ${name}${schema && schema.description ? ` — ${schema.description}` : ''}`);
    for (const [field, def] of Object.entries(props)) {
      const type = def.type || (def.anyOf ? def.anyOf.map((a) => a.type).join('|') : '?');
      const extra = [def.format, def.enum && `one of ${def.enum.join('/')}`, def.items && `of ${def.items.type || '?'}`].filter(Boolean).join(', ');
      console.log(`     ${required.has(field) ? '*' : ' '} ${field.padEnd(22)} ${type}${extra ? '  (' + extra + ')' : ''}`);
    }
    if (!Object.keys(props).length) console.log('       (no declared fields)');

    if (process.argv.includes('--records')) {
      try {
        const rows = await client.list(name, { limit: 3, sort: '-created_date' });
        const items = Array.isArray(rows) ? rows : (rows && rows.items) || [];
        console.log(`     ${items.length} recent record(s)`);
        for (const r of items) console.log('       ·', Object.entries(r).map(([k, v]) => `${k}=${short(v)}`).join('  '));
      } catch (e) { console.log(`     (could not read records: ${e.message})`); }
    }
    console.log();
  }
  console.log('* = required field');
})().catch((e) => {
  console.error(e instanceof Base44Error ? `\n${e.message}` : `\n${e.message}`);
  process.exit(1);
});
