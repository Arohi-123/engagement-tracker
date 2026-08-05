const { getAppGraphToken } = require('./graphAppToken');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

async function graphGet(url) {
  const token = await getAppGraphToken();
  let items = [], next = url;
  while (next) {
    const r = await fetch(next, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) throw new Error(`Graph GET ${r.status}: ${await r.text()}`);
    const j = await r.json();
    items = items.concat(j.value || []);
    next = j['@odata.nextLink'] || null;
  }
  return items;
}

async function graphPost(url, body) {
  const token = await getAppGraphToken();
  const r = await fetch(url, { method: 'POST', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error(`Graph POST ${r.status}: ${await r.text()}`);
  return r.json();
}

async function graphPatch(url, body) {
  const token = await getAppGraphToken();
  const r = await fetch(url, { method: 'PATCH', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error(`Graph PATCH ${r.status}: ${await r.text()}`);
  return r.json();
}

module.exports = { GRAPH_BASE, graphGet, graphPost, graphPatch };
