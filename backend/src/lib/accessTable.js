const { TableClient } = require('@azure/data-tables');

const TABLE_NAME = 'AccessControl';
// Fixed partition — row count here is small (one row per person who has access),
// no need to shard by anything.
const PARTITION_KEY = 'user';

let client;
function getClient() {
  if (!client) {
    client = TableClient.fromConnectionString(process.env.TABLES_CONNECTION_STRING, TABLE_NAME);
  }
  return client;
}

// Table Storage has no array type, so multiple regions (a Regional role can now
// hold more than one) are stored as a single comma-separated string and split
// back into an array on every read.
function parseRegions(raw) {
  return raw ? String(raw).split(',').map(r => r.trim()).filter(Boolean) : [];
}

// RowKey is the lowercased email — callers must lowercase before calling both
// this and any write path (see Super Admin panel, added in Phase 4).
async function getAccessRow(email) {
  try {
    const entity = await getClient().getEntity(PARTITION_KEY, email);
    return { role: entity.role, regions: parseRegions(entity.regions) };
  } catch (err) {
    if (err.statusCode === 404) return null;
    throw err;
  }
}

// Phase 4 (Super Admin panel) — every row, for the management screen. Never
// exposed to non-SuperAdmin callers (enforced in src/functions/users.js, not here).
async function listAccessRows() {
  const rows = [];
  for await (const e of getClient().listEntities()) {
    rows.push({ email: e.rowKey, role: e.role, regions: parseRegions(e.regions) });
  }
  return rows.sort((a, b) => a.email.localeCompare(b.email));
}

async function upsertAccessRow(email, role, regions) {
  const entity = { partitionKey: PARTITION_KEY, rowKey: email, role, regions: (regions && regions.length) ? regions.join(',') : null };
  await getClient().upsertEntity(entity, 'Replace');
}

async function deleteAccessRow(email) {
  try {
    await getClient().deleteEntity(PARTITION_KEY, email);
  } catch (err) {
    if (err.statusCode !== 404) throw err;
  }
}

module.exports = { getAccessRow, listAccessRows, upsertAccessRow, deleteAccessRow, PARTITION_KEY, TABLE_NAME };
