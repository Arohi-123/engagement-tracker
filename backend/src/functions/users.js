const { app } = require('@azure/functions');
const { authenticate } = require('../lib/authenticate');
const { listAccessRows, upsertAccessRow, deleteAccessRow } = require('../lib/accessTable');
const { REGIONS } = require('../lib/regions');

const VALID_ROLES = ['SuperAdmin', 'GlobalUser', 'GlobalViewer', 'RegionalUser', 'RegionalViewer'];
const REGIONAL_ROLES = ['RegionalUser', 'RegionalViewer'];

// Every endpoint here is SuperAdmin-only — this is the one place allowed to see
// or change the full AccessControl table. All other endpoints only ever resolve
// the caller's own row (see authenticate.js / me.js).
async function requireSuperAdmin(request) {
  const user = await authenticate(request);
  if (user.role !== 'SuperAdmin') {
    const err = new Error('Only Super Admin can manage access'); err.status = 403; throw err;
  }
  return user;
}

function errResponse(err) {
  return { status: err.status || 500, jsonBody: { error: err.message || 'Internal error' } };
}

// GET /api/users — full access list, Super Admin panel's main table.
app.http('listUsers', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'users',
  handler: async (request, context) => {
    try { await requireSuperAdmin(request); } catch (err) { return errResponse(err); }
    try {
      return { status: 200, jsonBody: await listAccessRows() };
    } catch (err) {
      context.error(err);
      return { status: 502, jsonBody: { error: 'Failed to read access list' } };
    }
  }
});

// POST /api/users — add or change someone's role/regions (upsert on email).
// A Regional role can now hold more than one region.
app.http('upsertUser', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'users',
  handler: async (request, context) => {
    let caller;
    try { caller = await requireSuperAdmin(request); } catch (err) { return errResponse(err); }

    let body;
    try { body = await request.json(); } catch { return { status: 400, jsonBody: { error: 'Invalid JSON body' } }; }

    const email = String(body.email || '').trim().toLowerCase();
    const role = String(body.role || '').trim();
    const regions = Array.isArray(body.regions) ? [...new Set(body.regions.map(r => String(r).trim().toUpperCase()))] : [];

    if (!email || !email.includes('@')) return { status: 400, jsonBody: { error: 'A valid email is required' } };
    if (!VALID_ROLES.includes(role)) return { status: 400, jsonBody: { error: `Role must be one of: ${VALID_ROLES.join(', ')}` } };
    if (REGIONAL_ROLES.includes(role)) {
      const invalid = regions.filter(r => !REGIONS[r]);
      if (!regions.length || invalid.length) return { status: 400, jsonBody: { error: `${role} requires at least one valid region (${Object.keys(REGIONS).join(', ')})` } };
    }

    try {
      const savedRegions = REGIONAL_ROLES.includes(role) ? regions : [];
      await upsertAccessRow(email, role, savedRegions);
      return { status: 200, jsonBody: { email, role, regions: savedRegions } };
    } catch (err) {
      context.error(err);
      return { status: 502, jsonBody: { error: 'Failed to save access row' } };
    }
  }
});

// DELETE /api/users/{email} — revoke access. Blocks removing your own row so a
// Super Admin can't accidentally lock everyone (including themselves) out.
app.http('deleteUser', {
  methods: ['DELETE'],
  authLevel: 'anonymous',
  route: 'users/{email}',
  handler: async (request, context) => {
    let caller;
    try { caller = await requireSuperAdmin(request); } catch (err) { return errResponse(err); }

    const email = decodeURIComponent(request.params.email || '').trim().toLowerCase();
    if (!email) return { status: 400, jsonBody: { error: 'Email is required' } };
    if (email === caller.email) return { status: 400, jsonBody: { error: "You can't remove your own access. Have another Super Admin do it." } };

    try {
      await deleteAccessRow(email);
      return { status: 200, jsonBody: { success: true } };
    } catch (err) {
      context.error(err);
      return { status: 502, jsonBody: { error: 'Failed to remove access row' } };
    }
  }
});
