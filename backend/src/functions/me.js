const { app } = require('@azure/functions');
const { verifyToken } = require('../lib/verifyToken');
const { getAccessRow } = require('../lib/accessTable');

// GET /api/me — resolves the caller's own {role, regions} from their MSAL token.
// Deliberately returns only the caller's own row, never the full AccessControl
// table, so no signed-in user can see anyone else's access rights.
app.http('me', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'me',
  handler: async (request, context) => {
    const authHeader = request.headers.get('authorization') || '';
    const token = authHeader.replace(/^Bearer\s+/i, '');
    if (!token) {
      return { status: 401, jsonBody: { error: 'Missing bearer token' } };
    }

    let claims;
    try {
      claims = await verifyToken(token);
    } catch (err) {
      context.warn('Token verification failed:', err.message);
      return { status: 401, jsonBody: { error: 'Invalid token' } };
    }

    const email = String(claims.preferred_username || claims.upn || claims.email || '').toLowerCase();
    if (!email) {
      return { status: 401, jsonBody: { error: 'Token has no email/UPN claim' } };
    }

    const row = await getAccessRow(email);
    if (!row) {
      return { status: 200, jsonBody: { role: null, regions: [] } };
    }

    return { status: 200, jsonBody: { role: row.role, regions: row.regions } };
  }
});
