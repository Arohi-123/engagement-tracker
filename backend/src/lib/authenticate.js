const { verifyToken } = require('./verifyToken');
const { getAccessRow } = require('./accessTable');

// Shared by every data endpoint: validates the bearer token, resolves the
// caller's own {email, role, regions}. Throws an Error with a `.status` on
// any failure so handlers can just catch-and-return it.
async function authenticate(request) {
  const authHeader = request.headers.get('authorization') || '';
  const token = authHeader.replace(/^Bearer\s+/i, '');
  if (!token) {
    const err = new Error('Missing bearer token'); err.status = 401; throw err;
  }

  let claims;
  try {
    claims = await verifyToken(token);
  } catch {
    const err = new Error('Invalid token'); err.status = 401; throw err;
  }

  const email = String(claims.preferred_username || claims.upn || claims.email || '').toLowerCase();
  if (!email) {
    const err = new Error('Token has no email/UPN claim'); err.status = 401; throw err;
  }

  const row = await getAccessRow(email);
  if (!row || !row.role) {
    const err = new Error('No access configured for this account'); err.status = 403; throw err;
  }

  return { email, role: row.role, regions: row.regions };
}

module.exports = { authenticate };
