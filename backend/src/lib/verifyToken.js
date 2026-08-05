const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const TENANT_ID = process.env.AAD_TENANT_ID;
const CLIENT_ID = process.env.AAD_CLIENT_ID;

const jwks = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`
});

function getSigningKey(header) {
  return new Promise((resolve, reject) => {
    jwks.getSigningKey(header.kid, (err, key) => {
      if (err) return reject(err);
      resolve(key.getPublicKey());
    });
  });
}

// Verifies the ID token MSAL issues to signed-in users of this same app registration
// (aud === our own client ID) — signature, issuer and expiry are all checked, so this
// is a real verification, not just a decode. Returns the token's claims on success.
async function verifyToken(token) {
  const decoded = jwt.decode(token, { complete: true });
  if (!decoded) throw new Error('Malformed token');

  const publicKey = await getSigningKey(decoded.header);

  return jwt.verify(token, publicKey, {
    algorithms: ['RS256'],
    audience: CLIENT_ID,
    issuer: [
      `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
      `https://sts.windows.net/${TENANT_ID}/`
    ]
  });
}

module.exports = { verifyToken };
