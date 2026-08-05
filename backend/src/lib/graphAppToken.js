// Acquires an app-only Graph access token via the client-credentials grant —
// this is what lets the backend read/write SharePoint using its own identity
// (the Sites.Selected grant from Phase 1), never the calling user's.
const TENANT_ID = process.env.AAD_TENANT_ID;
const CLIENT_ID = process.env.AAD_CLIENT_ID;
const CLIENT_SECRET = process.env.AAD_CLIENT_SECRET;

let cached = null; // {token, expiresAt}

async function getAppGraphToken() {
  if (cached && cached.expiresAt > Date.now() + 60000) return cached.token;

  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials'
    })
  });
  if (!res.ok) {
    throw new Error(`Failed to acquire app Graph token: ${res.status} ${await res.text()}`);
  }
  const json = await res.json();
  cached = { token: json.access_token, expiresAt: Date.now() + json.expires_in * 1000 };
  return cached.token;
}

module.exports = { getAppGraphToken };
