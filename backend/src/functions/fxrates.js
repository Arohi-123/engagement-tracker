const { app } = require('@azure/functions');
const { authenticate } = require('../lib/authenticate');
const { GRAPH_BASE, graphGet } = require('../lib/graph');

// Same admin/FX-rates SharePoint site as the old ACCESS_CONTROL_SITE_ID in app.js.
// Finance maintains this list directly in SharePoint — this endpoint just reads it
// through the app's own access instead of the signed-in user's.
const FX_SITE_ID = 'neovationsg.sharepoint.com,2a6704aa-ebec-4ac3-9c90-bb26c7471a7a,9f770810-80a0-43e9-8516-566420d17813';
const FX_LIST_ID = '5e75b9bf-db6b-41bf-b732-30922f3b95de';

// GET /api/fxrates — no region scoping needed, any authenticated (valid-role) user
// can read currency rates; they're not sensitive business data.
app.http('fxRates', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'fxrates',
  handler: async (request, context) => {
    try {
      await authenticate(request);
    } catch (err) {
      return { status: err.status || 500, jsonBody: { error: err.message } };
    }

    try {
      const items = await graphGet(`${GRAPH_BASE}/sites/${FX_SITE_ID}/lists/${FX_LIST_ID}/items?expand=fields&$top=999`);
      const rates = { SGD: 1 };
      items.forEach(i => {
        const code = String((i.fields || {}).Title || '').trim().toUpperCase();
        const rate = Number((i.fields || {}).RateToSGD);
        if (code && !isNaN(rate)) rates[code] = rate;
      });
      return { status: 200, jsonBody: rates };
    } catch (err) {
      context.error(err);
      return { status: 502, jsonBody: { error: 'Failed to read FX rates' } };
    }
  }
});
