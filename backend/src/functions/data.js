const { app } = require('@azure/functions');
const { authenticate } = require('../lib/authenticate');
const { canReadRegion, canWriteRegion } = require('../lib/permissions');
const { REGIONS } = require('../lib/regions');
const { spItemToRecord, recordToSpFields } = require('../lib/fieldMap');
const { GRAPH_BASE, graphGet, graphPost, graphPatch } = require('../lib/graph');

const KINDS = ['clients', 'opportunities', 'engagements', 'companies'];

function errResponse(err) {
  return { status: err.status || 500, jsonBody: { error: err.message || 'Internal error' } };
}

// GET /api/data/{region} — all 4 lists for one region. Global/SuperAdmin roles can
// request any region; Regional roles only their own (enforced here, not trusted
// from the frontend).
app.http('getData', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'data/{region}',
  handler: async (request, context) => {
    let user;
    try { user = await authenticate(request); } catch (err) { return errResponse(err); }

    const regionKey = request.params.region;
    const region = REGIONS[regionKey];
    if (!region) return { status: 404, jsonBody: { error: `Unknown region '${regionKey}'` } };
    if (!canReadRegion(user, regionKey)) return { status: 403, jsonBody: { error: 'Not authorized for this region' } };

    try {
      const [c, o, e, co] = await Promise.all(KINDS.map(kind =>
        graphGet(`${GRAPH_BASE}/sites/${region.siteId}/lists/${region.listIds[kind]}/items?expand=fields&$top=500`)
      ));
      return {
        status: 200,
        jsonBody: {
          clients: c.map(i => spItemToRecord('clients', i)),
          opportunities: o.map(i => spItemToRecord('opportunities', i)),
          engagements: e.map(i => spItemToRecord('engagements', i)),
          companies: co.map(i => spItemToRecord('companies', i))
        }
      };
    } catch (err) {
      context.error(err);
      return { status: 502, jsonBody: { error: 'Failed to read from SharePoint' } };
    }
  }
});

// POST /api/data/{region}/{kind} — create a record. Viewer roles (Global or
// Regional) are blocked here regardless of what the frontend shows/hides.
app.http('createData', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'data/{region}/{kind}',
  handler: async (request, context) => {
    let user;
    try { user = await authenticate(request); } catch (err) { return errResponse(err); }

    const regionKey = request.params.region;
    const kind = request.params.kind;
    const region = REGIONS[regionKey];
    if (!region) return { status: 404, jsonBody: { error: `Unknown region '${regionKey}'` } };
    if (!KINDS.includes(kind)) return { status: 400, jsonBody: { error: `Unknown list '${kind}'` } };
    if (!canWriteRegion(user, regionKey)) return { status: 403, jsonBody: { error: 'Not authorized to write to this region' } };

    let payload;
    try { payload = await request.json(); } catch { return { status: 400, jsonBody: { error: 'Invalid JSON body' } }; }

    try {
      const fields = recordToSpFields(kind, payload);
      const created = await graphPost(`${GRAPH_BASE}/sites/${region.siteId}/lists/${region.listIds[kind]}/items`, { fields });
      return { status: 201, jsonBody: spItemToRecord(kind, created) };
    } catch (err) {
      context.error(err);
      return { status: 502, jsonBody: { error: 'Failed to save to SharePoint' } };
    }
  }
});

// PATCH /api/data/{region}/{kind}/{id} — update an existing record. Same write
// check as create.
app.http('updateData', {
  methods: ['PATCH'],
  authLevel: 'anonymous',
  route: 'data/{region}/{kind}/{id}',
  handler: async (request, context) => {
    let user;
    try { user = await authenticate(request); } catch (err) { return errResponse(err); }

    const regionKey = request.params.region;
    const kind = request.params.kind;
    const id = request.params.id;
    const region = REGIONS[regionKey];
    if (!region) return { status: 404, jsonBody: { error: `Unknown region '${regionKey}'` } };
    if (!KINDS.includes(kind)) return { status: 400, jsonBody: { error: `Unknown list '${kind}'` } };
    if (!canWriteRegion(user, regionKey)) return { status: 403, jsonBody: { error: 'Not authorized to write to this region' } };

    let payload;
    try { payload = await request.json(); } catch { return { status: 400, jsonBody: { error: 'Invalid JSON body' } }; }

    try {
      const fields = recordToSpFields(kind, payload);
      await graphPatch(`${GRAPH_BASE}/sites/${region.siteId}/lists/${region.listIds[kind]}/items/${id}/fields`, fields);
      return { status: 200, jsonBody: { success: true } };
    } catch (err) {
      context.error(err);
      return { status: 502, jsonBody: { error: 'Failed to update SharePoint' } };
    }
  }
});
