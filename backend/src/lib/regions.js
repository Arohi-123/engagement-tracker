// Mirrors REGIONS in the frontend's app.js — one SharePoint site + list set per
// business region. Keep in sync if a region is ever added/changed there.
const REGIONS = {
  ME: {
    label: 'ME',
    siteId: 'neovationsg.sharepoint.com,31e6c0b5-79a5-4ab7-9a3d-d81b111fd9fb,370ce563-c0f0-4788-8152-5dac7fb721e5',
    listIds: {
      clients: 'ea38938c-a1bf-45fa-b011-41e7ce2e35f9',
      opportunities: '8ffc2b36-fc95-47b1-9412-c2383a1d935d',
      engagements: 'bc18a9fa-155c-45ff-98b1-c87cedb823b7',
      companies: '6aa2ab20-ad4c-4dcd-b931-c2b9eeda4585'
    }
  },
  US: {
    label: 'US',
    siteId: 'neovationsg.sharepoint.com,ff5ea6ec-f50a-4876-88b0-a8a96bdc326a,9f770810-80a0-43e9-8516-566420d17813',
    listIds: {
      clients: '65a6fa76-1997-4806-b61e-353ca0e7ece6',
      opportunities: '102e339b-23fa-428f-86b9-a8d07d6d6cd8',
      engagements: 'eccc6ebd-011f-44d9-91dd-526f25f8c2c3',
      companies: 'e7333c74-284a-4d2d-91df-325e419cee7b'
    }
  },
  APAC: {
    label: 'APAC',
    siteId: 'neovationsg.sharepoint.com,c85dc6e1-6c8b-4109-8928-a896237ec950,9f770810-80a0-43e9-8516-566420d17813',
    listIds: {
      clients: '4d8fec0d-db60-40e4-9844-d2a99e5f2207',
      opportunities: '91c0881b-b20f-4c37-97eb-9003524cdb68',
      engagements: '984e595a-e2af-4464-aa43-b522e2b8feba',
      companies: 'ea9723fb-424d-4f60-853d-fcc8cc0a30ea'
    }
  },
  IND: {
    label: 'IND',
    siteId: 'neovationsg.sharepoint.com,e5171710-9380-46be-bfce-b2223c5093a4,9f770810-80a0-43e9-8516-566420d17813',
    listIds: {
      clients: '5fbe4e4d-1423-470b-a840-af898f880cd4',
      opportunities: '2659aa31-3c50-46cf-8d99-00ee6beac153',
      engagements: 'd700724d-ac3c-4429-bc54-8664efa54086',
      companies: '134d7fd0-9383-40df-9ed5-593f6122036f'
    }
  }
};

module.exports = { REGIONS };
