/* ============================================================
   Neovation Client Engagement Tracker v2
   ============================================================ */

/* ---------- 1. CONFIG ---------- */
const MSAL_CLIENT_ID   = '997ad661-b120-4f9e-bc6f-f672a0339206';
const MSAL_TENANT_ID   = 'e1a8eec3-eb6b-4aba-bbca-c2512937361f';
const GRAPH_SCOPES  = ['User.Read','Sites.ReadWrite.All'];

// One SharePoint site + list set per business region. To onboard a new region,
// create its SharePoint site (same 4 lists, same columns as the others) and add
// an entry here — no other code changes are required.
const REGIONS = {
  ME: {
    label: 'ME',
    siteId: 'neovationsg.sharepoint.com,31e6c0b5-79a5-4ab7-9a3d-d81b111fd9fb,370ce563-c0f0-4788-8152-5dac7fb721e5',
    listIds: {
      clients:       'ea38938c-a1bf-45fa-b011-41e7ce2e35f9',
      opportunities: '8ffc2b36-fc95-47b1-9412-c2383a1d935d',
      engagements:   'bc18a9fa-155c-45ff-98b1-c87cedb823b7',
      companies:     '6aa2ab20-ad4c-4dcd-b931-c2b9eeda4585'
    }
  },
  US: {
    label: 'US',
    siteId: 'neovationsg.sharepoint.com,ff5ea6ec-f50a-4876-88b0-a8a96bdc326a,9f770810-80a0-43e9-8516-566420d17813',
    listIds: {
      clients:       '65a6fa76-1997-4806-b61e-353ca0e7ece6',
      opportunities: '102e339b-23fa-428f-86b9-a8d07d6d6cd8',
      engagements:   'eccc6ebd-011f-44d9-91dd-526f25f8c2c3',
      companies:     'e7333c74-284a-4d2d-91df-325e419cee7b'
    }
  },
  APAC: {
    label: 'APAC',
    siteId: 'neovationsg.sharepoint.com,c85dc6e1-6c8b-4109-8928-a896237ec950,9f770810-80a0-43e9-8516-566420d17813',
    listIds: {
      clients:       '4d8fec0d-db60-40e4-9844-d2a99e5f2207',
      opportunities: '91c0881b-b20f-4c37-97eb-9003524cdb68',
      engagements:   '984e595a-e2af-4464-aa43-b522e2b8feba',
      companies:     'ea9723fb-424d-4f60-853d-fcc8cc0a30ea'
    }
  },
  IND: {
    label: 'IND',
    siteId: 'neovationsg.sharepoint.com,e5171710-9380-46be-bfce-b2223c5093a4,9f770810-80a0-43e9-8516-566420d17813',
    listIds: {
      clients:       '5fbe4e4d-1423-470b-a840-af898f880cd4',
      opportunities: '2659aa31-3c50-46cf-8d99-00ee6beac153',
      engagements:   'd700724d-ac3c-4429-bc54-8664efa54086',
      companies:     '134d7fd0-9383-40df-9ed5-593f6122036f'
    }
  }
};

// Small dedicated SharePoint site/list holding Email / Role / Region rows.
// This is the single source of truth for who can access which region and as what role —
// add or change a row here to grant/revoke/change access, no code change or redeploy needed.
const ACCESS_CONTROL_SITE_ID = 'neovationsg.sharepoint.com,2a6704aa-ebec-4ac3-9c90-bb26c7471a7a,9f770810-80a0-43e9-8516-566420d17813';
const ACCESS_CONTROL_LIST_ID = '9d2f755c-3ebd-49ad-bd07-1ae7cb563606';

/* ---------- 2. LOOKUP LISTS ---------- */
const EMPLOYEES = [
  'Aditya Ghuman','Aishwarya Sutar','Aashka Kothari','Arohi Jain','Arpana Akella',
  'Arundhati Tiwari','Ashwini Maratha','Brenda Cheong','Burhan Patan','Chuin Hau Teo',
  'Disha Dayal','Dnyanesh Limaye','Fariza Nasreen','Hasnain Patel','Ibrahim Hamdy',
  'Ilmira Murni','Jaya Pandey','Kamayani Gupta','Kapil Vats','Kashmira Mangavkar',
  'Keegan Afonso','Keerthy Thatikonda','Linu Thomas','Manila Pasricha','Murtaza Bhugediwala',
  'Navneet Kaur','Neha Patnaik','Nehal Hassan','Padmashree Kanangi','Pargat Singh',
  'Pawan Kumar','Pooja Amuthan','Priya Joshi','Priyanka Kapoor','Sakshi Sharma',
  'Salini Asok','Saloni Kakkar','Sandhya Kaimal','Sanjay Joseph','Santosh Jha',
  'Santhosh','Sharon Mueller','Siddhi Parab','Simpy Sethi Khanna','Smrita Chaudhury',
  'Snehal Ingale','Sonal Nafade','Susan Linda Rajkumar','Tanzim Khan','Tanvi Sejpal',
  'Vaishnavi','Vibhuti Rana',
].sort();

const LOOKUPS = {
  therapyAreas:['Cardiovascular/ Cardiometabolic','Consumer healthcare','Dermatology','Endocrinology',
    'Fertility','Gastroenterology','Hematology','Immunology','Infectious Disease','Medical devices',
    'Nephrology','Neurology/ CNS','Obesity','Oncology','Ophthalmology','Rare Diseases','Respiratory',
    'Rheumatology','Vaccines and Diagnostics'],
  departments:['Medical','Market Access/ HEOR','Marketing/ Commercial','Procurement','Digital','Others'],
  regionGroupings:['Gulf','MENA','Intercontinental','MEAR','GCC','META','GEM & Levant','Others'],
  countries:['UAE','KSA','Qatar','Kuwait','Bahrain','Oman','Egypt','Jordan','Lebanon'],
  clientStatus:['Active business','Active engagement','Prospect','Inactive'],
  priority:['Low','Medium','High'],
  opportunityStatus:[
    '01 Blue Sky Opportunity','02 Identified Opportunity',
    '04 Proposal preparation','05 Proposal submitted','06 To Win',
    '07 Win','08 Loss','09 Client not interested','10 Missed Opportunity'
  ],
  engagementTypes:['In-Person Meeting','Virtual Meeting','Phone Call/ Whatsapp','Email','Conference/Event','Others'],
  stakeholderType:['Existing Stakeholder Engagement','Procurement Engagement','New Stakeholder Development'],
  engagementObjective:['Blue sky opportunity discussion','Identified opportunity discussion','Discussion on new RFP',
    'Follow up on ongoing RFP','Capability presentation','Request for referral','Intelligence gathering',
    'Relationship building','Other discussion'],
  engagementOutcome:['New opportunities discovered','Led to RFP','Progress on ongoing RFP','Led to referral','Others (Mentioned in notes)'],
  followUp:['Yes','No'],
  onboardingStatus:['Active MSA','Onboarded','Preferred Vendor'],
  clientPerception:['Vendor','Partner','Advisor']
};

/* ---------- 3. STATE ---------- */
let msalInstance = null;
let currentAccount = null;
let currentUser = null; // {username,name,role:'admin'|'regional_admin'|'regional_viewer',region:string|null}
let activeRegionKey = null; // key into REGIONS for whichever region's data is currently loaded
function activeRegion(){ return REGIONS[activeRegionKey]; }
const ROLE_LABELS = {admin:'ADMIN', regional_admin:'REGIONAL ADMIN', regional_viewer:'VIEWER'};
let DATA = { clients:[], opportunities:[], engagements:[], companies:[] };
let CHARTS = {};
let SORT_STATE = {}; // { viewKey: { col, dir } }
let bdStakeView='total'; // 'total' | 'active' — which Stakeholder List row is visible
let bdStakeExpanded=false; // collapsed by default: only the Total Clients tile shows until expanded
let bdEngExpanded=false; // collapsed by default: only the Total tile shows until expanded
let bdOppExpanded=false;
function toggleBdStakeExpanded(){ bdStakeExpanded=!bdStakeExpanded; renderBDFunnel(); }
function toggleBdEngExpanded(){ bdEngExpanded=!bdEngExpanded; renderBDFunnel(); }
function toggleBdOppExpanded(){ bdOppExpanded=!bdOppExpanded; renderBDFunnel(); }

// All-Regions Overview: same idea as the single-region BD funnel above, but its own state
// and data store (ALL_REGIONS_DATA) so switching in/out of it never touches the normal
// single-region DATA or its toggle states.
let ALL_REGIONS_DATA=null; // {merged:{clients,opportunities,engagements,companies}, byRegion:[{key,label,clients,opportunities,engagements,companies}]}
let arPeriod='ytd'; // weekly | monthly | quarterly | ytd — scopes engagement-activity figures only; pipeline/win-rate stay all-time, same convention as the rest of the app
let arStakeView='total';
let arStakeExpanded=false;
let arEngExpanded=false;
let arOppExpanded=false;
function toggleArBdStakeExpanded(){ arStakeExpanded=!arStakeExpanded; renderAllRegionsBDFunnel(); }
function toggleArBdEngExpanded(){ arEngExpanded=!arEngExpanded; renderAllRegionsBDFunnel(); }
function toggleArBdOppExpanded(){ arOppExpanded=!arOppExpanded; renderAllRegionsBDFunnel(); }

/* ---------- 4. UTILITIES ---------- */
function esc(str){
  if(str==null) return '';
  return String(str).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}
function fmtDate(d){
  if(!d) return '';
  const dt=new Date(d); if(isNaN(dt)) return '';
  return dt.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'});
}
// All monetary fields in SharePoint are stored in USD. displayCurrency only changes how
// fmtMoney() *shows* a number — every KPI/table/chart in the app already routes through
// this one function, so that's the only place a currency toggle needs to touch.
const CURRENCY_SYMBOLS={USD:'$',SGD:'S$'};
const FALLBACK_USD_SGD=1.35; // used only if the live rate fetch fails/hasn't completed yet
// localStorage can throw (not just be undefined) under some browser security contexts —
// private browsing, strict file:// origin policies, storage disabled by IT policy — and
// since this runs at the very top of the script, an uncaught throw here would silently
// kill every function defined after it. Never touch localStorage without this wrapper.
function safeLsGet(key){ try{ return localStorage.getItem(key); }catch(e){ return null; } }
function safeLsSet(key,val){ try{ localStorage.setItem(key,val); }catch(e){/* ignore */} }
let displayCurrency=safeLsGet('displayCurrency')||'USD';
let fxUsdSgd=null;
let fxFetchedAt=null;
async function ensureFxRate(){
  if(fxUsdSgd&&fxFetchedAt&&(Date.now()-fxFetchedAt<12*3600*1000)) return; // already fresh
  try{
    const cachedRaw=safeLsGet('fxUsdSgd');
    if(cachedRaw){
      const cached=JSON.parse(cachedRaw);
      if(cached.rate&&Date.now()-cached.at<12*3600*1000){ fxUsdSgd=cached.rate; fxFetchedAt=cached.at; return; }
    }
  }catch(e){/* ignore malformed cache */}
  try{
    const res=await fetch('https://api.frankfurter.app/latest?from=USD&to=SGD');
    const j=await res.json();
    if(j&&j.rates&&j.rates.SGD){
      fxUsdSgd=j.rates.SGD; fxFetchedAt=Date.now();
      safeLsSet('fxUsdSgd',JSON.stringify({rate:fxUsdSgd,at:fxFetchedAt}));
    }
  }catch(err){
    console.error('FX rate fetch failed, using fallback rate:',err);
  }
}
function setDisplayCurrency(cur){
  displayCurrency=cur;
  safeLsSet('displayCurrency',cur);
  syncCurrencySelects();
  if(document.getElementById('app')&&!document.getElementById('app').hidden) renderAllViews();
  const arScreen=document.getElementById('all-regions-screen');
  if(arScreen&&arScreen.style.display!=='none'&&ALL_REGIONS_DATA) renderAllRegionsOverview();
}
function syncCurrencySelects(){
  ['currency-select','ar-currency-select'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.value=displayCurrency;
  });
}
function fmtMoney(n){
  if(n==null||n==='') return '';
  const num=Number(n); if(isNaN(num)) return '';
  const rate=displayCurrency==='SGD'?(fxUsdSgd||FALLBACK_USD_SGD):1;
  const converted=num*rate;
  const sym=CURRENCY_SYMBOLS[displayCurrency]||'$';
  if(converted>=1000000) return sym+(converted/1000000).toFixed(1)+'M';
  if(converted>=1000) return sym+(converted/1000).toFixed(0)+'K';
  return sym+converted.toLocaleString('en-US',{maximumFractionDigits:0});
}
function fmtNum(n){ if(n==null||n==='') return ''; return Number(n).toLocaleString('en-US'); }
function debounce(fn,ms){ let t; return (...a)=>{clearTimeout(t);t=setTimeout(()=>fn(...a),ms);}; }
function isoWeek(dateStr){
  if(!dateStr) return null;
  const d=new Date(dateStr); if(isNaN(d)) return null;
  const dt=new Date(Date.UTC(d.getFullYear(),d.getMonth(),d.getDate()));
  const day=(dt.getUTCDay()+6)%7; dt.setUTCDate(dt.getUTCDate()-day+3);
  const jan4=new Date(Date.UTC(dt.getUTCFullYear(),0,4));
  const wk=1+Math.round(((dt-jan4)/86400000-3+((jan4.getUTCDay()+6)%7))/7);
  return `${dt.getUTCFullYear()}-W${String(wk).padStart(2,'0')}`;
}
function weekRangeLabel(d){
  const dt=new Date(d); if(isNaN(dt)) return '';
  const day=(dt.getDay()+6)%7;
  const mon=new Date(dt); mon.setDate(dt.getDate()-day);
  const sun=new Date(mon); sun.setDate(mon.getDate()+6);
  const o={day:'numeric',month:'short'};
  return `${mon.toLocaleDateString('en-GB',o)} – ${sun.toLocaleDateString('en-GB',o)}`;
}
function mondayOf(dateStr){
  const d=new Date(dateStr); if(isNaN(d)) return null;
  const day=(d.getDay()+6)%7;
  return new Date(d.getFullYear(),d.getMonth(),d.getDate()-day);
}
// Buckets a list of date strings into ISO weeks, returning them oldest→newest with
// human-readable, date-based labels (never raw "YYYY-Www" strings) for charts/KPIs.
function weeklySeries(dateList,count){
  const map={};
  (dateList||[]).forEach(ds=>{
    if(!ds) return;
    const wk=isoWeek(ds); if(!wk) return;
    if(!map[wk]) map[wk]={count:0,monday:mondayOf(ds)};
    map[wk].count++;
  });
  let keys=Object.keys(map).sort();
  if(count) keys=keys.slice(-count);
  return keys.map(k=>{
    const m=map[k];
    return {
      key:k,count:m.count,monday:m.monday,
      label:m.monday?m.monday.toLocaleDateString('en-GB',{day:'2-digit',month:'short'}):k,
      rangeLabel:m.monday?weekRangeLabel(m.monday):k
    };
  });
}
function initials(n){ return (n||'?').split(' ').map(p=>p[0]).join('').slice(0,2).toUpperCase(); }
function daysAgo(dateStr){
  if(!dateStr) return null;
  return Math.floor((Date.now()-new Date(dateStr))/(1000*60*60*24));
}
function sortAlpha(arr,key){ return [...arr].sort((a,b)=>(a[key]||'').localeCompare(b[key]||'')); }
function uniqueCompanies(){
  const s=new Set();
  DATA.companies.forEach(c=>{if(c.company)s.add(c.company.trim());});
  return [...s].sort();
}
// Department/Therapy area comboboxes allow typing a new value via "Add new…", but that
// only ever set the value on the record being saved — it never made it back into the
// static LOOKUPS list, so it silently vanished from the suggestions on every later entry.
// Merge in whatever's actually been used across the real client data so a custom value,
// once saved anywhere, becomes a normal pickable suggestion from then on.
function uniqueDepartments(){
  const s=new Set(LOOKUPS.departments);
  DATA.clients.forEach(c=>{if(c.department)s.add(c.department.trim());});
  return [...s].sort();
}
function uniqueTherapyAreas(){
  const s=new Set(LOOKUPS.therapyAreas);
  DATA.clients.forEach(c=>{if(c.therapy_area)s.add(c.therapy_area.trim());});
  return [...s].sort();
}

/* ---------- 5. COLOUR HELPERS ---------- */
const STAGE_NUM = s => { const m=(s||'').match(/^(\d+)/); return m?Number(m[1]):0; };
function stageTagClass(status){
  const n=STAGE_NUM(status);
  return n>=1&&n<=10 ? `tag stage-0${n<10?'0'+n:n}`.replace('stage-010','stage-10') : 'tag tag-grey';
}
function stageTagClassN(n){ return n>=1&&n<=9?`tag stage-0${n}`:(n===10?'tag stage-10':'tag tag-grey'); }
function clientStatusClass(s){
  const m={'active business':'tag-green','active engagement':'tag-sky','prospect':'tag-amber','inactive':'tag-coral'};
  return 'tag '+(m[(s||'').toLowerCase()]||'tag-grey');
}
function priorityClass(p){
  return 'tag '+(p==='High'?'tag-coral':p==='Medium'?'tag-amber':'tag-grey');
}
function engTypeClass(t){
  const m={'in-person meeting':'tag-green','virtual meeting':'tag-sky','phone call/ whatsapp':'tag-amber',
    'email':'tag-purple','conference/event':'tag-teal','others':'tag-grey'};
  return 'tag '+(m[(t||'').toLowerCase()]||'tag-grey');
}
function followUpClass(v){
  return 'tag '+(String(v||'').toLowerCase()==='yes'?'tag-green':'tag-amber');
}
function onboardingClass(s){
  const m={'onboarded':'tag-green','active msa':'tag-sky','preferred vendor':'tag-teal'};
  return 'tag '+(m[(s||'').toLowerCase()]||'tag-grey');
}
// Stage colors for Chart.js
// Value-only palette update (UI/UX redesign) — colors extracted from Neovation's brand
// deck and validated for categorical distinguishability (dataviz skill's validate_palette.js:
// lightness band, chroma floor, CVD-adjacency, contrast all pass). No logic changed below.
const STAGE_COLORS=['#E3F1FA','#1F5A9C','#5FA8E8','#08588A','#00A6C0','#3E8E4F','#C9A227','#1E6423','#E8836A','#95447F','#3A7DBA'];
const COLORS={ink:'#0A1E3D',sky:'#0070C0',skyDeep:'#0A4A85',skyPale:'#E3F1FA',amber:'#D98E2B',amberPale:'#FBF0DD',green:'#2E8B57',greenPale:'#E4F3E1',coral:'#D9634A',coralPale:'#FBE7E1',inkSoft:'#4B5C74',line:'#DEE2EA'};

function destroyChart(k){ if(CHARTS[k]){CHARTS[k].destroy();delete CHARTS[k];} }

/* ---------- 5b. PIVOT TABLE HELPER (shared by the 3 dashboard pivots) ---------- */
// Toggles a group row's child rows open/closed in-place (no re-render needed).
function togglePivotGroup(groupRow){
  groupRow.classList.toggle('open');
  const isOpen=groupRow.classList.contains('open');
  let sib=groupRow.nextElementSibling;
  while(sib&&sib.classList.contains('pivot-child')){
    sib.classList.toggle('show',isOpen);
    sib=sib.nextElementSibling;
  }
}

/* ---------- 6. STAGE LOGIC ---------- */
function isOpenStage(s){
  const closed=['07 win','08 loss','09 client not interested','10 missed opportunity'];
  return s && !closed.includes(String(s).toLowerCase().trim());
}
function isWinStage(s){ return String(s||'').toLowerCase().trim()==='07 win'; }
function isLossStage(s){ const v=String(s||'').toLowerCase().trim(); return v==='08 loss'||v==='09 client not interested'; }
function autoProb(status){
  if(isWinStage(status)) return 1;
  if(isLossStage(status)) return 0;
  return null; // no auto-set
}

/* ---------- 7. AUTH (MSAL) ---------- */
function initMsal(){
  if(msalInstance) return;
  msalInstance=new msal.PublicClientApplication({
    auth:{clientId:MSAL_CLIENT_ID,authority:`https://login.microsoftonline.com/${MSAL_TENANT_ID}`,redirectUri:window.location.origin+window.location.pathname},
    cache:{cacheLocation:'memoryStorage'}
  });
}
async function doLogin(){
  const errEl=document.getElementById('login-error');
  errEl.style.display='none';
  if(MSAL_CLIENT_ID.includes('YOUR_')||MSAL_TENANT_ID.includes('YOUR_')||ACCESS_CONTROL_SITE_ID.includes('YOUR_')||ACCESS_CONTROL_LIST_ID.includes('YOUR_')){
    errEl.textContent='Not connected to Microsoft 365. See SHAREPOINT_SETUP_GUIDE.md.';
    errEl.style.display='block'; return;
  }
  try{
    initMsal();
    await msalInstance.initialize();
    const result=await msalInstance.loginPopup({scopes:GRAPH_SCOPES});
    currentAccount=result.account;
    msalInstance.setActiveAccount(currentAccount);
    const email=(currentAccount.username||'').toLowerCase();
    const access=await fetchAccessControl(email);
    if(!access||!access.role){
      showAccessDeniedScreen('No access configured for this account. Contact an admin to be added to the Access Control list.');
      return;
    }
    currentUser={username:email,name:currentAccount.name||email,role:access.role,region:access.region||null};
    document.getElementById('login-screen').style.display='none';
    document.getElementById('user-name').textContent=currentUser.name;
    document.getElementById('user-role').textContent=ROLE_LABELS[currentUser.role]||currentUser.role.toUpperCase();
    document.getElementById('user-avatar').textContent=initials(currentUser.name);
    applyRolePermissions();
    if(currentUser.role==='admin'){
      showRegionScreen();
    }else{
      if(!REGIONS[currentUser.region]){
        showAccessDeniedScreen('Your assigned region is not configured yet. Contact an admin.');
        return;
      }
      await enterRegion(currentUser.region);
    }
  }catch(err){
    console.error(err);
    errEl.textContent='Sign-in failed: '+(err.errorMessage||err.message||err);
    errEl.style.display='block';
  }
}
function doLogout(){
  currentUser=null; currentAccount=null; activeRegionKey=null;
  document.getElementById('app').hidden=true;
  document.getElementById('region-screen').style.display='none';
  const arScreen=document.getElementById('all-regions-screen');
  if(arScreen) arScreen.style.display='none';
  document.getElementById('access-denied-screen').style.display='none';
  document.getElementById('login-screen').style.display='flex';
  document.getElementById('brand-mark').style.display='block';
  if(msalInstance&&msalInstance.getActiveAccount()) msalInstance.logoutPopup().catch(()=>{});
}
function showAccessDeniedScreen(msg){
  document.getElementById('login-screen').style.display='none';
  document.getElementById('region-screen').style.display='none';
  document.getElementById('app').hidden=true;
  document.getElementById('access-denied-msg').textContent=msg;
  document.getElementById('access-denied-screen').style.display='flex';
  document.getElementById('brand-mark').style.display='block';
}
function showRegionScreen(){
  document.getElementById('app').hidden=true;
  const arScreen=document.getElementById('all-regions-screen');
  if(arScreen) arScreen.style.display='none';
  document.getElementById('brand-mark').style.display='block';
  const grid=document.getElementById('region-screen-grid');
  grid.innerHTML=Object.entries(REGIONS).map(([key,r])=>
    `<button class="region-pick-card" onclick="enterRegion('${esc(key)}')">
       <div class="region-pick-label">${esc(r.label)}</div>
       <div class="region-pick-sub">Open this region's dashboard</div>
     </button>`
  ).join('')+
  (currentUser&&currentUser.role==='admin'?
  `<button class="region-pick-card" onclick="enterAllRegions()">
     <div class="region-pick-label">🌐 All Regions</div>
     <div class="region-pick-sub">Combined read-only view across every region</div>
   </button>`:'');
  document.getElementById('region-screen').style.display='flex';
}
async function enterRegion(key){
  activeRegionKey=key;
  document.getElementById('region-screen').style.display='none';
  document.getElementById('app').hidden=false;
  document.getElementById('brand-mark').style.display='none';
  document.getElementById('region-indicator').textContent=activeRegion().label;
  document.getElementById('region-switch-btn').style.display=(currentUser.role==='admin')?'inline-flex':'none';
  await boot();
}
async function getGraphToken(){
  initMsal();
  try{ return (await msalInstance.acquireTokenSilent({scopes:GRAPH_SCOPES,account:currentAccount})).accessToken; }
  catch(e){ return (await msalInstance.acquireTokenPopup({scopes:GRAPH_SCOPES})).accessToken; }
}
function canEdit(){ return !!currentUser && (currentUser.role==='admin' || currentUser.role==='regional_admin'); }
function applyRolePermissions(){
  const v=!canEdit();
  ['add-client-btn','add-opp-btn','add-eng-btn','add-company-btn'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.style.display=v?'none':'inline-flex';
  });
}

/* ---------- 8. NAV ---------- */
function switchView(v){
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.toggle('active',b.dataset.view===v));
  document.querySelectorAll('.view').forEach(s=>s.classList.toggle('active',s.id==='view-'+v));
  document.querySelectorAll('.sidebar').forEach(s=>s.classList.remove('open'));
}

/* ---------- 9. DATA LAYER ---------- */
const GRAPH_BASE='https://graph.microsoft.com/v1.0';
const FIELD_MAP={
  clients:{client_name:'Title',company:'field_1',designation:'field_2',department:'field_3',therapy_area:'field_4',region:'field_5',email:'field_6',phone:'field_7',linkedin_url:'field_8',status:'field_9',priority:'field_10',assigned_bd:'field_11',notes:'field_12'},
  opportunities:{opportunity:'Title',rfp_id:'field_1',company:'field_2',client_name:'field_3',opportunity_status:'field_4',discussion_date:'field_5',pitch_date:'field_6',proposal_submission_date:'field_7',expected_close_date:'field_8',bd_owner:'field_9',supporting_role:'field_10',estimated_value_usd:'field_11',probability_pct:'field_12',probability_weighted_value:'field_13',notes:'field_14',stage_history:'StageHistory',identified_date:'IdentifiedDate'},
  engagements:{client_name:'Title',eng_month:'field_1',eng_date:'field_2',designation:'field_3',company:'field_4',bd_pm:'field_5',engagement_type:'field_6',stakeholder_type:'field_8',engagement_objective:'field_9',engagement_outcome:'field_10',discussion_points:'field_11',cta_next_step:'field_12',cta_due_date:'field_13',cta_owner:'field_14',follow_up_done:'field_15',accompanied_by:'AccompaniedBy'},
  companies:{company:'Title',onboarding_status:'field_1',notes:'field_2',
    target_revenue:'TargetRevenue',overall_budget_potential:'BudgetPotential',overall_client_relationship:'ClientRelationship',
    client_perception:'ClientPerception',team_satisfaction:'TeamSatisfaction',degree_of_innovation:'DegreeOfInnovation'}
};
const DATE_FIELDS={clients:[],opportunities:['discussion_date','pitch_date','proposal_submission_date','expected_close_date','identified_date'],engagements:['eng_month','eng_date','cta_due_date'],companies:[]};

// Different regions/people have typed the same company inconsistently in SharePoint
// (e.g. "Abbvie" vs "AbbVie"). Canonicalize on read so every view — dropdowns, pivots,
// rollups — treats known variants as a single company.
const COMPANY_ALIASES={'abbvie':'AbbVie'};
function canonicalCompany(name){
  const t=String(name||'').trim();
  return COMPANY_ALIASES[t.toLowerCase()]||t;
}

function spItemToRecord(kind,item){
  const map=FIELD_MAP[kind],f=item.fields||{},rec={id:item.id};
  Object.entries(map).forEach(([k,sp])=>{
    let v=f[sp];
    if(v!=null&&DATE_FIELDS[kind].includes(k)) v=String(v).split('T')[0];
    rec[k]=v??null;
  });
  if('company' in rec) rec.company=canonicalCompany(rec.company);
  return rec;
}
function recordToSpFields(kind,payload){
  const map=FIELD_MAP[kind],fields={};
  Object.entries(payload).forEach(([k,v])=>{
    const sp=map[k]; if(!sp) return;
    if(v!=null&&DATE_FIELDS[kind].includes(k)) v=new Date(v).toISOString();
    fields[sp]=v;
  });
  return fields;
}

async function graphGet(url){
  const token=await getGraphToken();
  let items=[],next=url;
  while(next){
    const r=await fetch(next,{headers:{Authorization:`Bearer ${token}`}});
    if(!r.ok){const b=await r.text();throw new Error(`Graph GET ${r.status}: ${b}`);}
    const j=await r.json(); items=items.concat(j.value||[]); next=j['@odata.nextLink']||null;
  }
  return items;
}
async function graphPost(url,body){
  const token=await getGraphToken();
  const r=await fetch(url,{method:'POST',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify(body)});
  if(!r.ok){const b=await r.text();throw new Error(`Graph POST ${r.status}: ${b}`);}
  return r.json();
}
async function graphPatch(url,body){
  const token=await getGraphToken();
  const r=await fetch(url,{method:'PATCH',headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},body:JSON.stringify(body)});
  if(!r.ok){const b=await r.text();throw new Error(`Graph PATCH ${r.status}: ${b}`);}
  return r.json();
}

// Reads the Access-Control list (Email/Role/Region columns) and returns this user's
// {role,region}, or null if they have no row. This is the single source of truth for
// who can sign in as what — managed entirely in SharePoint, never in code.
async function fetchAccessControl(email){
  const items=await graphGet(`${GRAPH_BASE}/sites/${ACCESS_CONTROL_SITE_ID}/lists/${ACCESS_CONTROL_LIST_ID}/items?expand=fields&$top=999`);
  const row=items.find(i=>String(i.fields.Email||'').toLowerCase()===email);
  if(!row) return null;
  return {
    role: String(row.fields.Role||'').toLowerCase().replace(/\s+/g,'_'),
    region: row.fields.Region||null
  };
}

async function fetchAll(){
  const r=activeRegion();
  const [c,o,e,co]=await Promise.all([
    graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.clients}/items?expand=fields&$top=500`),
    graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.opportunities}/items?expand=fields&$top=500`),
    graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.engagements}/items?expand=fields&$top=500`),
    graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.companies}/items?expand=fields&$top=500`)
  ]);
  DATA.clients=c.map(i=>spItemToRecord('clients',i));
  DATA.opportunities=o.map(i=>spItemToRecord('opportunities',i));
  DATA.engagements=e.map(i=>spItemToRecord('engagements',i));
  DATA.companies=co.map(i=>spItemToRecord('companies',i));
}

async function insertRecord(kind,payload){
  try{
    const r=activeRegion();
    const fields=recordToSpFields(kind,payload);
    const j=await graphPost(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds[kind]}/items`,{fields});
    return spItemToRecord(kind,j);
  }catch(err){alert('Could not save: '+err.message);return null;}
}
async function updateRecord(kind,id,payload){
  try{
    const r=activeRegion();
    const fields=recordToSpFields(kind,payload);
    await graphPatch(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds[kind]}/items/${id}/fields`,fields);
    return true;
  }catch(err){alert('Could not update: '+err.message);return false;}
}

function renderAllViews(){
  renderOverview();renderClients();renderOpportunities();renderEngagements();renderCompanies();
}
async function refreshAndRenderAll(){
  setRefreshSpinning(true);
  await fetchAll();
  renderAllViews();
  setRefreshSpinning(false);
}
function setRefreshSpinning(on){
  const btn=document.getElementById('refresh-btn');
  if(btn) btn.classList.toggle('spinning',on);
}

/* ---------- 10. COLUMN SORT HELPER ---------- */
function makeSortableHeaders(theadEl, cols, viewKey, renderFn){
  const state=SORT_STATE[viewKey]||{col:null,dir:'asc'};
  theadEl.innerHTML=`<tr>${cols.map(c=>`
    <th data-col="${esc(c.key||c.label)}" class="${state.col===(c.key||c.label)?('sort-'+state.dir):''}">
      ${esc(c.label)}<span class="sort-icon"></span>
    </th>`).join('')}</tr>`;
  theadEl.querySelectorAll('th').forEach(th=>{
    th.addEventListener('click',()=>{
      const col=th.dataset.col;
      const cur=SORT_STATE[viewKey]||{col:null,dir:'asc'};
      SORT_STATE[viewKey]={col,dir:(cur.col===col&&cur.dir==='asc')?'desc':'asc'};
      renderFn();
    });
  });
}
function applySort(rows,viewKey,defaultKey){
  const {col,dir}=SORT_STATE[viewKey]||{col:defaultKey,dir:'asc'};
  if(!col) return rows;
  return [...rows].sort((a,b)=>{
    let av=a[col]??'', bv=b[col]??'';
    if(!isNaN(Number(av))&&!isNaN(Number(bv))){av=Number(av);bv=Number(bv);}
    const cmp = typeof av==='number'?av-bv:String(av).localeCompare(String(bv));
    return dir==='asc'?cmp:-cmp;
  });
}

/* ---------- 11. KPI CARDS ---------- */
function renderKpiCards(id,cards){
  document.getElementById(id).innerHTML=cards.map(c=>`
    <div class="kpi-card accent-${c.accent||'sky'}">
      <div class="kpi-label">${esc(c.label)}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-sub">${esc(c.sub||'')}</div>
    </div>`).join('');
}

/* ---------- 12. RENDER: OVERVIEW ---------- */
// Fiscal year runs April–March (matches the app's "FY 2026–27" branding and the
// workbook's own "Q1' April – June" quarter labels).
let ovPeriod='ytd';
function periodStart(period){
  const now=new Date();
  if(period==='weekly'){
    const day=(now.getDay()+6)%7;
    return new Date(now.getFullYear(),now.getMonth(),now.getDate()-day);
  }
  if(period==='monthly') return new Date(now.getFullYear(),now.getMonth(),1);
  if(period==='quarterly'){
    const m=now.getMonth();
    const blockStartMonth=(m>=3&&m<=5)?3:(m>=6&&m<=8)?6:(m>=9&&m<=11)?9:0;
    return new Date(now.getFullYear(),blockStartMonth,1);
  }
  const fy=now.getMonth()>=3?now.getFullYear():now.getFullYear()-1; // YTD
  return new Date(fy,3,1);
}
function lookupClientDepartment(company,clientName){
  const co=normCo(company),cl=(clientName||'').trim().toLowerCase();
  if(!co&&!cl) return null;
  const match=DATA.clients.find(c=>normCo(c.company)===co&&(c.client_name||'').trim().toLowerCase()===cl);
  return match?match.department:null;
}
function renderOverviewBDTables(){
  const start=periodStart(ovPeriod);
  const engs=DATA.engagements.filter(e=>e.eng_date&&new Date(e.eng_date)>=start);
  const bdNames=[...new Set(engs.map(e=>normStr(e.bd_pm)).filter(Boolean))].sort();

  const rangeText=`${fmtDate(start)} – ${fmtDate(new Date())}`;
  const rangeEl1=document.getElementById('ov-period-range'); if(rangeEl1) rangeEl1.textContent=rangeText;
  const rangeEl2=document.getElementById('ov-period-range-2'); if(rangeEl2) rangeEl2.textContent=rangeText;

  const types=LOOKUPS.engagementTypes;
  const tbl1=document.getElementById('ov-bd-type-table');
  tbl1.querySelector('thead').innerHTML=`<tr><th>BD Owner</th>${types.map(t=>`<th class="num">${esc(t)}</th>`).join('')}<th class="num">Total</th></tr>`;
  if(!bdNames.length){
    tbl1.querySelector('tbody').innerHTML=`<tr><td colspan="${types.length+2}" class="empty-state">No engagements in this period.</td></tr>`;
  }else{
    const colTotals=types.map(()=>0);let grandTotal=0;
    const rows=bdNames.map(bd=>{
      const list=engs.filter(e=>normStr(e.bd_pm)===bd);
      const counts=types.map((t,i)=>{const n=list.filter(e=>e.engagement_type===t).length;colTotals[i]+=n;return n;});
      const total=counts.reduce((a,b)=>a+b,0);grandTotal+=total;
      return `<tr><td>${esc(bd)}</td>${counts.map(n=>`<td class="num">${n||''}</td>`).join('')}<td class="num">${total}</td></tr>`;
    }).join('');
    tbl1.querySelector('tbody').innerHTML=rows+`<tr class="pivot-total"><td>Grand Total</td>${colTotals.map(n=>`<td class="num">${n}</td>`).join('')}<td class="num">${grandTotal}</td></tr>`;
  }

  const deptCols=[...LOOKUPS.departments,'Unknown'];
  const tbl2=document.getElementById('ov-bd-dept-table');
  tbl2.querySelector('thead').innerHTML=`<tr><th>BD Owner</th>${deptCols.map(d=>`<th class="num">${esc(d)}</th>`).join('')}<th class="num">Total</th></tr>`;
  if(!bdNames.length){
    tbl2.querySelector('tbody').innerHTML=`<tr><td colspan="${deptCols.length+2}" class="empty-state">No engagements in this period.</td></tr>`;
  }else{
    const colTotals2=deptCols.map(()=>0);let grandTotal2=0;
    const rows2=bdNames.map(bd=>{
      const list=engs.filter(e=>normStr(e.bd_pm)===bd);
      const counts=deptCols.map((d,i)=>{
        const n=list.filter(e=>{const dep=lookupClientDepartment(e.company,e.client_name);return d==='Unknown'?!dep:dep===d;}).length;
        colTotals2[i]+=n;return n;
      });
      const total=counts.reduce((a,b)=>a+b,0);grandTotal2+=total;
      return `<tr><td>${esc(bd)}</td>${counts.map(n=>`<td class="num">${n||''}</td>`).join('')}<td class="num">${total}</td></tr>`;
    }).join('');
    tbl2.querySelector('tbody').innerHTML=rows2+`<tr class="pivot-total"><td>Grand Total</td>${colTotals2.map(n=>`<td class="num">${n}</td>`).join('')}<td class="num">${grandTotal2}</td></tr>`;
  }
}

function renderOverview(){
  populateBDMonthFilter();
  renderBDFunnel();
  const openOpps=DATA.opportunities.filter(o=>isOpenStage(o.opportunity_status));
  const pipelineVal=openOpps.reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0);
  const weightedVal=openOpps.reduce((s,o)=>s+(Number(o.probability_weighted_value)||0),0);
  const wins=DATA.opportunities.filter(o=>isWinStage(o.opportunity_status));
  const losses=DATA.opportunities.filter(o=>isLossStage(o.opportunity_status));
  const winRate=(wins.length+losses.length)>0?Math.round((wins.length/(wins.length+losses.length))*100):null;
  const overdueFollowups=DATA.engagements.filter(e=>{
    if(String(e.follow_up_done||'').toLowerCase()==='yes') return false;
    if(!e.cta_due_date) return false;
    return new Date(e.cta_due_date)<new Date();
  }).length;

  renderKpiCards('ov-kpis',[
    {label:'Open pipeline',value:fmtMoney(pipelineVal),sub:`${openOpps.length} opportunities`,accent:'sky'},
    {label:'Weighted pipeline',value:fmtMoney(weightedVal),sub:'probability-adjusted',accent:'sky'},
    {label:'Win rate',value:winRate===null?'':winRate+'%',sub:`${wins.length} Win · ${losses.length} lost`,accent:'green'},
    {label:'Overdue CTAs',value:fmtNum(overdueFollowups),sub:'follow-up past due date',accent:overdueFollowups>0?'coral':'green'}
  ]);

  renderOverviewBDTables();

  // Funnel — count shown on hover only; est./weighted value shown as columns
  const stages=LOOKUPS.opportunityStatus;
  const stageOpps=stages.map(s=>DATA.opportunities.filter(o=>o.opportunity_status===s));
  const counts=stageOpps.map(list=>list.length);
  const stageEst=stageOpps.map(list=>list.reduce((a,o)=>a+(Number(o.estimated_value_usd)||0),0));
  const stageWeighted=stageOpps.map(list=>list.reduce((a,o)=>a+(Number(o.probability_weighted_value)||0),0));
  const maxCount=Math.max(1,...counts);
  const totalOpps=counts.reduce((a,b)=>a+b,0);
  document.getElementById('ov-funnel').innerHTML=
    `<div class="funnel-row funnel-head"><div>Stage</div><div></div><div class="funnel-val" style="text-transform:none;">Est. Value</div><div class="funnel-val" style="text-transform:none;">Weighted</div></div>`
    +stages.map((s,i)=>`
    <div class="funnel-row">
      <div class="funnel-label">${esc(s.replace(/^\d+\s/,''))}</div>
      <div class="funnel-track" title="${counts[i]} opportunit${counts[i]===1?'y':'ies'}"><div class="funnel-fill" style="width:${(counts[i]/maxCount*100).toFixed(0)}%;background:${STAGE_COLORS[i+1]||COLORS.sky};opacity:0.85;"></div></div>
      <div class="funnel-val">${fmtMoney(stageEst[i])}</div>
      <div class="funnel-val">${fmtMoney(stageWeighted[i])}</div>
    </div>`).join('')
    +`<div class="funnel-row" style="border-top:2px solid var(--ink);font-weight:700;">
      <div class="funnel-label">Total (${totalOpps} opportunit${totalOpps===1?'y':'ies'})</div>
      <div></div>
      <div class="funnel-val">${fmtMoney(stageEst.reduce((a,b)=>a+b,0))}</div>
      <div class="funnel-val">${fmtMoney(stageWeighted.reduce((a,b)=>a+b,0))}</div>
    </div>`;

  // Weekly chart — labelled by week-start date, not raw ISO week strings
  const ovWeekSeries=weeklySeries(DATA.engagements.map(e=>e.eng_date),10);
  destroyChart('ovWeekly');
  const c1=document.getElementById('ov-weekly-chart').getContext('2d');
  CHARTS.ovWeekly=new Chart(c1,{type:'line',data:{labels:ovWeekSeries.map(s=>s.label),datasets:[{label:'Engagements',data:ovWeekSeries.map(s=>s.count),borderColor:COLORS.sky,backgroundColor:COLORS.skyPale,fill:true,tension:0.35,pointRadius:3,pointBackgroundColor:COLORS.skyDeep}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{callbacks:{title:items=>ovWeekSeries[items[0].dataIndex]?.rangeLabel||''}}},scales:{y:{beginAtZero:true,ticks:{precision:0}},x:{grid:{display:false},title:{display:true,text:'Week starting'}}}}});

  // Account health — touch points (replaces the old "Pipeline by company" list)
  const ahRows=computeCompanyRollup()
    .filter(r=>r.engagements>0||r.wins>0||r.toWin>0||r.pipelineValue>0||r.losses>0)
    .map(r=>{
      const pipeline=r.pipelineValue-r.toWinValue;
      const total=r.winValue+r.toWinValue+pipeline+r.lossValue;
      return {...r,pipeline,total};
    })
    .sort((a,b)=>b.total-a.total);
  const ahTbl=document.getElementById('ov-account-health-table');
  ahTbl.querySelector('thead').innerHTML='<tr><th>Account</th><th class="num">Total Touch Points</th><th class="num">Win</th><th class="num">To Win</th><th class="num">Pipeline</th><th class="num">Lose</th><th class="num">Total</th></tr>';
  ahTbl.querySelector('tbody').innerHTML=ahRows.length?ahRows.map(r=>`
    <tr>
      <td><strong>${esc(r.company)}</strong></td>
      <td class="num">${fmtNum(r.engagements)}</td>
      <td class="num">${r.winValue?fmtMoney(r.winValue):''}</td>
      <td class="num">${r.toWinValue?fmtMoney(r.toWinValue):''}</td>
      <td class="num">${r.pipeline?fmtMoney(r.pipeline):''}</td>
      <td class="num">${r.lossValue?fmtMoney(r.lossValue):''}</td>
      <td class="num"><strong>${fmtMoney(r.total)}</strong></td>
    </tr>`).join('')
    :'<tr><td colspan="7" class="empty-state">No account activity yet.</td></tr>';

  // Follow-ups due / overdue
  const now=new Date(); const in7=new Date(); in7.setDate(in7.getDate()+7);
  const followups=DATA.engagements.filter(e=>{
    if(String(e.follow_up_done||'').toLowerCase()==='yes') return false;
    if(!e.cta_due_date) return false;
    return new Date(e.cta_due_date)<=in7;
  }).sort((a,b)=>new Date(a.cta_due_date)-new Date(b.cta_due_date)).slice(0,8);
  document.getElementById('ov-followup-list').innerHTML=followups.length?followups.map(e=>{
    const overdue=new Date(e.cta_due_date)<now;
    return `<div class="alert-row">
      <div class="alert-dot" style="background:${overdue?'var(--coral)':'var(--amber)'}"></div>
      <div><div class="alert-text">${esc(e.client_name||'')} · ${esc(e.company||'')}</div>
      <div class="alert-meta">${esc(e.cta_next_step||'')}</div></div>
      <div class="alert-date">${fmtDate(e.cta_due_date)}</div>
    </div>`}).join('')
    :'<div class="empty-state" style="padding:10px 0">No follow-ups due this week 🎉</div>';

  // Recent outcomes — the 10 most recent wins/losses, ranked by the date the stage
  // actually changed to Win/Loss (from stage_history), falling back to expected close date.
  const recentOpps=DATA.opportunities
    .filter(o=>isWinStage(o.opportunity_status)||isLossStage(o.opportunity_status))
    .map(o=>{
      const hist=getHistory(o);
      const decidedDate=(hist.length?hist[hist.length-1].date:null)||o.expected_close_date||null;
      return {o,decidedDate};
    })
    .sort((a,b)=>new Date(b.decidedDate||0)-new Date(a.decidedDate||0))
    .slice(0,10);
  document.getElementById('ov-outcomes-list').innerHTML=recentOpps.length?recentOpps.map(({o,decidedDate})=>{
    const won=isWinStage(o.opportunity_status);
    return `<div class="alert-row">
      <div class="alert-dot" style="background:${won?'var(--green)':'var(--coral)'}"></div>
      <div><div class="alert-text">${esc(o.opportunity||'')} · ${esc(o.company||'')}</div>
      <div class="alert-meta">${won?'Win':'Lost'} · ${esc(o.bd_owner||'')} · ${fmtDate(decidedDate)}</div></div>
      <div class="alert-date">${fmtMoney(o.estimated_value_usd)}</div>
    </div>`}).join('')
    :'<div class="empty-state" style="padding:10px 0">No recent outcomes</div>';
}

/* ---------- 13. RENDER: BD SALES FUNNEL TRACKER ---------- */
function populateBDMonthFilter(){
  const sel=document.getElementById('bd-month-sel');
  if(!sel) return;
  const cur=sel.value;
  // Collect unique months from engagement log
  const months=new Set();
  DATA.engagements.forEach(e=>{
    if(e.eng_date){const d=new Date(e.eng_date);if(!isNaN(d))months.add(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`);}
  });
  const sorted=[...months].sort().reverse();
  sel.innerHTML='<option value="all">All time</option>'+sorted.map(m=>{
    const [y,mo]=m.split('-');
    const label=new Date(Number(y),Number(mo)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'});
    return `<option value="${m}"${cur===m?' selected':''}>${label}</option>`;
  }).join('');
  if(cur) sel.value=cur;
}

function renderBDFunnel(){
  const sel=document.getElementById('bd-month-sel');
  const monthFilter=sel?sel.value:'all';

  // Helper: filter engagements by selected month
  function filterEngs(arr){
    if(monthFilter==='all') return arr;
    return arr.filter(e=>{
      if(!e.eng_date) return false;
      const d=new Date(e.eng_date);
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`===monthFilter;
    });
  }
  const engs=filterEngs(DATA.engagements);

  // Update section title to show selected month
  const titleEl=document.getElementById('bd-eng-title');
  if(titleEl){
    if(monthFilter==='all') titleEl.textContent='Engagement Overview — All Time';
    else {
      const [y,mo]=monthFilter.split('-');
      const label=new Date(Number(y),Number(mo)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'});
      titleEl.textContent=`Engagement Overview — ${label}`;
    }
  }

  /* ---- SECTION 1: STAKEHOLDER LIST ---- */
  const depts=[
    {label:'Total Clients',key:null},
    {label:'Medical',key:'Medical'},
    {label:'Access / HEOR',key:'Market Access/ HEOR'},
    {label:'Marketing / Comm.',key:'Marketing/ Commercial'},
    {label:'Procurement',key:'Procurement'},
    {label:'Digital',key:'Digital'},
    {label:'Others',key:'Others'}
  ];

  const isActiveStatus=c=>['active business','active engagement'].includes(String(c.status||'').toLowerCase());

  const stakeViewSel=document.getElementById('bd-stake-view');
  if(stakeViewSel) stakeViewSel.value=bdStakeView;
  const activeRowEl=document.getElementById('bd-s-active');
  if(activeRowEl) activeRowEl.style.display=bdStakeView==='active'?'':'none';

  const stakeDepts=bdStakeExpanded?depts:depts.slice(0,1);
  const stakeToggle=document.getElementById('bd-stake-toggle');
  if(stakeToggle) stakeToggle.textContent=bdStakeExpanded?'Hide breakdown ▴':'Show breakdown ▾';

  const totalEl=document.getElementById('bd-s-total');
  const activeEl=document.getElementById('bd-s-active');
  if(totalEl) totalEl.innerHTML=stakeDepts.map(d=>{
    const count=d.key?DATA.clients.filter(c=>c.department===d.key).length:DATA.clients.length;
    return `<div class="bd-metric">
      <div class="bd-metric-label">${esc(d.label)}</div>
      <div class="bd-metric-val">${count}</div>
      <div class="bd-metric-sub">total</div>
    </div>`;
  }).join('');
  if(activeEl) activeEl.innerHTML=stakeDepts.map(d=>{
    const count=d.key
      ?DATA.clients.filter(c=>c.department===d.key&&isActiveStatus(c)).length
      :DATA.clients.filter(c=>isActiveStatus(c)).length;
    return `<div class="bd-metric">
      <div class="bd-metric-label">${esc(d.label)}</div>
      <div class="bd-metric-val">${count}</div>
      <div class="bd-metric-sub">active</div>
    </div>`;
  }).join('');

  /* ---- SECTION 2: ENGAGEMENT OVERVIEW ---- */
  const ftf=engs.filter(e=>e.engagement_type==='In-Person Meeting').length;
  const virt=engs.filter(e=>e.engagement_type==='Virtual Meeting').length;
  const call=engs.filter(e=>e.engagement_type==='Phone Call/ Whatsapp').length;
  const email=engs.filter(e=>e.engagement_type==='Email').length;
  const conference=engs.filter(e=>e.engagement_type==='Conference/Event').length;
  const others=engs.filter(e=>e.engagement_type==='Others').length;

  const engMetrics=[
    {label:'Active Engagements',val:engs.length},
    {label:'Face to Face',val:ftf},
    {label:'Virtual Meetings',val:virt},
    {label:'Calls / WhatsApp',val:call},
    {label:'Email',val:email},
    {label:'Conference',val:conference},
    {label:'Others',val:others}
  ];
  const engEl=document.getElementById('bd-eng');
  if(engEl) engEl.innerHTML=(bdEngExpanded?engMetrics:engMetrics.slice(0,1)).map(m=>`<div class="bd-metric">
    <div class="bd-metric-label">${esc(m.label)}</div>
    <div class="bd-metric-val">${m.val}</div>
  </div>`).join('');
  const engToggle=document.getElementById('bd-eng-toggle');
  if(engToggle) engToggle.textContent=bdEngExpanded?'Hide breakdown ▴':'Show breakdown ▾';

  /* ---- SECTION 3: OPPORTUNITY MAPPING ---- */
  const opps=DATA.opportunities;
  const stN=o=>STAGE_NUM(o.opportunity_status);
  const oppMetrics=[
    {label:'Total Mapped',val:opps.length},
    {label:'Blue Sky',val:opps.filter(o=>stN(o)===1).length},
    {label:'Identified / Discussed',val:opps.filter(o=>stN(o)===2||stN(o)===3).length},
    {label:'Proposals',val:opps.filter(o=>stN(o)===4||stN(o)===5).length},
    {label:'To Win / Active',val:opps.filter(o=>stN(o)===6).length},
    {label:'Win',val:opps.filter(o=>isWinStage(o.opportunity_status)).length},
    {label:'Lost / Missed',val:opps.filter(o=>stN(o)>=8).length}
  ];
  const oppEl=document.getElementById('bd-opp');
  if(oppEl) oppEl.innerHTML=(bdOppExpanded?oppMetrics:oppMetrics.slice(0,1)).map(m=>`<div class="bd-metric">
    <div class="bd-metric-label">${esc(m.label)}</div>
    <div class="bd-metric-val">${m.val}</div>
  </div>`).join('');
  const oppToggle=document.getElementById('bd-opp-toggle');
  if(oppToggle) oppToggle.textContent=bdOppExpanded?'Hide breakdown ▴':'Show breakdown ▾';

  /* ---- SECTION 4: PIPELINE FUNNEL — the 7 open/won stages, individually, plus a Total ---- */
  function sumVal(arr){ return arr.reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0); }
  const pipCols=[1,2,3,4,5,6,7].map(n=>({
    label:LOOKUPS.opportunityStatus[n-1].replace(/^\d+\s/,''),
    stageNum:n,
    color:STAGE_COLORS[n]
  }));
  const pipBuckets=pipCols.map(col=>opps.filter(o=>stN(o)===col.stageNum));
  const totalBucket=pipBuckets.flat();
  const pipEl=document.getElementById('bd-pipeline');
  if(pipEl){
    const stageColsHtml=pipCols.map((col,i)=>{
      const bucket=pipBuckets[i];
      const isWin=col.stageNum===7;
      return `<div class="bd-pipeline-col">
        <div class="bd-pipeline-col-head" style="background:${col.color};">${esc(col.label)}</div>
        <div class="bd-pipeline-row">
          <div class="bd-pipeline-cell">
            <div class="bd-pipeline-cell-val${isWin?' green':''}">${bucket.length}</div>
            <div class="bd-pipeline-cell-label">opportunities</div>
          </div>
          <div class="bd-pipeline-cell">
            <div class="bd-pipeline-cell-val${isWin?' green':''}">${fmtMoney(sumVal(bucket))}</div>
            <div class="bd-pipeline-cell-label">est. value</div>
          </div>
        </div>
      </div>`;
    }).join('');
    const totalColHtml=`<div class="bd-pipeline-col">
      <div class="bd-pipeline-col-head" style="background:${COLORS.ink};">Total</div>
      <div class="bd-pipeline-row">
        <div class="bd-pipeline-cell">
          <div class="bd-pipeline-cell-val">${totalBucket.length}</div>
          <div class="bd-pipeline-cell-label">opportunities</div>
        </div>
        <div class="bd-pipeline-cell">
          <div class="bd-pipeline-cell-val">${fmtMoney(sumVal(totalBucket))}</div>
          <div class="bd-pipeline-cell-label">est. value</div>
        </div>
      </div>
    </div>`;
    pipEl.innerHTML=stageColsHtml+totalColHtml;
  }
}

/* ---------- 13b. ALL REGIONS OVERVIEW (admin-only, read-only rollup) ---------- */
async function enterAllRegions(){
  if(!currentUser||currentUser.role!=='admin'){ alert('All Regions is only available to admins.'); return; }
  document.getElementById('region-screen').style.display='none';
  document.getElementById('brand-mark').style.display='none';
  document.getElementById('all-regions-screen').style.display='block';
  const loadingEl=document.getElementById('ar-loading');
  if(loadingEl) loadingEl.style.display='block';
  syncCurrencySelects();
  ensureFxRate().then(()=>{ syncCurrencySelects(); if(displayCurrency==='SGD'&&ALL_REGIONS_DATA) renderAllRegionsOverview(); });
  try{
    ALL_REGIONS_DATA=await fetchAllRegionsData();
    renderAllRegionsOverview();
  }catch(err){
    console.error(err);
    alert('Could not load all-regions data: '+err.message);
  }finally{
    if(loadingEl) loadingEl.style.display='none';
  }
}

// Fetches every region's 4 SharePoint lists in parallel and tags each record with which
// region it came from, so per-region breakdowns (comparison chart/tables) stay possible
// even though everything else here treats it as one merged dataset.
async function fetchAllRegionsData(){
  const entries=Object.entries(REGIONS);
  const byRegion=await Promise.all(entries.map(async ([key,r])=>{
    const [c,o,e,co]=await Promise.all([
      graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.clients}/items?expand=fields&$top=500`),
      graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.opportunities}/items?expand=fields&$top=500`),
      graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.engagements}/items?expand=fields&$top=500`),
      graphGet(`${GRAPH_BASE}/sites/${r.siteId}/lists/${r.listIds.companies}/items?expand=fields&$top=500`)
    ]);
    const tag=rec=>({...rec,_region:key,_regionLabel:r.label});
    return {
      key, label:r.label,
      clients:c.map(i=>spItemToRecord('clients',i)).map(tag),
      opportunities:o.map(i=>spItemToRecord('opportunities',i)).map(tag),
      engagements:e.map(i=>spItemToRecord('engagements',i)).map(tag),
      companies:co.map(i=>spItemToRecord('companies',i)).map(tag)
    };
  }));
  const merged={
    clients:byRegion.flatMap(r=>r.clients),
    opportunities:byRegion.flatMap(r=>r.opportunities),
    engagements:byRegion.flatMap(r=>r.engagements),
    companies:byRegion.flatMap(r=>r.companies)
  };
  return {merged, byRegion};
}

// Mirrors populateBDMonthFilter() — built once per data load, never rebuilt from inside
// the month select's own change handler (that's what breaks a <select>'s native picker).
function populateArBdMonthFilter(){
  const sel=document.getElementById('ar-bd-month-sel');
  if(!sel) return;
  const cur=sel.value;
  const months=new Set();
  ALL_REGIONS_DATA.merged.engagements.forEach(e=>{
    if(e.eng_date){const d=new Date(e.eng_date);if(!isNaN(d))months.add(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`);}
  });
  const sorted=[...months].sort().reverse();
  sel.innerHTML='<option value="all">All time</option>'+sorted.map(m=>{
    const [y,mo]=m.split('-');
    const label=new Date(Number(y),Number(mo)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'});
    return `<option value="${m}"${cur===m?' selected':''}>${label}</option>`;
  }).join('');
  if(cur) sel.value=cur;
}

function renderAllRegionsOverview(){
  if(!ALL_REGIONS_DATA) return;
  populateArBdMonthFilter();
  renderAllRegionsBDFunnel();
  renderAllRegionsKpis();
  renderAllRegionsCompareChart();
  renderAllRegionsRegionTable();
  renderAllRegionsStageTable();
  renderAllRegionsTrendChart();
}

// Mirrors renderBDFunnel()'s Sections 1-3 exactly, just against the merged all-regions
// dataset and its own ar*-prefixed element ids / expand-state so the two pages never
// interfere with each other.
function renderAllRegionsBDFunnel(){
  const data=ALL_REGIONS_DATA.merged;
  const sel=document.getElementById('ar-bd-month-sel');
  const monthFilter=sel?sel.value:'all';

  function filterEngs(arr){
    if(monthFilter==='all') return arr;
    return arr.filter(e=>{
      if(!e.eng_date) return false;
      const d=new Date(e.eng_date);
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`===monthFilter;
    });
  }
  const engs=filterEngs(data.engagements);

  const titleEl=document.getElementById('ar-bd-eng-title');
  if(titleEl){
    if(monthFilter==='all') titleEl.textContent='Engagement Overview — All Time';
    else {
      const [y,mo]=monthFilter.split('-');
      const label=new Date(Number(y),Number(mo)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'});
      titleEl.textContent=`Engagement Overview — ${label}`;
    }
  }

  /* ---- SECTION 1: STAKEHOLDER LIST ---- */
  const depts=[
    {label:'Total Clients',key:null},
    {label:'Medical',key:'Medical'},
    {label:'Access / HEOR',key:'Market Access/ HEOR'},
    {label:'Marketing / Comm.',key:'Marketing/ Commercial'},
    {label:'Procurement',key:'Procurement'},
    {label:'Digital',key:'Digital'},
    {label:'Others',key:'Others'}
  ];
  const isActiveStatus=c=>['active business','active engagement'].includes(String(c.status||'').toLowerCase());

  const stakeViewSel=document.getElementById('ar-bd-stake-view');
  if(stakeViewSel) stakeViewSel.value=arStakeView;
  const activeRowEl=document.getElementById('ar-bd-s-active');
  if(activeRowEl) activeRowEl.style.display=arStakeView==='active'?'':'none';

  const stakeDepts=arStakeExpanded?depts:depts.slice(0,1);
  const stakeToggle=document.getElementById('ar-bd-stake-toggle');
  if(stakeToggle) stakeToggle.textContent=arStakeExpanded?'Hide breakdown ▴':'Show breakdown ▾';

  const totalEl=document.getElementById('ar-bd-s-total');
  const activeEl=document.getElementById('ar-bd-s-active');
  if(totalEl) totalEl.innerHTML=stakeDepts.map(d=>{
    const count=d.key?data.clients.filter(c=>c.department===d.key).length:data.clients.length;
    return `<div class="bd-metric">
      <div class="bd-metric-label">${esc(d.label)}</div>
      <div class="bd-metric-val">${count}</div>
      <div class="bd-metric-sub">total</div>
    </div>`;
  }).join('');
  if(activeEl) activeEl.innerHTML=stakeDepts.map(d=>{
    const count=d.key
      ?data.clients.filter(c=>c.department===d.key&&isActiveStatus(c)).length
      :data.clients.filter(c=>isActiveStatus(c)).length;
    return `<div class="bd-metric">
      <div class="bd-metric-label">${esc(d.label)}</div>
      <div class="bd-metric-val">${count}</div>
      <div class="bd-metric-sub">active</div>
    </div>`;
  }).join('');

  /* ---- SECTION 2: ENGAGEMENT OVERVIEW ---- */
  const ftf=engs.filter(e=>e.engagement_type==='In-Person Meeting').length;
  const virt=engs.filter(e=>e.engagement_type==='Virtual Meeting').length;
  const call=engs.filter(e=>e.engagement_type==='Phone Call/ Whatsapp').length;
  const email=engs.filter(e=>e.engagement_type==='Email').length;
  const conference=engs.filter(e=>e.engagement_type==='Conference/Event').length;
  const others=engs.filter(e=>e.engagement_type==='Others').length;

  const engMetrics=[
    {label:'Active Engagements',val:engs.length},
    {label:'Face to Face',val:ftf},
    {label:'Virtual Meetings',val:virt},
    {label:'Calls / WhatsApp',val:call},
    {label:'Email',val:email},
    {label:'Conference',val:conference},
    {label:'Others',val:others}
  ];
  const engEl=document.getElementById('ar-bd-eng');
  if(engEl) engEl.innerHTML=(arEngExpanded?engMetrics:engMetrics.slice(0,1)).map(m=>`<div class="bd-metric">
    <div class="bd-metric-label">${esc(m.label)}</div>
    <div class="bd-metric-val">${m.val}</div>
  </div>`).join('');
  const engToggle=document.getElementById('ar-bd-eng-toggle');
  if(engToggle) engToggle.textContent=arEngExpanded?'Hide breakdown ▴':'Show breakdown ▾';

  /* ---- SECTION 3: OPPORTUNITY MAPPING ---- */
  const opps=data.opportunities;
  const stN=o=>STAGE_NUM(o.opportunity_status);
  const oppMetrics=[
    {label:'Total Mapped',val:opps.length},
    {label:'Blue Sky',val:opps.filter(o=>stN(o)===1).length},
    {label:'Identified / Discussed',val:opps.filter(o=>stN(o)===2||stN(o)===3).length},
    {label:'Proposals',val:opps.filter(o=>stN(o)===4||stN(o)===5).length},
    {label:'To Win / Active',val:opps.filter(o=>stN(o)===6).length},
    {label:'Win',val:opps.filter(o=>isWinStage(o.opportunity_status)).length},
    {label:'Lost / Missed',val:opps.filter(o=>stN(o)>=8).length}
  ];
  const oppEl=document.getElementById('ar-bd-opp');
  if(oppEl) oppEl.innerHTML=(arOppExpanded?oppMetrics:oppMetrics.slice(0,1)).map(m=>`<div class="bd-metric">
    <div class="bd-metric-label">${esc(m.label)}</div>
    <div class="bd-metric-val">${m.val}</div>
  </div>`).join('');
  const oppToggle=document.getElementById('ar-bd-opp-toggle');
  if(oppToggle) oppToggle.textContent=arOppExpanded?'Hide breakdown ▴':'Show breakdown ▾';
}

// Filters engagements to the period picked in ar-period-filter. Every other KPI here
// (pipeline value, win rate) stays all-time — an open opportunity doesn't stop being open
// just because it was created outside the selected window, same convention used elsewhere
// in this app (e.g. the Opportunities page's own KPIs are never period-limited either).
function arPeriodEngagements(){
  const start=periodStart(arPeriod);
  return ALL_REGIONS_DATA.merged.engagements.filter(e=>e.eng_date&&new Date(e.eng_date)>=start);
}
function renderAllRegionsKpis(){
  const opps=ALL_REGIONS_DATA.merged.opportunities;
  const start=periodStart(arPeriod);
  const periodEngs=arPeriodEngagements();

  const rangeEl=document.getElementById('ar-period-range');
  if(rangeEl) rangeEl.textContent=`${fmtDate(start)} – ${fmtDate(new Date())}`;

  const open=opps.filter(o=>isOpenStage(o.opportunity_status));
  const wins=opps.filter(o=>isWinStage(o.opportunity_status));
  const losses=opps.filter(o=>isLossStage(o.opportunity_status));
  const winRate=(wins.length+losses.length)>0?Math.round((wins.length/(wins.length+losses.length))*100):null;
  renderKpiCards('ar-kpis',[
    {label:'Open pipeline',value:fmtMoney(open.reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0)),sub:`${open.length} opportunities · all time`,accent:'sky'},
    {label:'Weighted pipeline',value:fmtMoney(open.reduce((s,o)=>s+(Number(o.probability_weighted_value)||0),0)),sub:'across all regions · all time',accent:'sky'},
    {label:'Win rate',value:winRate===null?'':winRate+'%',sub:`${wins.length} Win · ${losses.length} lost · all time`,accent:'green'},
    {label:'Engagements logged',value:fmtNum(periodEngs.length),sub:'selected period',accent:'amber'}
  ]);
}

// Per-region weighted/open pipeline, side by side.
function renderAllRegionsCompareChart(){
  const byRegion=ALL_REGIONS_DATA.byRegion;
  const labels=byRegion.map(r=>r.label);
  const openVals=byRegion.map(r=>r.opportunities.filter(o=>isOpenStage(o.opportunity_status)).reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0));
  const weightedVals=byRegion.map(r=>r.opportunities.filter(o=>isOpenStage(o.opportunity_status)).reduce((s,o)=>s+(Number(o.probability_weighted_value)||0),0));
  destroyChart('arRegionCompare');
  const ctx=document.getElementById('ar-region-chart').getContext('2d');
  CHARTS.arRegionCompare=new Chart(ctx,{
    type:'bar',
    data:{labels,datasets:[
      {label:'Open pipeline (est.)',data:openVals,backgroundColor:COLORS.sky,borderRadius:4,maxBarThickness:48},
      {label:'Weighted pipeline',data:weightedVals,backgroundColor:COLORS.amber,borderRadius:4,maxBarThickness:48}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{beginAtZero:true,ticks:{callback:v=>fmtMoney(v)}}},
      plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:10}}},
        tooltip:{callbacks:{label:item=>`${item.dataset.label}: ${fmtMoney(item.raw)}`}}}}
  });
}

function renderAllRegionsRegionTable(){
  const byRegion=ALL_REGIONS_DATA.byRegion;
  const start=periodStart(arPeriod);
  const rows=byRegion.map(r=>{
    const open=r.opportunities.filter(o=>isOpenStage(o.opportunity_status));
    const wins=r.opportunities.filter(o=>isWinStage(o.opportunity_status));
    const losses=r.opportunities.filter(o=>isLossStage(o.opportunity_status));
    const winRate=(wins.length+losses.length)>0?Math.round((wins.length/(wins.length+losses.length))*100):null;
    const periodEngs=r.engagements.filter(e=>e.eng_date&&new Date(e.eng_date)>=start);
    return {
      label:r.label,
      clients:r.clients.length,
      activeClients:r.clients.filter(c=>['active business','active engagement'].includes(String(c.status||'').toLowerCase())).length,
      openOpps:open.length,
      pipeline:open.reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0),
      weighted:open.reduce((s,o)=>s+(Number(o.probability_weighted_value)||0),0),
      wins:wins.length,
      winRate,
      engagements:periodEngs.length
    };
  });
  const grand={
    clients:rows.reduce((s,r)=>s+r.clients,0),
    activeClients:rows.reduce((s,r)=>s+r.activeClients,0),
    openOpps:rows.reduce((s,r)=>s+r.openOpps,0),
    pipeline:rows.reduce((s,r)=>s+r.pipeline,0),
    weighted:rows.reduce((s,r)=>s+r.weighted,0),
    wins:rows.reduce((s,r)=>s+r.wins,0),
    engagements:rows.reduce((s,r)=>s+r.engagements,0)
  };
  const tbl=document.getElementById('ar-region-table');
  tbl.querySelector('thead').innerHTML=`<tr><th>Region</th><th class="num">Clients</th><th class="num">Active</th>
    <th class="num">Open Opps</th><th class="num">Pipeline</th><th class="num">Weighted</th>
    <th class="num">Wins</th><th class="num">Win Rate</th><th class="num">Engagements (period)</th></tr>`;
  tbl.querySelector('tbody').innerHTML=rows.map(r=>`<tr>
      <td><strong>${esc(r.label)}</strong></td>
      <td class="num">${fmtNum(r.clients)}</td>
      <td class="num">${fmtNum(r.activeClients)}</td>
      <td class="num">${fmtNum(r.openOpps)}</td>
      <td class="num">${fmtMoney(r.pipeline)}</td>
      <td class="num">${fmtMoney(r.weighted)}</td>
      <td class="num">${fmtNum(r.wins)}</td>
      <td class="num">${r.winRate===null?'':r.winRate+'%'}</td>
      <td class="num">${fmtNum(r.engagements)}</td>
    </tr>`).join('')+
    `<tr class="pivot-total"><td>Grand Total</td><td class="num">${grand.clients}</td><td class="num">${grand.activeClients}</td>
      <td class="num">${grand.openOpps}</td><td class="num">${fmtMoney(grand.pipeline)}</td><td class="num">${fmtMoney(grand.weighted)}</td>
      <td class="num">${grand.wins}</td><td class="num"></td><td class="num">${grand.engagements}</td></tr>`;
}

// Region x Stage — counts only, mirrors the single-region Pipeline Funnel's 7 open/won stages.
function renderAllRegionsStageTable(){
  const byRegion=ALL_REGIONS_DATA.byRegion;
  const stN=o=>STAGE_NUM(o.opportunity_status);
  const stageCols=[1,2,3,4,5,6,7].map(n=>LOOKUPS.opportunityStatus[n-1].replace(/^\d+\s/,''));
  const tbl=document.getElementById('ar-stage-table');
  tbl.querySelector('thead').innerHTML=`<tr><th>Region</th>${stageCols.map(l=>`<th class="num">${esc(l)}</th>`).join('')}<th class="num">Total</th></tr>`;
  const colTotals=stageCols.map(()=>0);
  let grandTotal=0;
  const rowsHtml=byRegion.map(r=>{
    const counts=[1,2,3,4,5,6,7].map(n=>r.opportunities.filter(o=>stN(o)===n).length);
    counts.forEach((n,i)=>colTotals[i]+=n);
    const total=counts.reduce((a,b)=>a+b,0);
    grandTotal+=total;
    return `<tr><td>${esc(r.label)}</td>${counts.map(n=>`<td class="num">${n||''}</td>`).join('')}<td class="num">${total}</td></tr>`;
  }).join('');
  tbl.querySelector('tbody').innerHTML=rowsHtml+
    `<tr class="pivot-total"><td>Grand Total</td>${colTotals.map(n=>`<td class="num">${n}</td>`).join('')}<td class="num">${grandTotal}</td></tr>`;
}

// One line per region, aligned on the same set of ISO week keys so they're comparable.
function renderAllRegionsTrendChart(){
  const byRegion=ALL_REGIONS_DATA.byRegion;
  const allDates=byRegion.flatMap(r=>r.engagements.map(e=>e.eng_date)).filter(Boolean);
  const mondayByKey={};
  allDates.forEach(ds=>{ const wk=isoWeek(ds); if(wk&&!mondayByKey[wk]) mondayByKey[wk]=mondayOf(ds); });
  let keys=Object.keys(mondayByKey).sort();
  keys=keys.slice(-12);
  const labels=keys.map(k=>mondayByKey[k]?mondayByKey[k].toLocaleDateString('en-GB',{day:'2-digit',month:'short'}):k);
  const palette=[COLORS.sky,COLORS.amber,COLORS.green,COLORS.coral];
  const datasets=byRegion.map((r,i)=>{
    const countByKey={};
    r.engagements.forEach(e=>{ if(!e.eng_date) return; const wk=isoWeek(e.eng_date); if(!wk) return; countByKey[wk]=(countByKey[wk]||0)+1; });
    return {label:r.label,data:keys.map(k=>countByKey[k]||0),borderColor:palette[i%palette.length],backgroundColor:palette[i%palette.length],tension:0.3,pointRadius:3};
  });
  destroyChart('arTrend');
  const ctx=document.getElementById('ar-trend-chart').getContext('2d');
  CHARTS.arTrend=new Chart(ctx,{type:'line',data:{labels,datasets},
    options:{responsive:true,maintainAspectRatio:false,
      scales:{y:{beginAtZero:true,ticks:{precision:0}},x:{title:{display:true,text:'Week starting'}}},
      plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:10}}}}}});
}

/* ---------- 14. RENDER: CLIENTS ---------- */
let clFilters={status:'',company:'',department:'',priority:''};
let clPivotFilters={therapy:''};
let clPivotExpanded=new Set();
let clPivotSelectedKey=null;
function selectPivotRow(rowEl,key){
  clPivotSelectedKey=(clPivotSelectedKey===key)?null:key;
  document.querySelectorAll('#cl-pivot-table tr.pivot-row-selected').forEach(r=>r.classList.remove('pivot-row-selected'));
  if(clPivotSelectedKey===key) rowEl.classList.add('pivot-row-selected');
}
const STAKEHOLDER_STATUS_ORDER=['Active business','Active engagement','Prospect','Inactive'];

function toggleClPivotExpand(co){
  if(clPivotExpanded.has(co)) clPivotExpanded.delete(co); else clPivotExpanded.add(co);
}

// Mirrors the workbook's "Stakeholder List" pivot: Company > Department, by status, with grand totals.
function renderClientPivot(){
  const sel=document.getElementById('cl-pivot-therapy');
  if(sel){
    const areas=[...new Set(DATA.clients.map(c=>c.therapy_area).filter(Boolean))].sort();
    sel.innerHTML='<option value="">All therapy areas</option>'+areas.map(a=>`<option value="${esc(a)}">${esc(a)}</option>`).join('');
    sel.value=clPivotFilters.therapy;
  }

  const filtered=DATA.clients.filter(c=>!clPivotFilters.therapy||c.therapy_area===clPivotFilters.therapy);
  const companies={};
  filtered.forEach(c=>{
    const co=(c.company||'Unknown').trim()||'Unknown';
    const dept=(c.department||'Others').trim()||'Others';
    if(!companies[co]) companies[co]={depts:{},totals:{Inactive:0,Prospect:0,'Active business':0,'Active engagement':0}};
    if(!companies[co].depts[dept]) companies[co].depts[dept]={Inactive:0,Prospect:0,'Active business':0,'Active engagement':0};
    if(STAKEHOLDER_STATUS_ORDER.includes(c.status)){
      companies[co].depts[dept][c.status]++;
      companies[co].totals[c.status]++;
    }
  });
  const coNames=Object.keys(companies).sort((a,b)=>a.localeCompare(b));
  const grand={Inactive:0,Prospect:0,'Active business':0,'Active engagement':0};
  coNames.forEach(co=>STAKEHOLDER_STATUS_ORDER.forEach(s=>grand[s]+=companies[co].totals[s]));
  const grandTotal=STAKEHOLDER_STATUS_ORDER.reduce((a,s)=>a+grand[s],0);

  const tbl=document.getElementById('cl-pivot-table');
  tbl.querySelector('thead').innerHTML=`<tr><th>Company / Department</th>${STAKEHOLDER_STATUS_ORDER.map(s=>`<th class="num">${esc(s)}</th>`).join('')}<th class="num">Grand Total</th></tr>`;

  if(!coNames.length){
    tbl.querySelector('tbody').innerHTML=`<tr><td colspan="${STAKEHOLDER_STATUS_ORDER.length+2}" class="empty-state">No stakeholders match this filter.</td></tr>`;
    return;
  }

  let html='';
  coNames.forEach(co=>{
    const t=companies[co].totals;
    const total=STAKEHOLDER_STATUS_ORDER.reduce((a,s)=>a+t[s],0);
    const isOpen=clPivotExpanded.has(co);
    const groupKey=co;
    html+=`<tr class="pivot-group${isOpen?' open':''}${clPivotSelectedKey===groupKey?' pivot-row-selected':''}" onclick="togglePivotGroup(this);toggleClPivotExpand('${esc(co)}');selectPivotRow(this,'${esc(groupKey)}')">
      <td><span class="caret">▶</span> ${esc(co)}</td>
      ${STAKEHOLDER_STATUS_ORDER.map(s=>`<td class="num">${t[s]||''}</td>`).join('')}
      <td class="num">${total}</td>
    </tr>`;
    Object.keys(companies[co].depts).sort().forEach(d=>{
      const dt=companies[co].depts[d];
      const dtotal=STAKEHOLDER_STATUS_ORDER.reduce((a,s)=>a+dt[s],0);
      const childKey=co+'|||'+d;
      html+=`<tr class="pivot-child${isOpen?' show':''}${clPivotSelectedKey===childKey?' pivot-row-selected':''}" onclick="selectPivotRow(this,'${esc(childKey)}')">
        <td>${esc(d)}</td>
        ${STAKEHOLDER_STATUS_ORDER.map(s=>`<td class="num">${dt[s]||''}</td>`).join('')}
        <td class="num">${dtotal}</td>
      </tr>`;
    });
  });
  html+=`<tr class="pivot-total"><td>Grand Total</td>${STAKEHOLDER_STATUS_ORDER.map(s=>`<td class="num">${grand[s]}</td>`).join('')}<td class="num">${grandTotal}</td></tr>`;
  tbl.querySelector('tbody').innerHTML=html;
}

function renderClients(){
  renderClientPivot();
  const total=DATA.clients.length;
  const actBiz=DATA.clients.filter(c=>c.status==='Active business').length;
  const actEng=DATA.clients.filter(c=>c.status==='Active engagement').length;
  const prospect=DATA.clients.filter(c=>c.status==='Prospect').length;
  const inactive=DATA.clients.filter(c=>c.status==='Inactive').length;
  renderKpiCards('cl-kpis',[
    {label:'Total',value:fmtNum(total),sub:`${[...new Set(DATA.clients.map(c=>c.company).filter(Boolean))].length} companies`,accent:'sky'},
    {label:'Active business',value:fmtNum(actBiz),sub:'signed projects',accent:'green'},
    {label:'Active engagement',value:fmtNum(actEng),sub:'in progress',accent:'sky'},
    {label:'Prospects',value:fmtNum(prospect),sub:'nurturing',accent:'amber'},
    {label:'Inactive',value:fmtNum(inactive),sub:'No recent touchpoint',accent:'coral'}
  ]);

  // Populate filter dropdowns alphabetically
  const clSel=document.getElementById('cl-filter-company');
  clSel.innerHTML='<option value="">All companies</option>'+
    sortAlpha(DATA.clients,'company').map(c=>c.company).filter((v,i,a)=>v&&a.indexOf(v)===i).map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  clSel.value=clFilters.company;

  const deptSel=document.getElementById('cl-filter-department');
  deptSel.innerHTML='<option value="">All departments</option>'+
    sortAlpha(DATA.clients,'department').map(c=>c.department).filter((v,i,a)=>v&&a.indexOf(v)===i).map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  deptSel.value=clFilters.department;

  const stSel=document.getElementById('cl-filter-status');
  stSel.innerHTML='<option value="">All statuses</option>'+
    LOOKUPS.clientStatus.map(s=>`<option value="${esc(s)}">${esc(s)}</option>`).join('');
  stSel.value=clFilters.status;

  const priSel=document.getElementById('cl-filter-priority');
  priSel.innerHTML='<option value="">All priorities</option>'+
    LOOKUPS.priority.map(p=>`<option value="${esc(p)}">${esc(p)}</option>`).join('');
  priSel.value=clFilters.priority;

  let rows=DATA.clients.filter(c=>{
    if(clFilters.status&&c.status!==clFilters.status) return false;
    if(clFilters.company&&c.company!==clFilters.company) return false;
    if(clFilters.department&&c.department!==clFilters.department) return false;
    if(clFilters.priority&&c.priority!==clFilters.priority) return false;
    return true;
  });
  rows=applySort(rows,'clients','company');

  const cols=[
    {key:'company',label:'Company'},{key:'client_name',label:'Name'},{key:'designation',label:'Designation'},
    {key:'department',label:'Department'},{key:'therapy_area',label:'Therapy Area'},{key:'region',label:'Region'},
    {key:'status',label:'Status'},{key:'priority',label:'Priority'},{key:'assigned_bd',label:'Assigned BD'},
    {key:'email',label:'Email'},{key:'phone',label:'Phone'},{key:'linkedin_url',label:'LinkedIn'},{key:'notes',label:'Notes'},
    {key:'_actions',label:''}
  ];
  const tbl=document.getElementById('cl-table');
  makeSortableHeaders(tbl.querySelector('thead'),cols,'clients',renderClients);
  tbl.querySelector('tbody').innerHTML=rows.length?rows.map(c=>`
    <tr>
      <td><strong>${esc(c.company)}</strong></td>
      <td>${esc(c.client_name)}</td>
      <td>${esc(c.designation||'')}</td>
      <td>${esc(c.department||'')}</td>
      <td>${esc(c.therapy_area||'')}</td>
      <td>${esc(c.region||'')}</td>
      <td><span class="${clientStatusClass(c.status)}">${esc(c.status||'')}</span></td>
      <td><span class="${priorityClass(c.priority)}">${esc(c.priority||'')}</span></td>
      <td>${esc(c.assigned_bd||'')}</td>
      <td>${c.email?`<a href="mailto:${esc(c.email)}" style="color:var(--sky-deep)">${esc(c.email)}</a>`:''}</td>
      <td>${esc(c.phone||'')}</td>
      <td>${c.linkedin_url?`<a href="${esc(c.linkedin_url)}" target="_blank" style="color:var(--sky-deep)">View</a>`:''}</td>
      <td style="max-width:180px;white-space:normal">${esc(c.notes||'')}</td>
      <td class="td-actions">${canEdit()?`<button class="btn btn-edit btn-sm" onclick="openEditModal('clients','${esc(c.id)}')">Edit</button>`:''}</td>
    </tr>`).join('')
    :'<tr><td colspan="14" class="empty-state">No clients match these filters.</td></tr>';
  document.getElementById('cl-count').textContent=`${rows.length} of ${total}`;
}

/* ---------- 14. RENDER: OPPORTUNITIES ---------- */
let opFilters={search:'',status:'',company:'',bd:''};
let opStageChartCompany='';

// Mirrors the workbook's "Opportunity Tracker" pivot: stage-level count, estimated & weighted value.
function renderOpPivotTable(){
  const opps=opStageChartCompany?DATA.opportunities.filter(o=>o.company===opStageChartCompany):DATA.opportunities;
  const rows=LOOKUPS.opportunityStatus.map(s=>{
    const list=opps.filter(o=>o.opportunity_status===s);
    return {
      stage:s,count:list.length,
      est:list.reduce((a,o)=>a+(Number(o.estimated_value_usd)||0),0),
      weighted:list.reduce((a,o)=>a+(Number(o.probability_weighted_value)||0),0)
    };
  }).filter(r=>r.count>0);
  const grand={
    count:rows.reduce((a,r)=>a+r.count,0),
    est:rows.reduce((a,r)=>a+r.est,0),
    weighted:rows.reduce((a,r)=>a+r.weighted,0)
  };

  const tbl=document.getElementById('op-pivot-table');
  if(!tbl) return;
  tbl.querySelector('thead').innerHTML=`<tr><th>Opportunity Status</th><th class="num">No. of Opportunities</th><th class="num">Estimated Value (USD)</th><th class="num">Probability Weighted Value (USD)</th></tr>`;
  tbl.querySelector('tbody').innerHTML=rows.length
    ? rows.map(r=>`<tr>
        <td><span class="${stageTagClassN(STAGE_NUM(r.stage))}">${esc(r.stage.replace(/^\d+\s/,''))}</span></td>
        <td class="num">${r.count}</td><td class="num">${fmtMoney(r.est)}</td><td class="num">${fmtMoney(r.weighted)}</td>
      </tr>`).join('')+`<tr class="pivot-total"><td>Grand Total</td><td class="num">${grand.count}</td><td class="num">${fmtMoney(grand.est)}</td><td class="num">${fmtMoney(grand.weighted)}</td></tr>`
    : `<tr><td colspan="4" class="empty-state">No opportunities match this filter.</td></tr>`;
}

function renderOpStageChart(){
  const sel=document.getElementById('op-stage-company');
  if(sel){
    const cos=[...new Set(DATA.opportunities.map(o=>o.company).filter(Boolean))].sort();
    sel.innerHTML='<option value="">All companies</option>'+cos.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');
    sel.value=opStageChartCompany;
  }
  renderOpPivotTable();
  const opps=opStageChartCompany?DATA.opportunities.filter(o=>o.company===opStageChartCompany):DATA.opportunities;
  const stages=LOOKUPS.opportunityStatus;
  const stageCounts=stages.map(s=>opps.filter(o=>o.opportunity_status===s).length);
  const stageVals=stages.map(s=>opps.filter(o=>o.opportunity_status===s).reduce((sum,o)=>sum+(Number(o.probability_weighted_value)||0),0));
  destroyChart('opStage');
  const ctx=document.getElementById('op-stage-chart').getContext('2d');
  CHARTS.opStage=new Chart(ctx,{
    data:{labels:stages.map(s=>s.replace(/^\d+\s/,'')),datasets:[
      {type:'bar',label:'Count',data:stageCounts,backgroundColor:STAGE_COLORS.slice(1),yAxisID:'yCount',borderRadius:4,maxBarThickness:36},
      {type:'line',label:'Weighted value',data:stageVals,borderColor:COLORS.amber,backgroundColor:COLORS.amber,yAxisID:'yValue',tension:0.3,pointRadius:3}
    ]},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      scales:{x:{ticks:{font:{size:9},maxRotation:40,minRotation:40},grid:{display:false}},
        yCount:{position:'left',beginAtZero:true,ticks:{precision:0},title:{display:true,text:'Count'}},
        yValue:{position:'right',beginAtZero:true,grid:{drawOnChartArea:false},title:{display:true,text:'USD'}}},
      plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:10}}}}}
  });
}

function renderOpportunities(){
  renderOpStageChart();
  const open=DATA.opportunities.filter(o=>isOpenStage(o.opportunity_status));
  const wins=DATA.opportunities.filter(o=>isWinStage(o.opportunity_status));
  const losses=DATA.opportunities.filter(o=>isLossStage(o.opportunity_status));
  const winRate=(wins.length+losses.length)>0?Math.round((wins.length/(wins.length+losses.length))*100):null;
  renderKpiCards('op-kpis',[
    {label:'Total',value:fmtNum(DATA.opportunities.length),sub:`${open.length} open`,accent:'sky'},
    {label:'Open pipeline',value:fmtMoney(open.reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0)),sub:'estimated USD',accent:'sky'},
    {label:'Weighted',value:fmtMoney(open.reduce((s,o)=>s+(Number(o.probability_weighted_value)||0),0)),sub:'open pipeline',accent:'sky'},
    {label:'Win',value:fmtNum(wins.length),sub:fmtMoney(wins.reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0)),accent:'green'},
    {label:'Lost',value:fmtNum(losses.length),sub:fmtMoney(losses.reduce((s,o)=>s+(Number(o.estimated_value_usd)||0),0)),accent:'coral'},
    {label:'Win rate',value:winRate===null?'':winRate+'%',sub:`${wins.length}W · ${losses.length}L`,accent:winRate>=50?'green':'coral'}
  ]);

  // Populate filters alphabetically
  const opCoSel=document.getElementById('op-filter-company');
  opCoSel.innerHTML='<option value="">All companies</option>'+
    [...new Set(DATA.opportunities.map(o=>o.company).filter(Boolean))].sort().map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  opCoSel.value=opFilters.company;

  const bdSel=document.getElementById('op-filter-bd');
  bdSel.innerHTML='<option value="">All BD owners</option>'+
    [...new Set(DATA.opportunities.map(o=>o.bd_owner).filter(Boolean))].sort().map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  bdSel.value=opFilters.bd;

  const stSel=document.getElementById('op-filter-status');
  stSel.innerHTML='<option value="">All stages</option>'+
    LOOKUPS.opportunityStatus.map(s=>`<option value="${esc(s)}">${esc(s)}</option>`).join('');
  stSel.value=opFilters.status;

  let rows=DATA.opportunities.filter(o=>{
    if(opFilters.status&&o.opportunity_status!==opFilters.status) return false;
    if(opFilters.company&&o.company!==opFilters.company) return false;
    if(opFilters.bd&&o.bd_owner!==opFilters.bd) return false;
    if(opFilters.search){
      const h=`${o.opportunity||''} ${o.company||''} ${o.client_name||''}`.toLowerCase();
      if(!h.includes(opFilters.search.toLowerCase())) return false;
    }
    return true;
  });
  rows=applySort(rows,'opportunities','company');

  const cols=[
    {key:'company',label:'Company'},{key:'client_name',label:'Client'},{key:'opportunity',label:'Opportunity'},
    {key:'opportunity_status',label:'Stage'},{key:'bd_owner',label:'BD Owner'},{key:'supporting_role',label:'Supporting Role'},
    {key:'estimated_value_usd',label:'Est. Value'},{key:'probability_pct',label:'Prob.'},
    {key:'probability_weighted_value',label:'Weighted'},{key:'expected_close_date',label:'Close Date'},
    {key:'notes',label:'Notes'},{key:'_actions',label:''}
  ];
  const tbl=document.getElementById('op-table');
  makeSortableHeaders(tbl.querySelector('thead'),cols,'opportunities',renderOpportunities);
  tbl.querySelector('tbody').innerHTML=rows.length?rows.map(o=>{
    const stageN=STAGE_NUM(o.opportunity_status);
    return `<tr>
      <td><strong>${esc(o.company||'')}</strong></td>
      <td>${esc(o.client_name||'')}</td>
      <td>${esc(o.opportunity||'')}</td>
      <td><span class="${stageTagClassN(stageN)}">${esc((o.opportunity_status||'').replace(/^\d+\s/,''))}</span></td>
      <td>${esc(o.bd_owner||'')}</td>
      <td>${esc(msSelectedFromValue(o.supporting_role).join(', '))}</td>
      <td class="num">${fmtMoney(o.estimated_value_usd)}</td>
      <td class="num">${o.probability_pct!=null?Math.round(o.probability_pct*100)+'%':''}</td>
      <td class="num">${fmtMoney(o.probability_weighted_value)}</td>
      <td>${fmtDate(o.expected_close_date)}</td>
      <td style="max-width:160px;white-space:normal">${esc(o.notes||'')}</td>
      <td class="td-actions">${canEdit()?`<button class="btn btn-edit btn-sm" onclick="openEditModal('opportunities','${esc(o.id)}')">Edit</button>`:''}</td>
    </tr>`}).join('')
    :'<tr><td colspan="12" class="empty-state">No opportunities match these filters.</td></tr>';
  document.getElementById('op-count').textContent=`${rows.length} of ${DATA.opportunities.length}`;
}

/* ---------- 15. RENDER: ENGAGEMENTS ---------- */
let egFilters={type:'',company:'',bd:'',accompanied:'',followup:'',week:''};
function filterEngByCompany(company){
  egFilters.company=(egFilters.company===company)?'':company; // toggle off if clicking the same company again
  renderEngagements();
  const sel=document.getElementById('eg-filter-company'); if(sel) sel.value=egFilters.company;
  document.getElementById('eg-table-panel')?.scrollIntoView({behavior:'smooth',block:'start'});
}
function filterEngByWeek(weekKey){
  egFilters.week=(egFilters.week===weekKey)?'':weekKey; // toggle off if clicking the same week again
  renderEngagements();
  document.getElementById('eg-table-panel')?.scrollIntoView({behavior:'smooth',block:'start'});
}
function clearEgWeekFilter(){ egFilters.week=''; renderEngagements(); }

// Trims stray whitespace before comparing/deduping — live SharePoint data has known
// trailing-space variants (e.g. "Takeda " vs "Takeda"), and an exact === match against
// a clean canonical string would silently zero out results, which looks like "broken filters".
function normStr(v){ return String(v==null?'':v).trim(); }
// Orders a set of actually-present values to match the canonical LOOKUPS order first,
// then appends anything present in the data but not in LOOKUPS (alphabetically) so nothing is ever hidden.
function orderByLookup(presentValues,lookupArr){
  const known=lookupArr.filter(v=>presentValues.includes(v));
  const unknown=presentValues.filter(v=>!lookupArr.includes(v)).sort();
  return [...known,...unknown];
}

/* ---- Engagement pivot: Outcome > Stakeholder Type, cross-tabbed by Engagement Type ----
   Mirrors the workbook's "Engagement Tracker" pivot. Filter option lists are built once
   per data load (EGP_OPTION_FIELDS below); a filter's own change handler only recomputes
   the table and syncs .value — it never touches any <select>'s <option> children. Rebuilding
   a select's options synchronously inside its own 'change' handler corrupts the native
   dropdown's open/focus state in real browsers (invisible in headless/simulated testing,
   but it makes the dropdown stop responding to clicks after a pick or two). */
const EGP_IDS={bd:'egp-bd',month:'egp-month',company:'egp-company',objective:'egp-objective',type:'egp-type'};
let egpFilters={bd:'',month:'',company:'',objective:'',type:''};

function egpMonthKey(dateStr){
  if(!dateStr) return null;
  const d=new Date(dateStr);
  return isNaN(d)?null:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
}
function egpMonthLabel(key){
  const [y,mo]=key.split('-');
  return new Date(Number(y),Number(mo)-1,1).toLocaleDateString('en-GB',{month:'long',year:'numeric'});
}
function egpFillSelect(id,optionsHtml){
  const el=document.getElementById(id);
  if(el) el.innerHTML=optionsHtml;
}

// Populates the 5 dropdowns' option lists from the full (unfiltered) engagement dataset.
// Call once per data refresh — not from a filter's change handler.
function populateEngagementPivotFilters(){
  const engs=DATA.engagements;

  const bds=[...new Set(engs.map(e=>normStr(e.bd_pm)).filter(Boolean))].sort();
  egpFillSelect(EGP_IDS.bd,'<option value="">All BD/PM</option>'+bds.map(b=>`<option value="${esc(b)}">${esc(b)}</option>`).join(''));

  const monthKeys=[...new Set(engs.map(e=>egpMonthKey(e.eng_date)).filter(Boolean))].sort().reverse();
  egpFillSelect(EGP_IDS.month,'<option value="">All months</option>'+monthKeys.map(m=>`<option value="${m}">${egpMonthLabel(m)}</option>`).join(''));

  const cos=[...new Set(engs.map(e=>normStr(e.company)).filter(Boolean))].sort();
  egpFillSelect(EGP_IDS.company,'<option value="">All companies</option>'+cos.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join(''));

  const objs=orderByLookup([...new Set(engs.map(e=>normStr(e.engagement_objective)).filter(Boolean))],LOOKUPS.engagementObjective);
  egpFillSelect(EGP_IDS.objective,'<option value="">All objectives</option>'+objs.map(o=>`<option value="${esc(o)}">${esc(o)}</option>`).join(''));

  const types=orderByLookup([...new Set(engs.map(e=>normStr(e.engagement_type)).filter(Boolean))],LOOKUPS.engagementTypes);
  egpFillSelect(EGP_IDS.type,'<option value="">All engagement types</option>'+types.map(t=>`<option value="${esc(t)}">${esc(t)}</option>`).join(''));

  egpSyncSelectValues();
}

// Sets each <select>'s .value from egpFilters (state is the single source of truth — a
// pick can never desync from what's applied). Cheap: never touches <option> children.
function egpSyncSelectValues(){
  Object.entries(EGP_IDS).forEach(([key,id])=>{
    const el=document.getElementById(id);
    if(el) el.value=egpFilters[key];
  });
}

function egpMatchesFilters(e){
  if(egpFilters.bd&&normStr(e.bd_pm)!==egpFilters.bd) return false;
  if(egpFilters.month&&egpMonthKey(e.eng_date)!==egpFilters.month) return false;
  if(egpFilters.company&&normStr(e.company)!==egpFilters.company) return false;
  if(egpFilters.objective&&normStr(e.engagement_objective)!==egpFilters.objective) return false;
  if(egpFilters.type&&normStr(e.engagement_type)!==egpFilters.type) return false;
  return true;
}

function renderEngagementPivotTable(){
  egpSyncSelectValues();

  const tbl=document.getElementById('egp-table');
  if(!tbl) return;

  const filtered=DATA.engagements.filter(egpMatchesFilters);
  const types=orderByLookup([...new Set(filtered.map(e=>normStr(e.engagement_type)).filter(Boolean))],LOOKUPS.engagementTypes);
  const outcomes=orderByLookup([...new Set(filtered.map(e=>normStr(e.engagement_outcome)).filter(Boolean))],LOOKUPS.engagementOutcome);

  if(!types.length||!outcomes.length){
    tbl.querySelector('thead').innerHTML='';
    tbl.querySelector('tbody').innerHTML='<tr><td class="empty-state">No engagements match these filters.</td></tr>';
    return;
  }

  tbl.querySelector('thead').innerHTML=`<tr><th>Outcome / Stakeholder Type</th>${types.map(t=>`<th class="num">${esc(t)}</th>`).join('')}<th class="num">Grand Total</th></tr>`;

  const colTotals=types.map(()=>0);
  let grandTotal=0;
  const rowsHtml=outcomes.map(outcome=>{
    const outcomeEngs=filtered.filter(e=>normStr(e.engagement_outcome)===outcome);
    if(!outcomeEngs.length) return '';
    const stakeholderTypes=orderByLookup([...new Set(outcomeEngs.map(e=>normStr(e.stakeholder_type)).filter(Boolean))],LOOKUPS.stakeholderType);
    const outcomeCounts=types.map(t=>outcomeEngs.filter(e=>normStr(e.engagement_type)===t).length);
    const outcomeTotal=outcomeCounts.reduce((a,b)=>a+b,0);
    outcomeCounts.forEach((n,i)=>colTotals[i]+=n);
    grandTotal+=outcomeTotal;

    const groupRow=`<tr class="pivot-group open" onclick="togglePivotGroup(this)"><td><span class="caret">▶</span> ${esc(outcome)}</td>${outcomeCounts.map(n=>`<td class="num">${n||''}</td>`).join('')}<td class="num">${outcomeTotal}</td></tr>`;
    const childRows=stakeholderTypes.map(st=>{
      const stEngs=outcomeEngs.filter(e=>normStr(e.stakeholder_type)===st);
      const stCounts=types.map(t=>stEngs.filter(e=>normStr(e.engagement_type)===t).length);
      const stTotal=stCounts.reduce((a,b)=>a+b,0);
      return `<tr class="pivot-child show"><td>${esc(st)}</td>${stCounts.map(n=>`<td class="num">${n||''}</td>`).join('')}<td class="num">${stTotal}</td></tr>`;
    }).join('');
    return groupRow+childRows;
  }).join('');

  const totalRow=`<tr class="pivot-total"><td>Grand Total</td>${colTotals.map(n=>`<td class="num">${n}</td>`).join('')}<td class="num">${grandTotal}</td></tr>`;
  tbl.querySelector('tbody').innerHTML=rowsHtml+totalRow;
}

function renderEngagements(){
  populateEngagementPivotFilters();
  renderEngagementPivotTable();
  const thisWeek=isoWeek(new Date().toISOString());
  const lastWeekStart=mondayOf(new Date().toISOString()); lastWeekStart.setDate(lastWeekStart.getDate()-7);
  const lastWeek=isoWeek(lastWeekStart.toISOString());
  const engThisWeek=DATA.engagements.filter(e=>isoWeek(e.eng_date)===thisWeek).length;
  const engLastWeek=DATA.engagements.filter(e=>isoWeek(e.eng_date)===lastWeek).length;
  const wowDelta=engThisWeek-engLastWeek;
  const pending=DATA.engagements.filter(e=>String(e.follow_up_done||'').toLowerCase()!=='yes').length;
  const ledToRfp=DATA.engagements.filter(e=>String(e.engagement_outcome||'').toLowerCase().includes('led to rfp')).length;
  renderKpiCards('eg-kpis',[
    {label:'Total',value:fmtNum(DATA.engagements.length),sub:'logged touchpoints',accent:'sky'},
    {label:'This week',value:fmtNum(engThisWeek),sub:`${wowDelta>0?'▲':wowDelta<0?'▼':'–'} ${Math.abs(wowDelta)} vs last week`,accent:'amber'},
    {label:'Last week',value:fmtNum(engLastWeek),sub:weekRangeLabel(lastWeekStart),accent:'sky'},
    {label:'Follow-up pending',value:fmtNum(pending),sub:'needs attention',accent:'coral'},
    {label:'Led to RFP',value:fmtNum(ledToRfp),sub:'conversion events',accent:'green'}
  ]);

  // Weekly bar chart — labelled by week-start date, not raw ISO week strings. Click a bar to drill into that week.
  const egWeekSeries=weeklySeries(DATA.engagements.map(e=>e.eng_date),12);
  destroyChart('egWeekly');
  const ctx=document.getElementById('eg-weekly-chart').getContext('2d');
  CHARTS.egWeekly=new Chart(ctx,{type:'bar',data:{labels:egWeekSeries.map(s=>s.label),datasets:[{label:'Engagements',data:egWeekSeries.map(s=>s.count),backgroundColor:egWeekSeries.map(s=>egFilters.week===s.key?COLORS.skyDeep:COLORS.sky),borderRadius:4,maxBarThickness:36}]},options:{responsive:true,maintainAspectRatio:false,onClick:(evt,elements)=>{if(!elements.length)return;const s=egWeekSeries[elements[0].index];if(s)filterEngByWeek(s.key);},onHover:(evt,elements)=>{if(evt.native?.target) evt.native.target.style.cursor=elements.length?'pointer':'default';},plugins:{legend:{display:false},tooltip:{callbacks:{title:items=>egWeekSeries[items[0].dataIndex]?.rangeLabel||''}}},scales:{y:{beginAtZero:true,ticks:{precision:0}},x:{grid:{display:false},title:{display:true,text:'Week starting'}}}}});

  // Company-wise breakdown — click a company to drill into its engagements
  const egCoMap={};
  DATA.engagements.forEach(e=>{const co=(e.company||'Unknown').trim();egCoMap[co]=(egCoMap[co]||0)+1;});
  const egCoTop=Object.entries(egCoMap).sort((a,b)=>b[1]-a[1]);
  const egCoMax=Math.max(1,...egCoTop.map(([,n])=>n));
  document.getElementById('eg-company-breakdown').innerHTML=egCoTop.length?egCoTop.map(([c,n])=>`
    <div class="funnel-row" style="grid-template-columns:120px 1fr 30px;cursor:pointer;${egFilters.company===c?'background:var(--sky-pale);border-radius:6px;':''}" onclick="filterEngByCompany('${esc(c)}')" title="Show only ${esc(c)}">
      <div class="funnel-label" style="overflow:hidden;text-overflow:ellipsis;">${esc(c)}</div>
      <div class="funnel-track"><div class="funnel-fill" style="width:${(n/egCoMax*100).toFixed(0)}%;background:${COLORS.sky};opacity:.85;"></div></div>
      <div class="funnel-val">${n}</div>
    </div>`).join('')
    :'<div class="empty-state">No engagements logged yet.</div>';

  // Active week-filter chip (set by clicking a bar in the weekly chart above)
  const weekChipEl=document.getElementById('eg-week-chip');
  if(weekChipEl){
    if(egFilters.week){
      const wkInfo=egWeekSeries.find(s=>s.key===egFilters.week);
      weekChipEl.innerHTML=`<span class="tag tag-sky" style="cursor:pointer;" onclick="clearEgWeekFilter()" title="Clear week filter">📅 ${esc(wkInfo?wkInfo.rangeLabel:egFilters.week)} ✕</span>`;
    } else {
      weekChipEl.innerHTML='';
    }
  }

  // Company filter alphabetically — value is always set FROM egFilters (state is the single
  // source of truth), never read back from the DOM, so Clear can't desync from what's shown.
  const egCoSel=document.getElementById('eg-filter-company');
  egCoSel.innerHTML='<option value="">All companies</option>'+
    [...new Set(DATA.engagements.map(e=>e.company).filter(Boolean))].sort().map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  egCoSel.value=egFilters.company;

  const typeSel=document.getElementById('eg-filter-type');
  typeSel.innerHTML='<option value="">All types</option>'+
    LOOKUPS.engagementTypes.map(t=>`<option value="${esc(t)}">${esc(t)}</option>`).join('');
  typeSel.value=egFilters.type;

  const bdSel=document.getElementById('eg-filter-bd');
  if(bdSel){
    const bdNames=[...new Set(DATA.engagements.map(e=>normStr(e.bd_pm)).filter(Boolean))].sort();
    bdSel.innerHTML='<option value="">All BD/PM</option>'+bdNames.map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join('');
    bdSel.value=egFilters.bd;
  }

  const accSel=document.getElementById('eg-filter-accompanied');
  if(accSel){
    const accNames=[...new Set(DATA.engagements.flatMap(e=>msSelectedFromValue(e.accompanied_by)))].sort();
    accSel.innerHTML='<option value="">Accompanied by: any</option>'+accNames.map(n=>`<option value="${esc(n)}">${esc(n)}</option>`).join('');
    accSel.value=egFilters.accompanied;
  }

  let rows=DATA.engagements.filter(e=>{
    if(egFilters.type&&e.engagement_type!==egFilters.type) return false;
    if(egFilters.company&&e.company!==egFilters.company) return false;
    if(egFilters.bd&&normStr(e.bd_pm)!==egFilters.bd) return false;
    if(egFilters.accompanied&&!msSelectedFromValue(e.accompanied_by).includes(egFilters.accompanied)) return false;
    if(egFilters.week&&isoWeek(e.eng_date)!==egFilters.week) return false;
    if(egFilters.followup==='yes'&&String(e.follow_up_done||'').toLowerCase()!=='yes') return false;
    if(egFilters.followup==='no'&&String(e.follow_up_done||'').toLowerCase()==='yes') return false;
    return true;
  });
  rows=applySort(rows,'engagements','company');

  const cols=[
    {key:'eng_date',label:'Date'},{key:'client_name',label:'Client'},{key:'company',label:'Company'},
    {key:'designation',label:'Designation'},{key:'bd_pm',label:'BD / PM'},{key:'accompanied_by',label:'Accompanied By'},
    {key:'engagement_type',label:'Type'},
    {key:'stakeholder_type',label:'Stakeholder'},
    {key:'engagement_objective',label:'Objective'},{key:'engagement_outcome',label:'Outcome'},
    {key:'discussion_points',label:'Discussion'},{key:'cta_next_step',label:'Next Step'},
    {key:'cta_due_date',label:'CTA Due'},{key:'cta_owner',label:'CTA Owner'},
    {key:'follow_up_done',label:'Follow-up'},{key:'_actions',label:''}
  ];
  const tbl=document.getElementById('eg-table');
  makeSortableHeaders(tbl.querySelector('thead'),cols,'engagements',renderEngagements);
  tbl.querySelector('tbody').innerHTML=rows.length?rows.map(e=>`
    <tr>
      <td style="white-space:nowrap">${fmtDate(e.eng_date)}</td>
      <td>${esc(e.client_name||'')}</td>
      <td><strong>${esc(e.company||'')}</strong></td>
      <td>${esc(e.designation||'')}</td>
      <td>${esc(e.bd_pm||'')}</td>
      <td>${esc(msSelectedFromValue(e.accompanied_by).join(', '))}</td>
      <td><span class="${engTypeClass(e.engagement_type)}">${esc(e.engagement_type||'')}</span></td>
      <td>${esc(e.stakeholder_type||'')}</td>
      <td>${esc(e.engagement_objective||'')}</td>
      <td>${esc(e.engagement_outcome||'')}</td>
      <td style="max-width:180px;white-space:normal">${esc(e.discussion_points||'')}</td>
      <td style="max-width:160px;white-space:normal">${esc(e.cta_next_step||'')}</td>
      <td style="white-space:nowrap">${fmtDate(e.cta_due_date)}</td>
      <td>${esc(e.cta_owner||'')}</td>
      <td><span class="${followUpClass(e.follow_up_done)}">${esc(e.follow_up_done||'No')}</span></td>
      <td class="td-actions">${canEdit()?`<button class="btn btn-edit btn-sm" onclick="openEditModal('engagements','${esc(e.id)}')">Edit</button>`:''}</td>
    </tr>`).join('')
    :'<tr><td colspan="17" class="empty-state">No engagements match these filters.</td></tr>';
  document.getElementById('eg-count').textContent=`${rows.length} of ${DATA.engagements.length}`;
}

/* ---------- 16. RENDER: COMPANIES ---------- */
let coFilters={search:'',onboarding:''};

function normCo(n){ return (n||'').trim().toLowerCase(); }
function computeCompanyRollup(){
  const map={};
  function ensure(name){
    const norm=normCo(name); if(!norm) return null;
    if(!map[norm]) map[norm]={company:name.trim(),onboarding_status:null,targetRevenue:null,totalClients:0,activeClients:0,openOpps:0,pipelineValue:0,weightedValue:0,wins:0,winValue:0,toWin:0,toWinValue:0,losses:0,lossValue:0,engagements:0,lastEngagement:null};
    return map[norm];
  }
  DATA.companies.forEach(c=>{const r=ensure(c.company);if(r){r.onboarding_status=c.onboarding_status;r.targetRevenue=c.target_revenue;}});
  DATA.clients.forEach(c=>{const r=ensure(c.company);if(!r)return;r.totalClients++;if(['active business','active engagement'].includes(String(c.status||'').toLowerCase()))r.activeClients++;});
  DATA.opportunities.forEach(o=>{
    const r=ensure(o.company);if(!r)return;
    const val=Number(o.estimated_value_usd)||0;
    if(isOpenStage(o.opportunity_status)){
      r.openOpps++;r.pipelineValue+=val;r.weightedValue+=Number(o.probability_weighted_value)||0;
      if(STAGE_NUM(o.opportunity_status)===6){r.toWin++;r.toWinValue+=val;}
    }
    if(isWinStage(o.opportunity_status)){r.wins++;r.winValue+=val;}
    if(isLossStage(o.opportunity_status)){r.losses++;r.lossValue+=val;}
  });
  DATA.engagements.forEach(e=>{const r=ensure(e.company);if(!r)return;r.engagements++;if(!r.lastEngagement||new Date(e.eng_date)>new Date(r.lastEngagement))r.lastEngagement=e.eng_date;});
  return Object.values(map).sort((a,b)=>a.company.localeCompare(b.company));
}

// Read-only summary — editing these scores happens through a company's Edit modal, like every other field.
let coHealthFilters={search:'',onboarding:''};

// Autocomplete dropdown on the health matrix's company search — typing still filters the
// table live (via renderCompanyHealthMatrix below), this just adds quick-pick suggestions.
function coHealthSearchOptions(query){
  const q=(query||'').toLowerCase();
  const names=[...new Set(DATA.companies.map(c=>c.company).filter(Boolean))].sort();
  return names.filter(n=>n.toLowerCase().includes(q));
}
function coHealthSearchOpen(inp){
  const list=document.getElementById('co-health-search-list');
  if(!list) return;
  list.innerHTML=coHealthSearchOptions(inp.value).map(n=>`<div class="combo-opt" onmousedown="coHealthSearchPick('${esc(n)}')">${esc(n)}</div>`).join('')||'<div class="combo-opt no-match">No matches</div>';
  list.classList.add('open');
}
function coHealthSearchFilter(inp){
  coHealthFilters.search=inp.value;
  renderCompanyHealthMatrix();
  coHealthSearchOpen(inp);
}
function coHealthSearchPick(name){
  const inp=document.getElementById('co-health-search');
  if(inp) inp.value=name;
  coHealthFilters.search=name;
  renderCompanyHealthMatrix();
  const list=document.getElementById('co-health-search-list');
  if(list) list.classList.remove('open');
}

function renderCompanyHealthMatrix(){
  const tbl=document.getElementById('co-health-table');
  if(!tbl) return;
  // Merge any duplicate SharePoint rows for the same company (e.g. leftover "Abbvie"/"AbbVie"
  // list items) into one column, preferring whichever row has a value for each measure.
  const byKey={};
  DATA.companies.filter(c=>c.company).forEach(c=>{
    const k=normCo(c.company);
    if(!byKey[k]) byKey[k]={...c};
    else Object.keys(c).forEach(f=>{ if(byKey[k][f]==null&&c[f]!=null) byKey[k][f]=c[f]; });
  });
  let companies=Object.values(byKey).sort((a,b)=>a.company.localeCompare(b.company));

  const onbSel=document.getElementById('co-health-filter-onboarding');
  if(onbSel){
    const onbStatuses=[...new Set(DATA.companies.map(c=>c.onboarding_status).filter(Boolean))].sort();
    onbSel.innerHTML='<option value="">All onboarding statuses</option>'+onbStatuses.map(s=>`<option value="${esc(s)}">${esc(s)}</option>`).join('');
    onbSel.value=coHealthFilters.onboarding;
  }
  companies=companies.filter(c=>{
    if(coHealthFilters.onboarding&&c.onboarding_status!==coHealthFilters.onboarding) return false;
    if(coHealthFilters.search&&!c.company.toLowerCase().includes(coHealthFilters.search.toLowerCase())) return false;
    return true;
  });
  const countEl=document.getElementById('co-health-count');
  if(countEl) countEl.textContent=`${companies.length} of ${Object.values(byKey).length}`;

  const measures=[
    {key:'overall_budget_potential',label:'Overall Budget Potential',fmt:v=>v?`${v}/5`:''},
    {key:'overall_client_relationship',label:'Overall Client Relationship',fmt:v=>v?`${v}/5`:''},
    {key:'client_perception',label:'Client Perception',fmt:v=>v||''},
    {key:'team_satisfaction',label:'Team Satisfaction',fmt:v=>v?`${v}/5`:''},
    {key:'degree_of_innovation',label:'Degree of Innovation',fmt:v=>v?`${v}/5`:''}
  ];
  tbl.querySelector('thead').innerHTML=`<tr><th>Company</th>${measures.map(m=>`<th>${esc(m.label)}</th>`).join('')}</tr>`;
  tbl.querySelector('tbody').innerHTML=companies.length?companies.map(c=>`
    <tr><td>${esc(c.company)}</td>${measures.map(m=>`<td class="num">${esc(String(m.fmt(c[m.key])))}</td>`).join('')}</tr>`).join('')
    :'<tr><td class="empty-state">No companies match this search.</td></tr>';
}

function renderCompanies(){
  const rollup=computeCompanyRollup();
  const withBiz=rollup.filter(r=>r.activeClients>0).length;
  renderKpiCards('co-kpis',[
    {label:'Companies',value:fmtNum(rollup.length),sub:`${withBiz} with active business`,accent:'sky'},
    {label:'Combined pipeline',value:fmtMoney(rollup.reduce((s,r)=>s+r.pipelineValue,0)),sub:'open opportunities',accent:'sky'},
    {label:'Total wins',value:fmtNum(rollup.reduce((s,r)=>s+r.wins,0)),sub:'closed Win',accent:'green'},
    {label:'Onboarding on file',value:fmtNum(DATA.companies.filter(c=>c.onboarding_status).length),sub:`of ${DATA.companies.length} listed`,accent:'amber'}
  ]);

  // Onboarding filter
  const onbSel=document.getElementById('co-filter-onboarding');
  const onbStatuses=[...new Set(DATA.companies.map(c=>c.onboarding_status).filter(Boolean))].sort();
  onbSel.innerHTML='<option value="">All onboarding statuses</option>'+onbStatuses.map(s=>`<option value="${esc(s)}">${esc(s)}</option>`).join('');
  onbSel.value=coFilters.onboarding;

  renderCompanyHealthMatrix();

  let rows=rollup.filter(r=>{
    if(coFilters.onboarding&&r.onboarding_status!==coFilters.onboarding) return false;
    if(coFilters.search&&!r.company.toLowerCase().includes(coFilters.search.toLowerCase())) return false;
    return true;
  });
  rows=applySort(rows,'companies','company');

  const cols=[
    {key:'company',label:'Company'},{key:'onboarding_status',label:'Onboarding'},
    {key:'targetRevenue',label:'Target Revenue'},
    {key:'totalClients',label:'Clients'},{key:'activeClients',label:'Active'},
    {key:'openOpps',label:'Open Opps'},{key:'pipelineValue',label:'Pipeline'},
    {key:'weightedValue',label:'Weighted'},{key:'wins',label:'Wins'},
    {key:'engagements',label:'Engagements'},{key:'lastEngagement',label:'Last Engagement'},
    {key:'_actions',label:''}
  ];
  const tbl=document.getElementById('co-table');
  makeSortableHeaders(tbl.querySelector('thead'),cols,'companies',renderCompanies);
  tbl.querySelector('tbody').innerHTML=rows.length?rows.map(r=>{
    const noEng90=r.lastEngagement&&daysAgo(r.lastEngagement)>90;
    return `<tr${noEng90?' style="background:rgba(200,84,61,0.04)"':''}>
      <td><strong>${esc(r.company)}</strong></td>
      <td>${r.onboarding_status?`<span class="${onboardingClass(r.onboarding_status)}">${esc(r.onboarding_status)}</span>`:'<span class="tag tag-grey">Not on file</span>'}</td>
      <td class="num">${r.targetRevenue?fmtMoney(r.targetRevenue):''}</td>
      <td class="num">${fmtNum(r.totalClients)}</td>
      <td class="num">${fmtNum(r.activeClients)}</td>
      <td class="num">${fmtNum(r.openOpps)}</td>
      <td class="num">${fmtMoney(r.pipelineValue)}</td>
      <td class="num">${fmtMoney(r.weightedValue)}</td>
      <td class="num">${fmtNum(r.wins)}</td>
      <td class="num">${fmtNum(r.engagements)}</td>
      <td${noEng90?' style="color:var(--coral)"':''}>${r.lastEngagement?fmtDate(r.lastEngagement):''}</td>
      <td class="td-actions">${canEdit()?`<button class="btn btn-edit btn-sm" onclick="openCoEditModal('${esc(r.company)}')">Edit</button>`:''}</td>
    </tr>`}).join('')
    :'<tr><td colspan="12" class="empty-state">No companies match this search.</td></tr>';
  document.getElementById('co-count').textContent=`${rows.length} of ${rollup.length}`;
}

/* ---------- 17. COMBOBOX HELPER ---------- */
// Builds a searchable dropdown with optional "Add new" capability.
// opts: {id, value, options[], placeholder, allowNew, required}
function buildCombo(opts){
  const {id,value='',options=[],placeholder='',allowNew=false,required=false}=opts;
  return `<div class="combo-wrap" data-combo="${id}">
    <input type="text" id="${id}" name="${id}" placeholder="${esc(placeholder)}"
      value="${esc(value)}" autocomplete="off" ${required?'required':''}
      oninput="comboFilter(this)" onfocus="comboOpen(this)" onblur="comboBlur(this,300)">
    <div class="combo-list" id="${id}-list">
      ${options.map(o=>`<div class="combo-opt" onmousedown="comboPick(this,'${id}','${esc(o)}')">${esc(o)}</div>`).join('')}
      ${allowNew?`<div class="combo-opt add-new" onmousedown="comboAddNew(event,'${id}')">Add new…</div>`:''}
    </div>
  </div>`;
}

/* ---- Multi-select combobox (chips + checkbox list) — e.g. "Accompanied by" ----
   Value in/out is a single "; "-joined string, so it reads/writes through the same
   generic el.value plumbing every other field in the modal already uses (validation,
   payload building, edit-mode prefill) without any special-casing there. */
function msSelectedFromValue(value){
  return String(value||'').split(';').map(s=>s.trim()).filter(Boolean);
}
function buildMultiselect(opts){
  const {id,value='',options=[],placeholder='Select…',required=false}=opts;
  const selected=msSelectedFromValue(value);
  return `<div class="multiselect" id="${id}-wrap">
    <div class="multiselect-box" onclick="msToggle('${id}')">${msChipsHtml(id,selected,placeholder)}</div>
    <input type="hidden" id="${id}" name="${id}" value="${esc(selected.join('; '))}" ${required?'required':''}>
    <div class="multiselect-list" id="${id}-list">
      ${options.map(o=>`<label class="ms-opt"><input type="checkbox" value="${esc(o)}"${selected.includes(o)?' checked':''} onchange="msToggleOption('${id}','${esc(o)}')"> ${esc(o)}</label>`).join('')}
    </div>
  </div>`;
}
function msChipsHtml(id,selected,placeholder){
  if(!selected.length) return `<span class="ms-placeholder">${esc(placeholder)}</span>`;
  return selected.map(v=>`<span class="ms-chip">${esc(v)}<button type="button" class="ms-chip-x" onclick="msRemove(event,'${id}','${esc(v)}')">×</button></span>`).join('');
}
function msToggle(id){
  const list=document.getElementById(id+'-list');
  if(list) list.classList.toggle('open');
}
function msSetSelected(id,selected){
  const hidden=document.getElementById(id);
  if(hidden) hidden.value=selected.join('; ');
  const box=document.querySelector(`#${id}-wrap .multiselect-box`);
  if(box) box.innerHTML=msChipsHtml(id,selected,'Select…');
}
function msToggleOption(id,val){
  const hidden=document.getElementById(id);
  let selected=msSelectedFromValue(hidden?hidden.value:'');
  selected=selected.includes(val)?selected.filter(v=>v!==val):[...selected,val];
  msSetSelected(id,selected);
}
function msRemove(evt,id,val){
  evt.stopPropagation();
  const hidden=document.getElementById(id);
  const selected=msSelectedFromValue(hidden?hidden.value:'').filter(v=>v!==val);
  msSetSelected(id,selected);
  const cb=document.querySelector(`#${id}-list input[type=checkbox][value="${CSS.escape(val)}"]`);
  if(cb) cb.checked=false;
}
// Closes any open multiselect list when the user clicks outside of it.
document.addEventListener('click',e=>{
  document.querySelectorAll('.multiselect-list.open').forEach(list=>{
    const wrap=list.closest('.multiselect');
    if(wrap&&!wrap.contains(e.target)) list.classList.remove('open');
  });
});

// Combobox interaction functions
function comboFilter(inp){
  const q=inp.value.toLowerCase();
  const list=document.getElementById(inp.id+'-list');
  if(!list) return;
  list.classList.add('open');
  list.querySelectorAll('.combo-opt:not(.add-new)').forEach(opt=>{
    opt.style.display=opt.textContent.toLowerCase().includes(q)?'':'none';
  });
  // Show no-match message
  const visible=[...list.querySelectorAll('.combo-opt:not(.add-new):not(.no-match)')].filter(o=>o.style.display!=='none');
  let nm=list.querySelector('.no-match');
  if(!visible.length&&!list.querySelector('.add-new')){
    if(!nm){nm=document.createElement('div');nm.className='combo-opt no-match';nm.textContent='No matches';list.appendChild(nm);}
  }else if(nm) nm.remove();
}
function comboOpen(inp){
  const list=document.getElementById(inp.id+'-list');
  if(list) list.classList.add('open');
}
function comboBlur(inp,delay){
  setTimeout(()=>{
    const list=document.getElementById(inp.id+'-list');if(list)list.classList.remove('open');
    if(inp.id==='modal-eng-client') autofillEngClientFields(inp.value);
    if(inp.id==='modal-opp-client-name') autofillOppClientFields(inp.value);
  },delay);
}
function comboPick(el,id,val){
  const inp=document.getElementById(id); if(inp) inp.value=val;
  const list=document.getElementById(id+'-list'); if(list) list.classList.remove('open');
  if(id==='modal-eng-client') autofillEngClientFields(val);
  if(id==='modal-opp-client-name') autofillOppClientFields(val);
}

// Looks up the picked client in the Clients master list and fills in Designation / BD-PM
// from their record, since the person logging the engagement rarely remembers this by heart.
// Clears and re-enables a client combo (and an optional dependent field, e.g.
// engagement's designation) after the company changes — the previous pick no longer
// applies, since it belonged to the old company's client list.
function resetClientLock(clientId,extraId){
  const clientEl=document.getElementById(clientId);
  if(clientEl){clientEl.readOnly=false;clientEl.value='';}
  if(extraId){
    const extraEl=document.getElementById(extraId);
    if(extraEl){extraEl.readOnly=false;extraEl.value='';}
  }
}

function autofillEngClientFields(clientName){
  const coEl=document.getElementById('modal-eng-company');
  const desigEl=document.getElementById('designation');
  const bdEl=document.getElementById('bd_pm');
  const clientEl=document.getElementById('modal-eng-client');
  if(!coEl||!clientName) return;
  const match=DATA.clients.find(c=>normCo(c.company)===normCo(coEl.value)&&normStr(c.client_name).toLowerCase()===normStr(clientName).toLowerCase());
  if(!match) return;
  // Designation only locks once it's actually been filled in — if the client record has
  // none on file, leave it open so the user can type it in themselves (it's required).
  if(desigEl&&match.designation){desigEl.value=match.designation;desigEl.readOnly=true;}
  if(bdEl&&match.assigned_bd) bdEl.value=match.assigned_bd;
  if(clientEl) clientEl.readOnly=true;
}

// Same idea for the opportunity form: once a client's picked, fill BD owner from their
// assigned BD in the Clients master list, since the opportunity BD is usually whoever
// already owns that client relationship.
function autofillOppClientFields(clientName){
  const coEl=document.getElementById('modal-opp-company');
  const bdEl=document.getElementById('bd_owner');
  const clientEl=document.getElementById('modal-opp-client-name');
  if(!coEl||!clientName) return;
  const match=DATA.clients.find(c=>normCo(c.company)===normCo(coEl.value)&&normStr(c.client_name).toLowerCase()===normStr(clientName).toLowerCase());
  if(!match) return;
  if(bdEl&&match.assigned_bd) bdEl.value=match.assigned_bd;
  if(clientEl) clientEl.readOnly=true;
}
async function comboAddNew(evt,id){
  evt.preventDefault();
  const inp=document.getElementById(id);
  const val=inp?inp.value.trim():'';
  const name=prompt(`Add new entry for this field:`+(val?` (pre-filled: "${val}")`:''),(val||''));
  if(!name||!name.trim()) return;
  inp.value=name.trim();
  // Unlike comboPick (a normal suggestion click), this never closed the dropdown —
  // it stayed open showing the old, now-irrelevant option list, which both looked like
  // the new entry hadn't taken and made it easy to accidentally click a stale suggestion
  // right afterward, silently overwriting the value that was just typed in.
  const list=document.getElementById(id+'-list');
  if(list) list.classList.remove('open');
}

function updateModalOppClientOptions(company){
  const listEl=document.getElementById('modal-opp-client-name-list');
  if(!listEl) return;
  const clients=DATA.clients.filter(c=>normCo(c.company)===normCo(company)).map(c=>c.client_name).filter(Boolean).sort();
  const existing=listEl.querySelectorAll('.combo-opt:not(.add-new)');
  existing.forEach(e=>e.remove());
  const addNew=listEl.querySelector('.add-new');
  clients.forEach(cl=>{
    const div=document.createElement('div'); div.className='combo-opt';
    div.textContent=cl; div.onmousedown=()=>comboPick(div,'modal-opp-client-name',cl);
    listEl.insertBefore(div,addNew||null);
  });
}

// Region detail is a manual text field now (was a combo dropdown) — this just toggles
// its visibility and swaps its label to match whichever type was picked, since "Regional"
// and "Country" each need their own heading rather than one generic "Region detail".
function updateRegionDetailUI(type){
  const wrap=document.getElementById('region-detail-wrap');
  if(!wrap) return;
  wrap.style.display=(type==='Regional'||type==='Country')?'':'none';
  const label=wrap.querySelector('label');
  if(label) label.textContent=type==='Country'?'Country detail':'Region detail';
}

/* ---------- 18. STAGE HISTORY HELPERS ---------- */
function getHistory(opp){
  try{ return JSON.parse(opp.stage_history||'[]')||[]; }catch{ return []; }
}
function addHistoryEntry(existingJson, newStage, userEmail){
  const hist=(() => { try{ return JSON.parse(existingJson||'[]'); }catch{ return []; } })();
  hist.push({stage:newStage,date:new Date().toISOString().slice(0,10),by:userEmail||'unknown'});
  return JSON.stringify(hist);
}
function renderHistoryTimeline(hist){
  if(!hist||!hist.length) return '<div style="color:var(--ink-faint);font-size:12px;">No stage history recorded.</div>';
  return hist.map((h,i)=>{
    const n=STAGE_NUM(h.stage)||1;
    const clr=STAGE_COLORS[n]||COLORS.sky;
    return `<div class="history-step">
      <div class="history-dot" style="background:${clr}">${n}</div>
      <div class="history-info">
        <div class="history-stage">${esc((h.stage||'').replace(/^\d+\s/,''))}</div>
        <div class="history-meta">${esc(h.date)} · ${esc(h.by)}</div>
      </div>
    </div>`;
  }).join('');
}

/* ---------- 19. MODAL: ADD ---------- */
const ADD_CONFIGS={
  clients:{title:'Add client',kind:'clients',fields:[
    {name:'modal-company',label:'Company',type:'select',required:true,opts:()=>uniqueCompanies()},
    {name:'modal-client-name',label:'Client name',type:'text',required:true},
    {name:'designation',label:'Designation',type:'text',required:true},
    {name:'department',label:'Department',combo:true,allowNew:true,required:true,comboOpts:()=>uniqueDepartments()},
    {name:'therapy_area',label:'Therapy area',combo:true,allowNew:true,required:true,comboOpts:()=>uniqueTherapyAreas()},
    {name:'region_type',label:'Region',type:'select',required:true,opts:['Global','Regional','Country']},
    {name:'region_detail',label:'Region detail',type:'text',wrapId:'region-detail-wrap'},
    {name:'email',label:'Email',type:'text',required:true},
    {name:'phone',label:'Phone',type:'text',required:true},
    {name:'linkedin_url',label:'LinkedIn URL',type:'text'},
    {name:'status',label:'Status',type:'select',opts:LOOKUPS.clientStatus,required:true},
    {name:'priority',label:'Priority',type:'select',opts:LOOKUPS.priority,required:true},
    {name:'assigned_bd',label:'Assigned BD',combo:true,comboOpts:()=>EMPLOYEES,required:true},
    {name:'notes',label:'Notes',type:'textarea',full:true}
  ]},
  opportunities:{title:'Add opportunity',kind:'opportunities',fields:[
    {name:'modal-opp-company',label:'Company',type:'select',required:true,opts:()=>uniqueCompanies()},
    {name:'modal-opp-client-name',label:'Client name',combo:true,comboOpts:()=>[],comboDepends:'modal-opp-company',required:true},
    {name:'opportunity',label:'Opportunity',type:'text',required:true,full:true},
    {name:'opportunity_status',label:'Stage',type:'select',opts:LOOKUPS.opportunityStatus,required:true},
    {name:'bd_owner',label:'BD owner',combo:true,comboOpts:()=>EMPLOYEES,required:true},
    {name:'supporting_role',label:'Supporting role',type:'multiselect',opts:()=>EMPLOYEES,placeholder:'Supporting team members…'},
    {name:'estimated_value_usd',label:'Estimated value (USD)',type:'number',required:true,min:0},
    {name:'probability_pct',label:'Probability (%)',type:'select',opts:['10','25','75','100']},
    {name:'identified_date',label:'Opportunity identified date',type:'date'},
    {name:'discussion_date',label:'Discussion date',type:'date',required:true},
    {name:'pitch_date',label:'Pitch date',type:'date'},
    {name:'proposal_submission_date',label:'Proposal submission',type:'date'},
    {name:'expected_close_date',label:'Expected close',type:'date',required:true},
    {name:'notes',label:'Notes',type:'textarea',full:true}
  ]},
  engagements:{title:'Log engagement',kind:'engagements',fields:[
    {name:'eng_date',label:'Date',type:'date',required:true},
    {name:'modal-eng-company',label:'Company',type:'select',required:true,opts:()=>uniqueCompanies()},
    {name:'modal-eng-client',label:'Client name',combo:true,comboOpts:()=>[],comboDepends:'modal-eng-company',required:true},
    {name:'designation',label:'Designation',type:'text',required:true},
    {name:'bd_pm',label:'BD / PM',combo:true,comboOpts:()=>EMPLOYEES,required:true},
    {name:'accompanied_by',label:'Accompanied by',type:'multiselect',opts:()=>EMPLOYEES,placeholder:'Other team members at this meeting…'},
    {name:'engagement_type',label:'Engagement type',type:'select',opts:LOOKUPS.engagementTypes,required:true},
    {name:'stakeholder_type',label:'Stakeholder type',type:'select',opts:LOOKUPS.stakeholderType,required:true},
    {name:'engagement_objective',label:'Objective',type:'select',opts:LOOKUPS.engagementObjective,required:true},
    {name:'engagement_outcome',label:'Outcome',type:'select',opts:LOOKUPS.engagementOutcome,required:true},
    {name:'discussion_points',label:'Discussion points',type:'textarea',full:true,required:true},
    {name:'cta_next_step',label:'CTA / next step',type:'textarea',full:true,required:true},
    {name:'cta_due_date',label:'CTA due date',type:'date',required:true},
    {name:'cta_owner',label:'CTA owner',combo:true,comboOpts:()=>EMPLOYEES,required:true},
    {name:'follow_up_done',label:'Follow-up done?',type:'select',opts:LOOKUPS.followUp,required:true}
  ]},
  companies:{title:'Add company',kind:'companies',fields:[
    {name:'company',label:'Company name',type:'text',required:true},
    {name:'onboarding_status',label:'Onboarding status',type:'select',opts:LOOKUPS.onboardingStatus},
    {name:'target_revenue',label:'Target Revenue (USD)',type:'number',min:0},
    {name:'overall_budget_potential',label:'Overall Budget Potential',type:'number',min:1,max:5},
    {name:'overall_client_relationship',label:'Overall Client Relationship',type:'number',min:1,max:5},
    {name:'client_perception',label:'Client Perception',type:'select',opts:LOOKUPS.clientPerception},
    {name:'team_satisfaction',label:'Team Satisfaction',type:'number',min:1,max:5},
    {name:'degree_of_innovation',label:'Degree of Innovation',type:'number',min:1,max:5},
    {name:'notes',label:'Notes',type:'textarea',full:true}
  ]}
};

function renderModalField(f,val=''){
  const req=f.required?`<span class="req"> *</span>`:'';
  const wrap=f.full?'full':'';
  let inp='';
  if(f.combo){
    const id=f.comboId||f.name;
    inp=buildCombo({id,value:val,options:f.comboOpts(),placeholder:`Type to search…`,allowNew:!!f.allowNew,required:f.required});
  } else if(f.type==='select'){
    const opts=typeof f.opts==='function'?f.opts():f.opts;
    // Match case/whitespace-insensitively — SharePoint data entered outside this exact
    // dropdown (imports, manual list edits) can differ in casing from the option list.
    // If the stored value still matches nothing, keep it as its own selected option
    // instead of silently defaulting to blank — that would submit null and wipe out
    // real data on save without the user ever noticing.
    const norm=s=>String(s).trim().toLowerCase();
    const matched=val&&opts.some(o=>norm(o)===norm(val));
    const extraOpt=(val&&!matched)?`<option value="${esc(val)}" selected>${esc(val)}</option>`:'';
    inp=`<select name="${f.name}" id="${f.name}" ${f.required?'required':''}><option value="">—</option>${extraOpt}${opts.map(o=>`<option value="${esc(o)}"${norm(o)===norm(val)?' selected':''}>${esc(o)}</option>`).join('')}</select>`;
  } else if(f.type==='multiselect'){
    const opts=typeof f.opts==='function'?f.opts():f.opts;
    inp=buildMultiselect({id:f.name,value:val,options:opts,placeholder:f.placeholder||'Select…',required:f.required});
  } else if(f.type==='textarea'){
    inp=`<textarea name="${f.name}" id="${f.name}" ${f.required?'required':''}>${esc(val)}</textarea>`;
  } else {
    const extra=f.type==='number'?`min="${f.min??''}" max="${f.max??''}" step="any"`:'';
    inp=`<input type="${f.type||'text'}" name="${f.name}" id="${f.name}" value="${esc(val)}" ${extra} ${f.required?'required':''}>`;
  }
  return `<div class="form-field ${wrap}"${f.wrapId?` id="${f.wrapId}"`:''}><label>${esc(f.label)}${req}</label>${inp}</div>`;
}

function openAddModal(kind){
  const cfg=ADD_CONFIGS[kind];
  const root=document.getElementById('modal-root');
  root.innerHTML=`<div class="modal-backdrop" id="modal-backdrop">
    <div class="modal">
      <div class="modal-head"><h3>${esc(cfg.title)}</h3><button class="modal-close" onclick="closeModal()">×</button></div>
      <form id="modal-form" novalidate>
        <div class="modal-body">
          <div class="form-grid">${cfg.fields.map(f=>renderModalField(f)).join('')}</div>
          ${kind==='opportunities'?`<div id="prob-lock-msg" class="prob-locked" style="display:none">🔒 Probability auto-set by stage</div>`:''}
          ${kind==='clients'?`<div class="form-footnote"><strong>Active business</strong> is for clients with whom we have current active projects.<br><strong>Active engagement</strong> is for clients with whom we are actively engaging in conversations, but there is no current active project.</div>`:''}
        </div>
        <div class="modal-foot">
          <button type="button" class="btn" onclick="closeModal()">Cancel</button>
          <button type="submit" class="btn btn-primary" id="modal-save">Save</button>
        </div>
      </form>
    </div>
  </div>`;

  document.getElementById('modal-backdrop').addEventListener('click',e=>{if(e.target.id==='modal-backdrop')closeModal();});

  if(kind==='opportunities'){
    const stSel=document.getElementById('opportunity_status');
    if(stSel) stSel.addEventListener('change',()=>handleStageChange(stSel));
    const coSel=document.getElementById('modal-opp-company');
    if(coSel) coSel.addEventListener('change',()=>{resetClientLock('modal-opp-client-name');updateModalOppClientOptions(coSel.value);});
    // Default BD owner to whoever's logged in and creating this record — the person
    // logging an opportunity is almost always its owner. A client-specific match (if one
    // exists once they pick a client) overrides this, since that's more authoritative.
    const bdOwnerEl=document.getElementById('bd_owner');
    if(bdOwnerEl&&!bdOwnerEl.value&&currentUser){
      const meInList=EMPLOYEES.find(n=>n.toLowerCase()===String(currentUser.name||'').trim().toLowerCase());
      if(meInList) bdOwnerEl.value=meInList;
    }
  }
  if(kind==='engagements'){
    const coSel=document.getElementById('modal-eng-company');
    if(coSel) coSel.addEventListener('change',()=>{resetClientLock('modal-eng-client','designation');updateModalEngClientOptions(coSel.value);});
  }
  if(kind==='clients'){
    const rtSel=document.getElementById('region_type');
    if(rtSel){
      updateRegionDetailUI(rtSel.value);
      rtSel.addEventListener('change',()=>{
        updateRegionDetailUI(rtSel.value);
        const detInp=document.getElementById('region_detail'); if(detInp) detInp.value='';
      });
    }
  }

  document.getElementById('modal-form').addEventListener('submit',e=>handleAddSubmit(e,kind,cfg));
}

function handleStageChange(stSel){
  const prob=autoProb(stSel.value);
  const pctInp=document.getElementById('probability_pct');
  const lockMsg=document.getElementById('prob-lock-msg');
  if(prob!==null&&pctInp){
    const pctVal=String(Math.round(prob*100));
    // The dropdown only offers 10/25/75/100 as manual choices. Loss locks to 0%, which
    // isn't one of them — inject it as a one-off option so the lock still displays
    // correctly instead of silently landing on a blank selection.
    if(pctInp.tagName==='SELECT'&&![...pctInp.options].some(o=>o.value===pctVal)){
      const opt=document.createElement('option');
      opt.value=pctVal;opt.textContent=pctVal;
      pctInp.insertBefore(opt,pctInp.firstChild);
    }
    pctInp.value=pctVal;
    pctInp.disabled=true;
    if(lockMsg) lockMsg.style.display='block';
  } else if(pctInp){
    pctInp.disabled=false;
    if(lockMsg) lockMsg.style.display='none';
  }
}

async function handleAddSubmit(e,kind,cfg){
  e.preventDefault();
  const saveBtn=document.getElementById('modal-save');
  // Validate required fields
  let valid=true;
  cfg.fields.filter(f=>f.required).forEach(f=>{
    const id=f.comboId||f.name;
    const el=document.getElementById(id);
    if(el&&!el.value.trim()){el.classList.add('invalid');valid=false;}
    else if(el) el.classList.remove('invalid');
  });
  if(!valid){alert('Please fill in all required fields marked with *.');return;}

  // Enforce min/max on numeric fields (e.g. Target Revenue can't be negative).
  let inRange=true;
  cfg.fields.filter(f=>f.type==='number'&&(f.min!==undefined||f.max!==undefined)).forEach(f=>{
    const id=f.comboId||f.name;
    const el=document.getElementById(id);
    if(!el||!el.value.trim()){if(el) el.classList.remove('invalid');return;}
    const n=Number(el.value);
    if((f.min!==undefined&&n<f.min)||(f.max!==undefined&&n>f.max)){el.classList.add('invalid');inRange=false;}
    else el.classList.remove('invalid');
  });
  if(!inRange){alert('Please check the highlighted fields — some values are out of the allowed range.');return;}

  saveBtn.disabled=true;saveBtn.textContent='Saving…';

  // Build payload
  const payload={};
  cfg.fields.forEach(f=>{
    const id=f.comboId||f.name;
    const el=document.getElementById(id);
    let v=el?el.value.trim()||null:null;
    if(f.type==='number'&&v!==null) v=Number(v);
    // Map combo company/client fields to real keys
    const keyMap={'modal-company':'company','modal-client-name':'client_name','modal-opp-company':'company',
      'modal-opp-client-name':'client_name','modal-eng-company':'company','modal-eng-client':'client_name'};
    const realKey=keyMap[id]||f.name;
    payload[realKey]=v;
  });

  if(kind==='clients'){
    const type=payload.region_type,detail=payload.region_detail;
    if((type==='Regional'||type==='Country')&&!detail){
      alert('Please fill in the region detail field.');
      saveBtn.disabled=false;saveBtn.textContent='Save';
      return;
    }
    payload.region=type==='Global'?'Global':(type&&detail?`${type}: ${detail}`:null);
    delete payload.region_type;delete payload.region_detail;
  }
  if(kind==='companies'&&DATA.companies.some(c=>normCo(c.company)===normCo(payload.company))){
    alert(`"${payload.company}" already exists in the Companies list. Edit the existing record instead of adding a duplicate.`);
    saveBtn.disabled=false;saveBtn.textContent='Save';
    return;
  }
  if(kind==='clients'&&DATA.clients.some(c=>normCo(c.company)===normCo(payload.company)&&normStr(c.client_name).toLowerCase()===normStr(payload.client_name).toLowerCase())){
    alert(`"${payload.client_name}" already exists as a client at "${payload.company}". Edit the existing record instead of adding a duplicate.`);
    saveBtn.disabled=false;saveBtn.textContent='Save';
    return;
  }
  // Opportunities/engagements must attach to a client that's already on file for that
  // company — the "Client name" field is a free-text combo (autocomplete suggestions,
  // not a constrained dropdown), so typing past the suggestions has to be blocked here.
  if((kind==='opportunities'||kind==='engagements')&&!DATA.clients.some(c=>normCo(c.company)===normCo(payload.company)&&normStr(c.client_name).toLowerCase()===normStr(payload.client_name).toLowerCase())){
    alert(`"${payload.client_name}" is not an existing client at "${payload.company}". Add them on the Clients page first, then pick them here from the dropdown.`);
    saveBtn.disabled=false;saveBtn.textContent='Save';
    return;
  }
  if(kind==='opportunities'){
    const pct=payload.probability_pct;
    payload.probability_pct=pct!==null?pct/100:null;
    payload.probability_weighted_value=(pct!==null&&payload.estimated_value_usd!=null)?Number(payload.estimated_value_usd)*(pct/100):null;
    // Start stage history
    payload.stage_history=payload.opportunity_status?JSON.stringify([{stage:payload.opportunity_status,date:new Date().toISOString().slice(0,10),by:currentUser.username}]):null;
  }
  if(kind==='engagements'&&payload.eng_date){
    const d=new Date(payload.eng_date);
    payload.eng_month=new Date(d.getFullYear(),d.getMonth(),1).toISOString().slice(0,10);
  }

  const result=await insertRecord(kind,payload);
  saveBtn.disabled=false;saveBtn.textContent='Save';
  if(result){
    closeModal();await refreshAndRenderAll();
  }
}

/* ---------- 20. MODAL: EDIT ---------- */
async function openEditModal(kind,id){
  const record=(DATA[kind]||[]).find(r=>r.id===id);
  if(!record){alert('Record not found.');return;}

  const cfg=ADD_CONFIGS[kind];
  const root=document.getElementById('modal-root');

  // Build initial values map (handle combo key aliases)
  const vals={...record};
  const comboKeyMap={company:'modal-company',client_name:'modal-client-name'};
  if(kind==='opportunities') comboKeyMap.company='modal-opp-company', comboKeyMap.client_name='modal-opp-client-name';
  if(kind==='engagements') comboKeyMap.company='modal-eng-company', comboKeyMap.client_name='modal-eng-client';

  function getVal(f){
    if(f.name==='region_type'||f.name==='region_detail'){
      const raw=vals.region||'';
      const isRegional=raw.startsWith('Regional:'), isCountry=raw.startsWith('Country:');
      if(f.name==='region_type'){
        if(raw==='Global') return 'Global';
        if(isRegional) return 'Regional';
        if(isCountry) return 'Country';
        return raw?'Country':''; // legacy flat values (pre-cascade) default into the Country bucket
      }
      if(isRegional||isCountry) return raw.split(':').slice(1).join(':').trim();
      if(raw==='Global') return '';
      return raw; // legacy flat value shown as-is
    }
    const reverseMap={'modal-company':'company','modal-client-name':'client_name',
      'modal-opp-company':'company','modal-opp-client-name':'client_name',
      'modal-eng-company':'company','modal-eng-client':'client_name'};
    const realKey=reverseMap[f.comboId||f.name]||(f.comboId||f.name);
    let v=vals[realKey]??'';
    if(f.name==='probability_pct'&&v!==''&&v!==null) v=Math.round(Number(v)*100);
    return String(v);
  }

  const hist=kind==='opportunities'?getHistory(record):[];
  const histHtml=kind==='opportunities'?`<div class="stage-history" style="border-top:1px solid var(--line);padding-top:14px;margin-top:4px;"><h4>Stage history</h4><div class="history-timeline">${renderHistoryTimeline(hist)}</div></div>`:'';

  root.innerHTML=`<div class="modal-backdrop" id="modal-backdrop">
    <div class="modal">
      <div class="modal-head"><h3>Edit · ${esc(cfg.title.replace('Add ',''))}</h3><button class="modal-close" onclick="closeModal()">×</button></div>
      <form id="modal-form" novalidate>
        <div class="modal-body">
          <div class="form-grid">${cfg.fields.map(f=>renderModalField(f,getVal(f))).join('')}</div>
          ${kind==='opportunities'?`<div id="prob-lock-msg" class="prob-locked" style="display:none">🔒 Probability auto-set by stage</div>`:''}
          ${kind==='clients'?`<div class="form-footnote"><strong>Active business</strong> is for clients with whom we have current active projects.<br><strong>Active engagement</strong> is for clients with whom we are actively engaging in conversations, but there is no current active project.</div>`:''}
          ${histHtml}
        </div>
        <div class="modal-foot">
          <button type="button" class="btn" onclick="closeModal()">Cancel</button>
          <button type="submit" class="btn btn-primary" id="modal-save">Save changes</button>
        </div>
      </form>
    </div>
  </div>`;

  document.getElementById('modal-backdrop').addEventListener('click',e=>{if(e.target.id==='modal-backdrop')closeModal();});

  if(kind==='opportunities'){
    const stSel=document.getElementById('opportunity_status');
    if(stSel) stSel.addEventListener('change',()=>handleStageChange(stSel));
    // Check current stage
    const cur=document.getElementById('opportunity_status');
    if(cur&&autoProb(cur.value)!==null) handleStageChange(cur);
    // Cascade client options
    const coCur=document.getElementById('modal-opp-company');
    if(coCur){
      if(coCur.value) updateModalOppClientOptions(coCur.value);
      coCur.addEventListener('change',()=>{resetClientLock('modal-opp-client-name');updateModalOppClientOptions(coCur.value);});
    }
  }
  if(kind==='engagements'){
    const coCur=document.getElementById('modal-eng-company');
    if(coCur){
      if(coCur.value) updateModalEngClientOptions(coCur.value);
      coCur.addEventListener('change',()=>{resetClientLock('modal-eng-client','designation');updateModalEngClientOptions(coCur.value);});
    }
  }
  if(kind==='clients'){
    const rtSel=document.getElementById('region_type');
    if(rtSel){
      updateRegionDetailUI(rtSel.value);
      rtSel.addEventListener('change',()=>{
        updateRegionDetailUI(rtSel.value);
        const detInp=document.getElementById('region_detail'); if(detInp) detInp.value='';
      });
    }
  }

  document.getElementById('modal-form').addEventListener('submit',e=>handleEditSubmit(e,kind,id,record,cfg));
}

function updateModalEngClientOptions(company){
  const listEl=document.getElementById('modal-eng-client-list');
  if(!listEl) return;
  const clients=DATA.clients.filter(c=>normCo(c.company)===normCo(company)).map(c=>c.client_name).filter(Boolean).sort();
  listEl.querySelectorAll('.combo-opt:not(.add-new)').forEach(e=>e.remove());
  const addNew=listEl.querySelector('.add-new');
  clients.forEach(cl=>{
    const div=document.createElement('div'); div.className='combo-opt';
    div.textContent=cl; div.onmousedown=()=>comboPick(div,'modal-eng-client',cl);
    listEl.insertBefore(div,addNew||null);
  });
}

async function openCoEditModal(companyName){
  const rec=DATA.companies.find(c=>normCo(c.company)===normCo(companyName));
  if(!rec){alert('Company record not found in SharePoint companies list.');return;}
  openEditModal('companies',rec.id);
}

async function handleEditSubmit(e,kind,id,original,cfg){
  e.preventDefault();
  const saveBtn=document.getElementById('modal-save');
  let valid=true;
  cfg.fields.filter(f=>f.required).forEach(f=>{
    const elId=f.comboId||f.name;
    const el=document.getElementById(elId);
    if(el&&!el.value.trim()){el.classList.add('invalid');valid=false;}
    else if(el) el.classList.remove('invalid');
  });
  if(!valid){alert('Please fill in all required fields marked with *.');return;}

  // Enforce min/max on numeric fields (e.g. Target Revenue can't be negative).
  let inRange=true;
  cfg.fields.filter(f=>f.type==='number'&&(f.min!==undefined||f.max!==undefined)).forEach(f=>{
    const elId=f.comboId||f.name;
    const el=document.getElementById(elId);
    if(!el||!el.value.trim()){if(el) el.classList.remove('invalid');return;}
    const n=Number(el.value);
    if((f.min!==undefined&&n<f.min)||(f.max!==undefined&&n>f.max)){el.classList.add('invalid');inRange=false;}
    else el.classList.remove('invalid');
  });
  if(!inRange){alert('Please check the highlighted fields — some values are out of the allowed range.');return;}

  saveBtn.disabled=true;saveBtn.textContent='Saving…';

  const payload={};
  cfg.fields.forEach(f=>{
    const id2=f.comboId||f.name;
    const el=document.getElementById(id2);
    let v=el?el.value.trim()||null:null;
    if(f.type==='number'&&v!==null) v=Number(v);
    const reverseMap={'modal-company':'company','modal-client-name':'client_name',
      'modal-opp-company':'company','modal-opp-client-name':'client_name',
      'modal-eng-company':'company','modal-eng-client':'client_name'};
    const realKey=reverseMap[id2]||f.name;
    payload[realKey]=v;
  });

  if(kind==='clients'){
    const type=payload.region_type,detail=payload.region_detail;
    if((type==='Regional'||type==='Country')&&!detail){
      alert('Please fill in the region detail field.');
      saveBtn.disabled=false;saveBtn.textContent='Save changes';
      return;
    }
    payload.region=type==='Global'?'Global':(type&&detail?`${type}: ${detail}`:null);
    delete payload.region_type;delete payload.region_detail;
  }
  if(kind==='opportunities'){
    const pct=payload.probability_pct;
    payload.probability_pct=pct!==null?pct/100:null;
    payload.probability_weighted_value=(pct!==null&&payload.estimated_value_usd!=null)?Number(payload.estimated_value_usd)*(pct/100):null;
    // Update stage history if stage changed
    if(payload.opportunity_status&&payload.opportunity_status!==original.opportunity_status){
      payload.stage_history=addHistoryEntry(original.stage_history||'[]',payload.opportunity_status,currentUser.username);
    } else {
      payload.stage_history=original.stage_history;
    }
  }
  if(kind==='engagements'&&payload.eng_date){
    const d=new Date(payload.eng_date);
    payload.eng_month=new Date(d.getFullYear(),d.getMonth(),1).toISOString().slice(0,10);
  }

  const ok=await updateRecord(kind,id,payload);
  saveBtn.disabled=false;saveBtn.textContent='Save changes';
  if(ok){
    closeModal();await refreshAndRenderAll();
  }
}

function closeModal(){ document.getElementById('modal-root').innerHTML=''; }

/* ---------- 21. BOOT ---------- */
async function boot(){
  const r=activeRegion();
  const placeholders=[r.siteId,...Object.values(r.listIds)];
  if(placeholders.some(v=>v.includes('YOUR_'))){
    document.getElementById('config-banner').style.display='block'; return;
  }
  document.getElementById('config-banner').style.display='none';
  ensureFxRate().then(()=>{ syncCurrencySelects(); if(displayCurrency==='SGD') renderAllViews(); });
  try{ await refreshAndRenderAll(); }
  catch(err){
    console.error(err);
    alert('Could not load data from SharePoint: '+err.message+'\n\nCheck Site ID, List IDs and that this account has access.');
  }
}

/* ---------- 22. EVENT WIRING ---------- */
// Guarded event binding: a single missing/renamed element logs a console error and is
// skipped, instead of throwing and silently aborting every wireEvents() call after it
// (which is exactly the kind of bug that makes unrelated dropdowns look "broken").
function on(id,event,handler){
  const el=document.getElementById(id);
  if(!el){ console.error(`wireEvents: #${id} not found — its ${event} handler was not attached.`); return; }
  el.addEventListener(event,handler);
}

function wireEvents(){
  on('login-submit','click',doLogin);
  on('logout-btn','click',doLogout);
  on('refresh-btn','click',refreshAndRenderAll);
  syncCurrencySelects();
  on('bd-stake-view','change',e=>{bdStakeView=e.target.value;renderBDFunnel();});
  on('ar-bd-stake-view','change',e=>{arStakeView=e.target.value;renderAllRegionsBDFunnel();});
  on('ar-period-filter','change',e=>{arPeriod=e.target.value;renderAllRegionsKpis();renderAllRegionsRegionTable();});
  on('currency-select','change',e=>{setDisplayCurrency(e.target.value);});
  on('ar-currency-select','change',e=>{setDisplayCurrency(e.target.value);});
  on('ov-period-filter','change',e=>{ovPeriod=e.target.value;renderOverviewBDTables();});

  document.querySelectorAll('.nav-item').forEach(btn=>{
    btn.addEventListener('click',()=>switchView(btn.dataset.view));
  });

  on('add-client-btn','click',()=>openAddModal('clients'));
  on('add-opp-btn','click',()=>openAddModal('opportunities'));
  on('add-eng-btn','click',()=>openAddModal('engagements'));
  on('add-company-btn','click',()=>openAddModal('companies'));

  // Client filters
  on('cl-filter-status','change',e=>{clFilters.status=e.target.value;renderClients();});
  on('cl-filter-company','change',e=>{clFilters.company=e.target.value;renderClients();});
  on('cl-filter-department','change',e=>{clFilters.department=e.target.value;renderClients();});
  on('cl-filter-priority','change',e=>{clFilters.priority=e.target.value;renderClients();});
  on('cl-clear','click',()=>{clFilters={status:'',company:'',department:'',priority:''};renderClients();});
  on('cl-pivot-therapy','change',e=>{clPivotFilters.therapy=e.target.value;renderClientPivot();});

  // Opportunity filters
  on('op-search','input',debounce(e=>{opFilters.search=e.target.value;renderOpportunities();},200));
  on('op-filter-status','change',e=>{opFilters.status=e.target.value;renderOpportunities();});
  on('op-filter-company','change',e=>{opFilters.company=e.target.value;renderOpportunities();});
  on('op-filter-bd','change',e=>{opFilters.bd=e.target.value;renderOpportunities();});
  on('op-clear','click',()=>{opFilters={search:'',status:'',company:'',bd:''};document.getElementById('op-search').value='';renderOpportunities();});
  on('op-stage-company','change',e=>{opStageChartCompany=e.target.value;renderOpStageChart();});

  // Engagement filters
  on('eg-filter-type','change',e=>{egFilters.type=e.target.value;renderEngagements();});
  on('eg-filter-company','change',e=>{egFilters.company=e.target.value;renderEngagements();});
  on('eg-filter-bd','change',e=>{egFilters.bd=e.target.value;renderEngagements();});
  on('eg-filter-accompanied','change',e=>{egFilters.accompanied=e.target.value;renderEngagements();});
  on('eg-filter-followup','change',e=>{egFilters.followup=e.target.value;renderEngagements();});
  on('eg-clear','click',()=>{egFilters={type:'',company:'',bd:'',accompanied:'',followup:'',week:''};renderEngagements();});
  on('egp-bd','change',e=>{egpFilters.bd=e.target.value;renderEngagementPivotTable();});
  on('egp-month','change',e=>{egpFilters.month=e.target.value;renderEngagementPivotTable();});
  on('egp-company','change',e=>{egpFilters.company=e.target.value;renderEngagementPivotTable();});
  on('egp-objective','change',e=>{egpFilters.objective=e.target.value;renderEngagementPivotTable();});
  on('egp-type','change',e=>{egpFilters.type=e.target.value;renderEngagementPivotTable();});
  on('egp-clear','click',()=>{egpFilters={bd:'',month:'',company:'',objective:'',type:''};renderEngagementPivotTable();});

  // Company filters
  on('co-search','input',debounce(e=>{coFilters.search=e.target.value;renderCompanies();},200));
  on('co-filter-onboarding','change',e=>{coFilters.onboarding=e.target.value;renderCompanies();});
  on('co-clear','click',()=>{coFilters={search:'',onboarding:''};document.getElementById('co-search').value='';renderCompanies();});

  // Account health matrix filters — independent of the company table's own filters above.
  on('co-health-search','input',debounce(e=>{coHealthFilters.search=e.target.value;renderCompanyHealthMatrix();},200));
  on('co-health-filter-onboarding','change',e=>{coHealthFilters.onboarding=e.target.value;renderCompanyHealthMatrix();});
  on('co-health-clear','click',()=>{coHealthFilters={search:'',onboarding:''};document.getElementById('co-health-search').value='';renderCompanyHealthMatrix();});
}

document.addEventListener('DOMContentLoaded',wireEvents);

