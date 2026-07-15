'use strict';

// ── Globals ────────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const apiWv   = $('api-wv');
const spWv    = $('sp-wv');
const loginWv = $('login-wv');

// SharePoint PTO calendar (source of truth for who's out)
const SP_SITE = 'https://YOUR-TENANT.sharepoint.com/sites/YOUR-SITE';
const SP_TIMEOFF_LIST = 'YOUR-TIME-OFF-LIST';
async function spFetch(apiPath) {
  const url = SP_SITE + apiPath;
  const script = `fetch(${JSON.stringify(url)},{headers:{Accept:'application/json;odata=nometadata'},credentials:'include'}).then(r=>r.text()).catch(e=>JSON.stringify({__err:String(e&&e.message||e)}))`;
  const txt = await spWv.executeJavaScript(script);
  try { return JSON.parse(txt); } catch (_) { return { __raw: String(txt).slice(0, 300) }; }
}
async function fetchSharePointTimeOff(fromIso, toIso) {
  try {
    const from = fromIso.replace(/\.\d+Z$/, 'Z');
    const to   = toIso.replace(/\.\d+Z$/, 'Z');
    const path = `/_api/web/lists/getbytitle('${SP_TIMEOFF_LIST}')/items` +
      `?$select=Title,EventDate,EndDate,Category` +
      `&$filter=EventDate le '${to}' and EndDate ge '${from}'&$top=1000`;
    const r = await spFetch(path);
    if (!r || !r.value) { if (r && r['odata.error']) console.warn('SP PTO filter error:', r['odata.error']?.message?.value); return []; }
    // All-day events store dates at UTC midnight but mean a local date — use the date part only.
    const dpart = s => { const [y, m, d] = String(s).slice(0, 10).split('-').map(Number); return new Date(y, (m || 1) - 1, d || 1); };
    return r.value.map(i => {
      const start = dpart(i.EventDate);
      const end   = dpart(i.EndDate); end.setHours(23, 59, 59, 999);
      return { name: i.Title, start, end, category: i.Category };
    });
  } catch (e) { console.warn('SharePoint PTO fetch failed:', e); return []; }
}
let orgUrl = '';
let xrmReady = false;
let activeTab = 'bookings';
let weekOffset = 0;

const cache = { bookings: null, schedule: null, scheduleWeek: null, accounts: null, contacts: null };

function toast(msg, isError) {
  let el = document.getElementById('app-toast');
  if (!el) {
    el = document.createElement('div');
    el.id = 'app-toast';
    el.style.cssText = 'position:fixed;bottom:22px;left:50%;transform:translateX(-50%);z-index:500;padding:12px 20px;border-radius:10px;font-size:13px;font-weight:600;box-shadow:0 12px 40px rgba(0,0,0,.5);max-width:560px;text-align:center;transition:opacity .3s;';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.background = isError ? 'var(--danger)' : 'var(--bg2)';
  el.style.color = isError ? '#fff' : 'var(--text)';
  el.style.border = '1px solid ' + (isError ? 'var(--danger)' : 'var(--border)');
  el.style.opacity = '1';
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.style.opacity = '0'; }, isError ? 6500 : 4000);
}

// ── Utility ────────────────────────────────────────────────────────────────
const esc  = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const sleep = ms => new Promise(r => setTimeout(r, ms));

function fmtDate(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleDateString(undefined, { month:'short', day:'numeric', year:'numeric' });
}
function fmtTime(iso) {
  if (!iso) return '';
  return new Date(iso).toLocaleTimeString(undefined, { hour:'2-digit', minute:'2-digit' });
}
function fmtDateTime(iso) {
  if (!iso) return '—';
  const d = new Date(iso);
  return d.toLocaleDateString(undefined, { month:'short', day:'numeric', year:'numeric' }) + ' ' +
         d.toLocaleTimeString(undefined, { hour:'2-digit', minute:'2-digit' });
}

const RESOURCE_COLORS = [
  '#7c6af7','#5b8af7','#12b76a','#f79009','#f04438',
  '#06b6d4','#8b5cf6','#ec4899','#14b8a6','#f97316'
];
const resourceColorMap = {};
let colorIdx = 0;
function resourceColor(name) {
  if (!resourceColorMap[name]) resourceColorMap[name] = RESOURCE_COLORS[colorIdx++ % RESOURCE_COLORS.length];
  return resourceColorMap[name];
}

function statusBadge(statusName) {
  const s = (statusName || '').toLowerCase();
  let cls = 'badge-default';
  if (s === 'free') cls = 'badge-free';
  else if (s.includes('scheduled') || s.includes('committed')) cls = 'badge-scheduled';
  else if (s.includes('travel'))     cls = 'badge-traveling';
  else if (s.includes('progress') || s.includes('started')) cls = 'badge-inprogress';
  else if (s.includes('complete'))   cls = 'badge-completed';
  else if (s.includes('cancel'))     cls = 'badge-cancelled';
  return `<span class="badge ${cls}">${esc(statusName || 'Unknown')}</span>`;
}

function substatusBadge(sub) {
  if (!sub) return '<span class="muted">—</span>';
  const s = sub.toLowerCase();
  // Reuse the Status badge styles: green = complete, yellow/orange = follow up,
  // blue = unscheduled, grey = sent to bc / everything else.
  let cls = 'badge-cancelled'; // grey default
  if (s.includes('complete'))       cls = 'badge-completed';  // green
  else if (s.includes('follow'))    cls = 'badge-traveling';  // yellow/orange
  else if (s.includes('unschedul')) cls = 'badge-scheduled';  // blue
  return `<span class="badge ${cls}">${esc(sub)}</span>`;
}

function bookingBlockClass(statusName) {
  const s = (statusName || '').toLowerCase();
  if (s.includes('travel'))      return 'traveling';
  if (s.includes('progress') || s.includes('started')) return 'inprogress';
  if (s.includes('complete'))    return 'completed';
  if (s.includes('cancel'))      return 'cancelled';
  return '';
}

// ── Open a Dynamics record in a new window ────────────────────────────────
function openRecord(r) {
  const bookingId = r.bookableresourcebookingid;
  if (!bookingId) return;
  window.api.openWorkOrder(bookingId, orgUrl, r.name || 'Booking');
}

// ── Setup ──────────────────────────────────────────────────────────────────
let appSettings = {};
async function persistSettings() { await window.api.saveSettings(appSettings); }

function applyTeam(team) {
  SCHEDULE_RESOURCES = Array.isArray(team) ? team.filter(Boolean) : [];
  DEFAULT_BOOKINGS_RESOURCE = SCHEDULE_RESOURCES[0] || ALL_MEMBERS_VALUE;
  if (!bookingsResource) bookingsResource = DEFAULT_BOOKINGS_RESOURCE;
  populateBookingsResourceFilter();
}

async function init() {
  appSettings = (await window.api.getSettings()) || {};
  const teamNeverConfigured = appSettings.team === undefined;
  applyTeam(appSettings.team);
  if (appSettings.orgUrl) {
    orgUrl = appSettings.orgUrl.replace(/\/$/, '');
    startApp();
    // Existing install that predates configurable teams — prompt once to set them up
    if (teamNeverConfigured) openTeamManager(true);
  } else {
    $('setup-overlay').classList.remove('hidden');
  }
}

$('setup-btn').addEventListener('click', async () => {
  const val = $('setup-url').value.trim().replace(/\/$/, '');
  $('setup-error').textContent = '';
  if (!val.startsWith('http')) { $('setup-error').textContent = 'Enter a valid URL starting with https://'; return; }
  appSettings.orgUrl = val;
  await persistSettings();
  orgUrl = val;
  startApp();
  // First-run: if no team configured yet, prompt the user to add their team members
  if (!SCHEDULE_RESOURCES.length) openTeamManager(true);
});
$('setup-url').addEventListener('keydown', e => { if (e.key === 'Enter') $('setup-btn').click(); });

$('btn-change-org').addEventListener('click', () => {
  $('setup-url').value = orgUrl;
  $('setup-error').textContent = '';
  $('shell').classList.add('hidden');
  $('setup-overlay').classList.remove('hidden');
  xrmReady = false;
  cache.bookings = cache.accounts = cache.contacts = null;
});

// ── App start ──────────────────────────────────────────────────────────────
function startApp() {
  $('setup-overlay').classList.add('hidden');
  $('shell').classList.remove('hidden');
  try {
    const host = new URL(orgUrl).hostname.split('.')[0];
    $('user-org').textContent  = host;
    $('user-avatar').textContent = host.slice(0, 2).toUpperCase();
  } catch (_) {}
  apiWv.src = orgUrl + '/main.aspx';
  spWv.src  = SP_SITE + '/SitePages/Home.aspx';
  setupApiWebview();
}

// ── API webview ────────────────────────────────────────────────────────────
function setupApiWebview() {
  apiWv.addEventListener('did-navigate', e => {
    const u = e.url || '';
    if (u.includes('login.microsoftonline.com') || u.includes('login.microsoft.com') || u.includes('login.live.com')) {
      showLoginOverlay(u);
    }
  });

  apiWv.addEventListener('did-finish-load', async () => {
    if ((apiWv.getURL() || '').startsWith(orgUrl)) {
      hideLoginOverlay();
      await waitForXrm();
      if (xrmReady) { setStatus('Connected'); loadTab(activeTab); }
    }
  });
}

function showLoginOverlay(initialUrl) {
  $('login-overlay').classList.remove('hidden');
  if (loginWv && initialUrl) loginWv.src = initialUrl;
  loginWv.addEventListener('did-navigate', function onNav(e) {
    if ((e.url || '').startsWith(orgUrl)) {
      loginWv.removeEventListener('did-navigate', onNav);
      apiWv.src = orgUrl + '/main.aspx';
    }
  });
}
function hideLoginOverlay() { $('login-overlay').classList.add('hidden'); }

async function waitForXrm() {
  setStatus('Loading…');
  // Poll frequently so we proceed the instant Xrm is available. ~30s timeout.
  for (let i = 0; i < 200; i++) {
    try {
      const ready = await apiWv.executeJavaScript('typeof Xrm !== "undefined" && typeof Xrm.WebApi !== "undefined"');
      if (ready) {
        xrmReady = true;
        // Share this warm session with child windows so they don't reload the shell.
        try { await window.api.registerApiWebview(apiWv.getWebContentsId()); } catch(_) {}
        try {
          const uname = await apiWv.executeJavaScript('Xrm.Utility.getGlobalContext().getUserName()');
          if (uname) $('user-name').textContent = uname;
        } catch(_) {}
        return;
      }
    } catch (_) {}
    await sleep(150);
  }
  setStatus('Could not connect');
}

function setStatus(msg) {
  $('titlebar-status').textContent = msg;
  const inProgress = msg === 'Connected' || msg === 'Loading…' || msg === 'Reconnecting…';
  $('reconnect-btn').classList.toggle('hidden', inProgress);
}

function reconnect() {
  xrmReady = false;
  setStatus('Reconnecting…');
  apiWv.src = orgUrl + '/main.aspx';
}
$('reconnect-btn').addEventListener('click', reconnect);

// ── Xrm.WebApi fetch ──────────────────────────────────────────────────────
async function xrmFetchPage(entity, query) {
  if (!xrmReady) throw new Error('Xrm not ready');
  const safeQuery = query.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/'/g, "\\'");
  const script = `
    new Promise((resolve, reject) => {
      try {
        Xrm.WebApi.retrieveMultipleRecords('${entity}', '${safeQuery}')
          .then(r => resolve(JSON.stringify({ entities: r.entities, nextLink: r.nextLink })))
          .catch(e => reject(String(e.message || e)));
      } catch(ex) { reject(String(ex.message || ex)); }
    })
  `;
  const json = await apiWv.executeJavaScript(script);
  return JSON.parse(json);
}

window.__dumpEntityFields = async function(entityLogicalName, prefix) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes?$select=LogicalName,DisplayName,AttributeType`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  if (r?.__err) { console.error('fetch error:', r.__err); return; }
  if (!Array.isArray(r?.value)) { console.error('Unexpected response:', json.slice(0, 500)); return; }
  let out = r.value.map(a => ({
    LogicalName: a.LogicalName,
    Label: a.DisplayName?.UserLocalizedLabel?.Label || '',
    Type: a.AttributeType
  })).sort((a,b) => a.LogicalName.localeCompare(b.LogicalName));
  if (prefix) out = out.filter(a => (a.LogicalName.startsWith(prefix) || a.LogicalName.includes(prefix) || a.Label.toLowerCase().includes(prefix.toLowerCase())) && a.Label);
  console.log(JSON.stringify(out));
  return out;
};

// List entities whose name contains a substring — helps find the right logical name
window.__findEntities = async function(substr) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  if (!Array.isArray(r?.value)) { console.error('Unexpected:', json.slice(0,500)); return; }
  const s = (substr||'').toLowerCase();
  const out = r.value.map(e => ({
    LogicalName: e.LogicalName,
    Label: e.DisplayName?.UserLocalizedLabel?.Label || ''
  })).filter(e => !s || e.LogicalName.includes(s) || e.Label.toLowerCase().includes(s))
     .sort((a,b)=>a.LogicalName.localeCompare(b.LogicalName));
  console.log(JSON.stringify(out));
  return out;
};

window.__dumpWorkOrder = async function(number) {
  const skip = /^(createdon|modifiedon|_createdby|_modifiedby|_owningbusinessunit|_ownerid|_owninguser|_owningteam|versionnumber|overriddencreatedon|importsequencenumber|timezone|utcconversion|exchangerate|traversedpath|processid|stageid)/i;
  const clean = obj => {
    const out = {};
    for (const k of Object.keys(obj)) {
      if (k.includes('@') && !k.includes('FormattedValue')) continue;
      if (skip.test(k)) continue;
      if (obj[k] === null || obj[k] === '') continue;
      out[k] = obj[k];
    }
    return out;
  };
  const wos = await xrmFetch('msdyn_workorder', `?$filter=msdyn_name eq '${number.replace(/'/g,"''")}'&$top=1`);
  if (!wos[0]) { console.log('Work order not found:', number); return; }
  console.log('WORK ORDER:', JSON.stringify(clean(wos[0]), null, 1));
  const bookings = await xrmFetch('bookableresourcebooking', `?$filter=_msdyn_workorder_value eq ${wos[0].msdyn_workorderid}`);
  console.log('BOOKINGS (' + bookings.length + '):', bookings[0] ? JSON.stringify(clean(bookings[0]), null, 1) : 'none');
};

async function xrmCreate(entity, data) {
  if (!xrmReady) throw new Error('Xrm not ready');
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{const __d=${JSON.stringify(data)};const r=await Xrm.WebApi.createRecord("${entity}",__d);return JSON.stringify({id:r.id});}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
  return r.id;
}
async function xrmRetrieve(entity, id, query = '') {
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{return JSON.stringify(await Xrm.WebApi.retrieveRecord("${entity}","${id}","${query.replace(/"/g,'\\"')}"));}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
  return r;
}
const _navCache = {};
async function getNavPropMap(entity) {
  if (_navCache[entity]) return _navCache[entity];
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entity}')/ManyToOneRelationships?$select=ReferencingAttribute,ReferencingEntityNavigationPropertyName,ReferencedEntity`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)},{headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  const map = {};
  (r.value || []).forEach(rel => { map[rel.ReferencingAttribute] = { nav: rel.ReferencingEntityNavigationPropertyName, target: rel.ReferencedEntity }; });
  _navCache[entity] = map;
  return map;
}
const _setCache = {};
async function entitySetOf(logical) {
  if (_setCache[logical]) return _setCache[logical];
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${logical}')?$select=EntitySetName`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)},{headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  _setCache[logical] = r.EntitySetName || (logical + 's');
  return _setCache[logical];
}

async function xrmFetch(entity, query = '') {
  if (!xrmReady) throw new Error('Xrm not ready');
  let all = [];
  let page = await xrmFetchPage(entity, query);
  all = all.concat(page.entities);
  let nextLink = page.nextLink;
  while (nextLink) {
    const qs = nextLink.split('?')[1];
    if (!qs) break;
    page = await xrmFetchPage(entity, '?' + qs);
    all = all.concat(page.entities);
    nextLink = page.nextLink;
  }
  return all;
}

// ── Tab navigation ─────────────────────────────────────────────────────────
document.querySelectorAll('.nav-item').forEach(item => {
  item.addEventListener('click', () => {
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    item.classList.add('active');
    activeTab = item.dataset.tab;
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.add('hidden'));
    $(`tab-${activeTab}`).classList.remove('hidden');
    if (xrmReady) loadTab(activeTab);
  });
});

function loadTab(tab) {
  if (tab === 'bookings') loadBookings();
  else if (tab === 'schedule') renderSchedule().then(scrollToToday);
  else if (tab === 'accounts') loadAccounts();
  else if (tab === 'contacts') loadContacts();
  else if (tab === 'team') loadTeam();
  else if (tab === 'assets') showState('assets', 'empty');
}

// ── BOOKINGS ───────────────────────────────────────────────────────────────
let bookingsFilter = 'upcoming';
let bookingsSearch = '';
let bookingsStatusFilter = '';
let bookingsSubstatusFilter = '';
const BOOKINGS_SUBSTATUS_OPTIONS = ['5 Day Monitoring', 'Completed', 'Unscheduled', 'Follow-up Required', 'Parts Required'];
(function initBookingsSubstatusFilter() {
  const sel = $('bookings-substatus-filter');
  sel.innerHTML = '<option value="">All Substatuses</option>' +
    BOOKINGS_SUBSTATUS_OPTIONS.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
  sel.addEventListener('change', e => {
    bookingsSubstatusFilter = e.target.value;
    renderBookings(currentBookingsDataset(), cache.bookingsCustomerMap || {});
  });
})();
let bookingsMonthOffset = 0; // 0 = current month
let DEFAULT_BOOKINGS_RESOURCE = ''; // first team member; set once team is loaded
let bookingsResource = '';
const ALL_MEMBERS_VALUE = '__ALL__';

function getMonthRange(offset) {
  const now   = new Date();
  const year  = now.getFullYear();
  const month = now.getMonth() + offset;
  const from  = new Date(year, month, 1);
  const to    = new Date(year, month + 1, 0, 23, 59, 59);
  return { from, to };
}

function updateBookingsMonthLabel() {
  const { from } = getMonthRange(bookingsMonthOffset);
  $('bookings-month-label').textContent = from.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });
}

let allBookingStatusNames = null;
async function loadAllBookingStatusNames() {
  if (allBookingStatusNames) return allBookingStatusNames;
  try {
    const statuses = await xrmFetch('bookingstatus', '?$select=name,statuscode&$orderby=name asc');
    const seen = new Map();
    for (const s of statuses) {
      if (!seen.has(s.name) || s.statuscode < seen.get(s.name).statuscode) seen.set(s.name, s);
    }
    allBookingStatusNames = [...seen.keys()].sort((a, b) => a.localeCompare(b));
  } catch (_) {
    allBookingStatusNames = [];
  }
  populateStatusFilterOptions();
  return allBookingStatusNames;
}

function resourceFilterClause() {
  if (bookingsResource === ALL_MEMBERS_VALUE || !bookingsResource) {
    if (!SCHEDULE_RESOURCES.length) return `Resource/name eq '__no_team_configured__'`;
    return '(' + SCHEDULE_RESOURCES.map(n => `Resource/name eq '${n.replace(/'/g, "''")}'`).join(' or ') + ')';
  }
  return `Resource/name eq '${bookingsResource.replace(/'/g, "''")}'`;
}

function populateStatusFilterOptions() {
  const statusSel = $('bookings-status-filter');
  if (!allBookingStatusNames) return;
  const prevValue = statusSel.value;
  statusSel.innerHTML = '<option value="">All Statuses</option>' +
    allBookingStatusNames.map(s => `<option value="${esc(s)}">${esc(s)}</option>`).join('');
  statusSel.value = allBookingStatusNames.includes(prevValue) ? prevValue : '';
}

async function loadBookings(force = false) {
  if (!allBookingStatusNames) loadAllBookingStatusNames();
  const monthKey = `month:${bookingsMonthOffset}:${bookingsResource}`;
  updateBookingsMonthLabel();
  if (!force && cache.bookingsMonths?.[monthKey]) {
    renderBookings(cache.bookingsMonths[monthKey], cache.bookingsCustomerMap || {});
    return;
  }

  showState('bookings', 'loading');

  const { from, to } = getMonthRange(bookingsMonthOffset);
  const fromIso = from.toISOString();
  const toIso   = to.toISOString();

  try {
    const records = await xrmFetch('bookableresourcebooking',
      '?$select=name,starttime,endtime,_resource_value,_bookingstatus_value,_msdyn_workorder_value,bookableresourcebookingid' +
      '&$expand=BookingStatus($select=name),Resource($select=name)' +
      `&$filter=starttime ge ${fromIso} and starttime le ${toIso} and ${resourceFilterClause()}` +
      '&$orderby=starttime desc&$top=500'
    );

    if (!cache.bookingsMonths) cache.bookingsMonths = {};
    cache.bookingsMonths[monthKey] = records;

    // Fetch customer names (cached, rarely change) + substatuses (always refreshed, change often)
    if (!cache.bookingsCustomerMap) cache.bookingsCustomerMap = {};
    if (!cache.bookingsSubstatusMap) cache.bookingsSubstatusMap = {};
    const knownIds = new Set(Object.keys(cache.bookingsCustomerMap));
    const allWoIds = [...new Set(records.map(r => r._msdyn_workorder_value).filter(Boolean))];
    const woIdsToFetch = force ? allWoIds : allWoIds.filter(id => !knownIds.has(id));
    if (woIdsToFetch.length) {
      try {
        const CHUNK = 30;
        const chunks = [];
        for (let i = 0; i < woIdsToFetch.length; i += CHUNK) chunks.push(woIdsToFetch.slice(i, i + CHUNK));
        const results = await Promise.all(chunks.map(chunk => {
          const filter = chunk.map(id => `msdyn_workorderid eq ${id}`).join(' or ');
          return xrmFetch('msdyn_workorder',
            `?$select=msdyn_workorderid,_msdyn_serviceaccount_value,_msdyn_substatus_value&$filter=${filter}&$top=${chunk.length}`
          );
        }));
        results.flat().forEach(wo => {
          const name = wo['_msdyn_serviceaccount_value@OData.Community.Display.V1.FormattedValue'] || '';
          if (name) cache.bookingsCustomerMap[wo.msdyn_workorderid] = name;
          const substatus = wo['_msdyn_substatus_value@OData.Community.Display.V1.FormattedValue'] || '';
          cache.bookingsSubstatusMap[wo.msdyn_workorderid] = substatus;
        });
      } catch (_) {}
    }

    renderBookings(records, cache.bookingsCustomerMap);
  } catch (e) {
    showState('bookings', 'empty');
    console.error('Bookings error:', e);
  }
}

function renderBookings(records, customerMap = {}) {
  const now        = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const todayEnd   = new Date(todayStart.getTime() + 86400000);

  let rows = records.filter(r => {
    const resource = (r['_resource_value@OData.Community.Display.V1.FormattedValue'] || r.Resource?.name || '').toLowerCase();
    if (bookingsResource === ALL_MEMBERS_VALUE) {
      if (!SCHEDULE_RESOURCES.some(n => resource.includes(n.toLowerCase()))) return false;
    } else if (!resource.includes(bookingsResource.toLowerCase())) {
      return false;
    }
    if (bookingsSearch) return true; // search bypasses date filters
    const start = r.starttime ? new Date(r.starttime) : null;
    if (bookingsFilter === 'today')    return start && start >= todayStart && start < todayEnd;
    if (bookingsFilter === 'upcoming') return start && start >= todayStart;
    return true;
  });

  if (bookingsSearch) {
    const q = bookingsSearch.toLowerCase();
    rows = rows.filter(r =>
      (r.name || '').toLowerCase().includes(q) ||
      (customerMap[r._msdyn_workorder_value] || '').toLowerCase().includes(q) ||
      (r['_resource_value@OData.Community.Display.V1.FormattedValue'] || '').toLowerCase().includes(q) ||
      (r.BookingStatus?.name || '').toLowerCase().includes(q)
    );
  }

  populateStatusFilterOptions();
  bookingsStatusFilter = $('bookings-status-filter').value;

  if (bookingsStatusFilter) {
    rows = rows.filter(r =>
      (r.BookingStatus?.name || r['_bookingstatus_value@OData.Community.Display.V1.FormattedValue'] || '') === bookingsStatusFilter
    );
  }

  if (bookingsSubstatusFilter) {
    const substatusMap = cache.bookingsSubstatusMap || {};
    rows = rows.filter(r => substatusMap[r._msdyn_workorder_value] === bookingsSubstatusFilter);
  }

  // Default sort: by date (rows already ordered by starttime from the underlying query)

  $('bookings-count').textContent = rows.length;
  if (rows.length === 0) { showState('bookings', 'empty'); return; }

  $('bookings-body').innerHTML = rows.map((r, i) => {
    const resource = r['_resource_value@OData.Community.Display.V1.FormattedValue'] || '—';
    const status   = r.BookingStatus?.name || r['_bookingstatus_value@OData.Community.Display.V1.FormattedValue'] || '';
    const customer = customerMap[r._msdyn_workorder_value] || '—';
    const substatus = (cache.bookingsSubstatusMap || {})[r._msdyn_workorder_value] || '';
    return `<tr class="clickable-row" data-idx="${i}">
      <td>${fmtDateTime(r.starttime)}</td>
      <td class="muted">${fmtDateTime(r.endtime)}</td>
      <td>${esc(customer)}</td>
      <td>${esc(r.name || '—')}</td>
      <td class="muted">${esc(resource)}</td>
      <td>${statusBadge(status)}</td>
      <td>${substatusBadge(substatus)}</td>
    </tr>`;
  }).join('');

  $('bookings-body').querySelectorAll('.clickable-row').forEach(tr => {
    tr.addEventListener('click', () => openRecord(rows[+tr.dataset.idx]));
  });

  showState('bookings', 'table');
}

function currentBookingsDataset() {
  if (bookingsSearch || bookingsFilter === 'all') return cache.bookingsAllByResource?.[bookingsResource] || [];
  const monthKey = `month:${bookingsMonthOffset}:${bookingsResource}`;
  return cache.bookingsMonths?.[monthKey] || [];
}

$('bookings-status-filter').addEventListener('change', e => {
  bookingsStatusFilter = e.target.value;
  renderBookings(currentBookingsDataset(), cache.bookingsCustomerMap || {});
});

document.querySelectorAll('#bookings-filter .chip').forEach(c => {
  c.addEventListener('click', () => {
    document.querySelectorAll('#bookings-filter .chip').forEach(x => x.classList.remove('active'));
    c.classList.add('active');
    bookingsFilter = c.dataset.filter;
    // An active search always pulls the full unscoped history, regardless of which date chip is selected
    if (bookingsSearch || bookingsFilter === 'all') {
      loadAllBookingsForSearch();
    } else {
      const monthKey = `month:${bookingsMonthOffset}:${bookingsResource}`;
      const cached = cache.bookingsMonths?.[monthKey];
      if (cached) renderBookings(cached, cache.bookingsCustomerMap || {});
      else loadBookings();
    }
  });
});
$('bookings-search').addEventListener('input', e => {
  bookingsSearch = e.target.value.trim();
  if (bookingsSearch) {
    loadAllBookingsForSearch();
  } else {
    const monthKey = `month:${bookingsMonthOffset}:${bookingsResource}`;
    const cached = cache.bookingsMonths?.[monthKey];
    if (cached) renderBookings(cached, cache.bookingsCustomerMap || {});
    else loadBookings();
  }
});

async function loadAllBookingsForSearch() {
  if (!allBookingStatusNames) loadAllBookingStatusNames();
  if (!cache.bookingsAllByResource) cache.bookingsAllByResource = {};
  if (cache.bookingsAllByResource[bookingsResource]) {
    renderBookings(cache.bookingsAllByResource[bookingsResource], cache.bookingsCustomerMap || {});
    return;
  }
  showState('bookings', 'loading');
  try {
    const records = await xrmFetch('bookableresourcebooking',
      '?$select=name,starttime,endtime,_resource_value,_bookingstatus_value,_msdyn_workorder_value,bookableresourcebookingid' +
      `&$expand=BookingStatus($select=name),Resource($select=name)` +
      `&$filter=${resourceFilterClause()}` +
      '&$orderby=starttime desc'
    );
    cache.bookingsAllByResource[bookingsResource] = records;

    // Fetch customer names + substatuses for every work order currently in view
    // (substatus always refreshed since it changes often; customer name rarely does but refreshing it too is cheap here)
    if (!cache.bookingsCustomerMap) cache.bookingsCustomerMap = {};
    if (!cache.bookingsSubstatusMap) cache.bookingsSubstatusMap = {};
    const woIdsToFetch = [...new Set(records.map(r => r._msdyn_workorder_value).filter(Boolean))];
    if (woIdsToFetch.length) {
      try {
        const CHUNK = 30;
        const chunks = [];
        for (let i = 0; i < woIdsToFetch.length; i += CHUNK) chunks.push(woIdsToFetch.slice(i, i + CHUNK));
        const results = await Promise.all(chunks.map(chunk => {
          const filter = chunk.map(id => `msdyn_workorderid eq ${id}`).join(' or ');
          return xrmFetch('msdyn_workorder',
            `?$select=msdyn_workorderid,_msdyn_serviceaccount_value,_msdyn_substatus_value&$filter=${filter}&$top=${chunk.length}`
          );
        }));
        results.flat().forEach(wo => {
          const name = wo['_msdyn_serviceaccount_value@OData.Community.Display.V1.FormattedValue'] || '';
          if (name) cache.bookingsCustomerMap[wo.msdyn_workorderid] = name;
          const substatus = wo['_msdyn_substatus_value@OData.Community.Display.V1.FormattedValue'] || '';
          cache.bookingsSubstatusMap[wo.msdyn_workorderid] = substatus;
        });
      } catch (_) {}
    }
    renderBookings(cache.bookingsAllByResource[bookingsResource], cache.bookingsCustomerMap);
  } catch (e) {
    showState('bookings', 'empty');
    console.error('Bookings search error:', e);
  }
}
$('bookings-month-prev').addEventListener('click',  () => { bookingsMonthOffset--; loadBookings(); });
$('bookings-month-next').addEventListener('click',  () => { bookingsMonthOffset++; loadBookings(); });
$('bookings-month-today').addEventListener('click', () => { bookingsMonthOffset = 0; loadBookings(); });
$('bookings-refresh').addEventListener('click', async () => {
  const wrap = $('bookings-table-wrap');
  const scrollPos = wrap.scrollTop;
  if (bookingsFilter === 'all') {
    if (cache.bookingsAllByResource) cache.bookingsAllByResource[bookingsResource] = null;
    await loadAllBookingsForSearch();
  } else {
    await loadBookings(true);
  }
  wrap.scrollTop = scrollPos;
});

$('bookings-reset-filters').addEventListener('click', () => {
  bookingsSearch = '';
  $('bookings-search').value = '';

  bookingsFilter = 'upcoming';
  document.querySelectorAll('#bookings-filter .chip').forEach(c => c.classList.toggle('active', c.dataset.filter === 'upcoming'));

  bookingsStatusFilter = '';
  $('bookings-status-filter').value = '';

  bookingsSubstatusFilter = '';
  $('bookings-substatus-filter').value = '';

  bookingsResource = DEFAULT_BOOKINGS_RESOURCE;
  $('bookings-resource-filter').value = DEFAULT_BOOKINGS_RESOURCE;

  loadBookings();
});

// ── SCHEDULE ───────────────────────────────────────────────────────────────
function getWeekRange(offset = 0) {
  const now = new Date();
  const day = now.getDay();
  const monday = new Date(now);
  monday.setDate(now.getDate() - ((day === 0 ? 7 : day) - 1) + offset * 7);
  monday.setHours(0, 0, 0, 0);
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  sunday.setHours(23, 59, 59, 999);
  return { monday, sunday };
}

// Team members are user-configurable and persisted in settings (see team manager)
let SCHEDULE_RESOURCES = [];

// ── Bookings team-member filter ──────────────────────────────────────────────
function populateBookingsResourceFilter() {
  const sel = $('bookings-resource-filter');
  sel.innerHTML = `<option value="${ALL_MEMBERS_VALUE}">All Members</option>` +
    SCHEDULE_RESOURCES.map(name =>
      `<option value="${esc(name)}">${esc(name)}</option>`
    ).join('');
  sel.value = bookingsResource || ALL_MEMBERS_VALUE;
}
$('bookings-resource-filter').addEventListener('change', e => {
  bookingsResource = e.target.value;
  if (bookingsFilter === 'all' || bookingsSearch) {
    loadAllBookingsForSearch();
  } else {
    loadBookings();
  }
});

// Time-based layout constants — horizontal timeline
const DAY_START_H    = 6;              // 6 AM
const DAY_END_H      = 21;             // 9 PM
const HOURS_PER_DAY  = DAY_END_H - DAY_START_H; // 15
const PX_PER_HOUR    = 200;
const PX_PER_MIN     = PX_PER_HOUR / 60;
const DAY_W          = HOURS_PER_DAY * PX_PER_HOUR;
const TOTAL_W        = 7 * DAY_W;
const ROW_H          = 80;
const WEEKLY_CAPACITY_HOURS = 40; // basis for utilization % (standard work week)

function fmtBookedDuration(minutes) {
  const h = Math.floor(minutes / 60);
  const m = Math.round(minutes % 60);
  if (h && m) return `${h}h ${m}m`;
  if (h) return `${h}h`;
  return `${m}m`;
}

function timeToX(iso, monday) {
  const d = new Date(iso);
  const dayMs = new Date(d.getFullYear(), d.getMonth(), d.getDate()) -
                new Date(monday.getFullYear(), monday.getMonth(), monday.getDate());
  const dayIdx = Math.round(dayMs / 86400000);
  const fromDayStart = (d.getHours() - DAY_START_H) + d.getMinutes() / 60;
  return Math.max(0, (dayIdx * HOURS_PER_DAY + fromDayStart) * PX_PER_HOUR);
}
function bookingWidth(start, end) {
  const mins = (new Date(end) - new Date(start)) / 60000;
  return Math.max(12, mins * PX_PER_MIN);
}

function scrollToToday() {
  if (weekOffset !== 0) return;
  const { monday } = getWeekRange(0);
  const today = new Date();
  const dayIndex = Math.floor((today - monday) / 86400000);
  if (dayIndex < 0 || dayIndex > 6) return;
  const wrap = $('schedule-wrap');
  const dayLeft = dayIndex * DAY_W;
  const minutesIntoDay = Math.max(0, (today.getHours() - DAY_START_H) + today.getMinutes() / 60) * PX_PER_HOUR;
  const center = dayLeft + minutesIntoDay - wrap.clientWidth / 2;
  wrap.scrollLeft = Math.max(0, center);
}

async function renderSchedule(force = false) {
  const { monday, sunday } = getWeekRange(weekOffset);
  const opts = { month: 'short', day: 'numeric' };
  $('week-label').textContent =
    monday.toLocaleDateString(undefined, opts) + ' – ' +
    sunday.toLocaleDateString(undefined, { ...opts, year: 'numeric' });

  const weekKey = `${weekOffset}`;
  if (!cache.schedule || cache.scheduleWeek !== weekKey || force) {
    showState('schedule', 'loading');
    try {
      const from = monday.toISOString();
      const to   = sunday.toISOString();
      const records = await xrmFetch('bookableresourcebooking',
        '?$select=name,starttime,endtime,_resource_value,_bookingstatus_value,_msdyn_workorder_value,bookableresourcebookingid' +
        '&$expand=BookingStatus($select=name)' +
        `&$filter=starttime ge ${from} and starttime le ${to}` +
        '&$orderby=starttime asc&$top=1000'
      );
      cache.schedule     = records;
      cache.scheduleWeek = weekKey;

      // Fetch approved time off for the same week (rendered as black bars, like Dynamics)
      cache.timeoff = [];
      try {
        cache.timeoff = await xrmFetch('msdyn_timeoffrequest',
          '?$select=msdyn_name,msdyn_starttime,msdyn_endtime,_msdyn_resource_value,statecode' +
          `&$filter=msdyn_starttime le ${to} and msdyn_endtime ge ${from} and statecode eq 0` +
          '&$orderby=msdyn_starttime asc&$top=500'
        );
      } catch (e) { console.warn('Time off fetch failed:', e); }

      // Fetch PTO from the SharePoint "Time Off Calendar" (the real source of truth)
      cache.sptimeoff = await fetchSharePointTimeOff(from, to);

      // Fetch customer names for all work orders in this batch
      const woIds = [...new Set(records.map(r => r._msdyn_workorder_value).filter(Boolean))];
      cache.customerMap = {};
      cache.problemMap  = {};
      if (woIds.length) {
        try {
          const CHUNK = 30;
          const chunks = [];
          for (let i = 0; i < woIds.length; i += CHUNK) chunks.push(woIds.slice(i, i + CHUNK));
          const results = await Promise.all(chunks.map(chunk => {
            const filter = chunk.map(id => `msdyn_workorderid eq ${id}`).join(' or ');
            return xrmFetch('msdyn_workorder',
              `?$select=msdyn_workorderid,_msdyn_serviceaccount_value,wc_workorderproblemdescription&$filter=${filter}&$top=${chunk.length}`
            );
          }));
          results.flat().forEach(wo => {
            const name = wo['_msdyn_serviceaccount_value@OData.Community.Display.V1.FormattedValue'] || '';
            if (name) cache.customerMap[wo.msdyn_workorderid] = name;
            if (wo.wc_workorderproblemdescription) cache.problemMap[wo.msdyn_workorderid] = wo.wc_workorderproblemdescription;
          });
        } catch (_) {}
      }
    } catch (e) {
      showState('schedule', 'empty');
      console.error('Schedule error:', e);
      return;
    }
  }

  const customerMap = cache.customerMap || {};
  const problemMap  = cache.problemMap  || {};

  const weekBookings = (cache.schedule || []).filter(r => {
    if (!r.starttime) return false;
    const res = r['_resource_value@OData.Community.Display.V1.FormattedValue'] || '';
    return SCHEDULE_RESOURCES.some(n => res.toLowerCase().includes(n.toLowerCase()));
  });

  const days = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(monday); d.setDate(monday.getDate() + i); return d;
  });

  // Group approved time off by resource (list of {start,end} intervals)
  const timeoffMap = {};
  SCHEDULE_RESOURCES.forEach(r => { timeoffMap[r] = []; });
  (cache.timeoff || []).forEach(t => {
    if (!t.msdyn_starttime || !t.msdyn_endtime) return;
    const res   = t['_msdyn_resource_value@OData.Community.Display.V1.FormattedValue'] || '';
    const match = SCHEDULE_RESOURCES.find(n => res.toLowerCase().includes(n.toLowerCase()));
    if (!match) return;
    timeoffMap[match].push({ start: new Date(t.msdyn_starttime), end: new Date(t.msdyn_endtime), name: t.msdyn_name || 'Time Off' });
  });
  // Merge SharePoint PTO (matched to team members by name)
  (cache.sptimeoff || []).forEach(t => {
    if (!t.name || !t.start || !t.end) return;
    const n = t.name.toLowerCase().trim();
    const match = SCHEDULE_RESOURCES.find(r => r.toLowerCase().trim() === n)
      || SCHEDULE_RESOURCES.find(r => { const p = r.toLowerCase().split(/\s+/); return p.length >= 2 && n.includes(p[0]) && n.includes(p[p.length - 1]); });
    if (!match) return;
    timeoffMap[match].push({ start: t.start, end: t.end, name: t.category || 'PTO' });
  });

  const resourceMap = {};
  SCHEDULE_RESOURCES.forEach(r => { resourceMap[r] = {}; });
  weekBookings.forEach(r => {
    const res   = r['_resource_value@OData.Community.Display.V1.FormattedValue'] || '';
    const match = SCHEDULE_RESOURCES.find(n => res.toLowerCase().includes(n.toLowerCase()));
    if (!match) return;
    const dayKey = new Date(r.starttime).toDateString();
    if (!resourceMap[match][dayKey]) resourceMap[match][dayKey] = [];
    resourceMap[match][dayKey].push(r);
  });

  const today         = new Date().toDateString();
  const grid          = $('schedule-grid');
  const schedBlockMap = {};
  grid.innerHTML      = '';

  // ── Header row ──────────────────────────────────────────────────────────
  let headerHtml = `<div class="sched-header-row">
    <div class="sched-corner">Resource</div>
    <div class="sched-time-header" style="width:${TOTAL_W}px;">`;

  days.forEach((d, di) => {
    const isToday  = d.toDateString() === today;
    const label    = d.toLocaleDateString(undefined, { weekday: 'short', month: 'short', day: 'numeric' });
    const dayLeft  = di * DAY_W;
    headerHtml += `<div class="sched-day-label ${isToday ? 'today' : ''}" style="left:${dayLeft}px;width:${DAY_W}px;">${label}</div>`;
    for (let h = 0; h < HOURS_PER_DAY; h++) {
      const hr   = DAY_START_H + h;
      const tick = hr === 12 ? '12pm' : hr < 12 ? `${hr}am` : `${hr - 12}pm`;
      headerHtml += `<div class="sched-hour-tick" style="left:${dayLeft + h * PX_PER_HOUR}px;">${tick}</div>`;
    }
  });
  headerHtml += `</div></div>`;
  grid.innerHTML += headerHtml;

  // ── Resource rows ────────────────────────────────────────────────────────
  SCHEDULE_RESOURCES.forEach(res => {
    const color = resourceColor(res);

    // Assign overlapping bookings to lanes FIRST so we know the row height
    const allBookings = days.flatMap(d => resourceMap[res][d.toDateString()] || [])
      .sort((a, b) => new Date(a.starttime) - new Date(b.starttime));
    const laneEnds = [];
    const laneAssignments = allBookings.map(b => {
      const start = new Date(b.starttime).getTime();
      const end   = new Date(b.endtime).getTime();
      let lane = laneEnds.findIndex(e => e <= start);
      if (lane === -1) lane = laneEnds.length;
      laneEnds[lane] = end;
      return lane;
    });
    const laneCount = Math.max(1, laneEnds.length);
    const rowHeight = laneCount * ROW_H;

    // Booked time + utilization % for the week
    const bookedMin = allBookings.reduce((sum, b) => {
      const s = new Date(b.starttime).getTime(), e = new Date(b.endtime).getTime();
      return sum + (e > s ? (e - s) / 60000 : 0);
    }, 0);
    const utilPct = Math.round((bookedMin / (WEEKLY_CAPACITY_HOURS * 60)) * 100);

    let rowHtml = `<div class="sched-row">
      <div class="sched-resource-label" style="height:${rowHeight}px;">
        <div class="resource-dot" style="background:${color}"></div>
        <div class="resource-label-text">
          <div class="resource-name">${esc(res)}</div>
          <div class="resource-util">${fmtBookedDuration(bookedMin)} booked · ${utilPct}%</div>
        </div>
      </div>
      <div class="sched-timeline" style="width:${TOTAL_W}px;height:${rowHeight}px;">`;

    // Day backgrounds + grid lines
    days.forEach((d, di) => {
      const isToday = d.toDateString() === today;
      const dayLeft = di * DAY_W;
      if (isToday) {
        rowHtml += `<div style="position:absolute;left:${dayLeft}px;width:${DAY_W}px;top:0;bottom:0;background:rgba(108,192,245,.07);pointer-events:none;"></div>`;
      }
      rowHtml += `<div style="position:absolute;left:${dayLeft}px;top:0;bottom:0;width:1px;background:var(--border);pointer-events:none;"></div>`;
      for (let h = 1; h < HOURS_PER_DAY; h++) {
        rowHtml += `<div style="position:absolute;left:${dayLeft + h * PX_PER_HOUR}px;top:0;bottom:0;width:1px;background:rgba(37,43,59,.5);pointer-events:none;"></div>`;
      }
    });

    // Time off — solid black bar over each day the resource is off (like Dynamics)
    days.forEach((d, di) => {
      const dayStart = new Date(d); dayStart.setHours(0, 0, 0, 0);
      const dayEnd   = new Date(d); dayEnd.setHours(23, 59, 59, 999);
      const off = (timeoffMap[res] || []).find(t => t.start <= dayEnd && t.end >= dayStart);
      if (!off) return;
      const dayLeft = di * DAY_W;
      rowHtml += `<div class="sched-timeoff" title="${esc(off.name)}" style="position:absolute;left:${dayLeft}px;width:${DAY_W}px;top:0;bottom:0;">
        <span class="sched-timeoff-label">Time Off</span>
      </div>`;
    });

    // Booking blocks — each fixed ROW_H tall, positioned in its lane
    allBookings.forEach((b, i) => {
      const lane     = laneAssignments[i];
      const key      = b.bookableresourcebookingid || `${res}-${b.starttime}`;
      schedBlockMap[key] = b;
      const status   = b.BookingStatus?.name || b['_bookingstatus_value@OData.Community.Display.V1.FormattedValue'] || '';
      const cls      = bookingBlockClass(status);
      const time     = `${fmtTime(b.starttime)}–${fmtTime(b.endtime)}`;
      const customer = customerMap[b._msdyn_workorder_value] || '';
      const problem  = problemMap[b._msdyn_workorder_value]  || '';
      const left     = timeToX(b.starttime, monday);
      const width    = bookingWidth(b.starttime, b.endtime);
      const top      = lane * ROW_H + 4;
      const height   = ROW_H - 8;
      rowHtml += `<div class="booking-block ${cls}" data-key="${esc(key)}" data-time="${esc(time)}" data-problem="${esc(problem)}" data-woname="${esc(b.name||'')}" data-status="${esc(status)}"
        style="position:absolute;left:${left}px;width:${width}px;top:${top}px;height:${height}px;cursor:pointer">
        <div class="booking-block-time">${esc(time)}</div>
        <div class="booking-block-title">${esc(customer || b.name || '')}</div>
        ${customer ? `<div class="booking-block-sub">${esc(b.name || '')}</div>` : ''}
      </div>`;
    });

    rowHtml += `</div></div>`;
    grid.innerHTML += rowHtml;
  });

  showState('schedule', 'grid');

  grid.querySelectorAll('.booking-block[data-key]').forEach(el => {
    el.addEventListener('click', () => {
      const b = schedBlockMap[el.dataset.key];
      if (b) openRecord(b);
    });
  });
}

$('week-prev').addEventListener('click',        () => { weekOffset--; renderSchedule(true); });
$('week-next').addEventListener('click',        () => { weekOffset++; renderSchedule(true); });
$('week-today-btn').addEventListener('click',   () => { weekOffset = 0; renderSchedule(true).then(scrollToToday); });
$('schedule-refresh').addEventListener('click', () => renderSchedule(true));

// ── ACCOUNTS ───────────────────────────────────────────────────────────────
const ACCOUNTS_SELECT = 'accountid,name,emailaddress1,telephone1,address1_city';
let accountsRowMap = {};
let accountsSearch = '';
let accountsSearchTimer = null;

async function loadAccounts(force = false) {
  if (cache.accounts && !force) { renderAccounts(cache.accounts); return; }
  showState('accounts', 'loading');
  try {
    const page = await xrmFetchPage('account',
      `?$select=${ACCOUNTS_SELECT}&$orderby=name asc&$top=500`
    );
    cache.accounts = page.entities;
    renderAccounts(cache.accounts);
  } catch (e) { showState('accounts', 'empty'); console.error('Accounts error:', e); }
}

async function searchAccounts(q) {
  showState('accounts', 'loading');
  try {
    const safe = q.replace(/'/g, "''");
    const records = await xrmFetch('account',
      `?$select=${ACCOUNTS_SELECT}&$filter=contains(name,'${safe}')&$orderby=name asc&$top=500`
    );
    renderAccounts(records);
  } catch (e) { showState('accounts', 'empty'); console.error('Accounts search error:', e); }
}

function renderAccounts(records) {
  let rows = [...records].sort((a, b) => {
    const an = (a.name || '').trim(), bn = (b.name || '').trim();
    if (!an && bn) return 1;
    if (an && !bn) return -1;
    return an.localeCompare(bn);
  });
  $('accounts-count').textContent = rows.length;
  if (rows.length === 0) { showState('accounts', 'empty'); return; }
  accountsRowMap = {};
  $('accounts-body').innerHTML = rows.map((r, i) => {
    accountsRowMap[i] = r;
    return `<tr data-idx="${i}" class="row-clickable">
      <td><strong>${esc(r.name||'—')}</strong></td>
      <td class="muted">${esc(r.address1_city||'—')}</td>
      <td class="muted">${esc(r.telephone1||'—')}</td>
      <td class="muted">${esc(r.emailaddress1||'—')}</td>
    </tr>`;
  }).join('');
  $('accounts-body').querySelectorAll('tr[data-idx]').forEach(tr => {
    tr.addEventListener('click', () => openAccountDetail(accountsRowMap[+tr.dataset.idx]));
  });
  showState('accounts', 'table');
}

$('accounts-search').addEventListener('input', e => {
  accountsSearch = e.target.value.trim();
  clearTimeout(accountsSearchTimer);
  accountsSearchTimer = setTimeout(() => {
    if (accountsSearch) searchAccounts(accountsSearch);
    else if (cache.accounts) renderAccounts(cache.accounts);
  }, 300);
});
$('accounts-refresh').addEventListener('click', () => loadAccounts(true));

// ── CONTACTS ───────────────────────────────────────────────────────────────
const CONTACTS_SELECT = 'fullname,emailaddress1,mobilephone,telephone1,jobtitle,_parentcustomerid_value';
let contactsSearch = '';
let contactsSearchTimer = null;

async function loadContacts(force = false) {
  if (cache.contacts && !force) { renderContacts(cache.contacts); return; }
  showState('contacts', 'loading');
  try {
    const page = await xrmFetchPage('contact',
      `?$select=${CONTACTS_SELECT}&$orderby=fullname asc&$top=500`
    );
    cache.contacts = page.entities;
    renderContacts(cache.contacts);
  } catch (e) { showState('contacts', 'empty'); console.error('Contacts error:', e); }
}

async function searchContacts(q) {
  showState('contacts', 'loading');
  try {
    const safe = q.replace(/'/g, "''");
    const records = await xrmFetch('contact',
      `?$select=${CONTACTS_SELECT}&$filter=contains(fullname,'${safe}')&$orderby=fullname asc&$top=500`
    );
    renderContacts(records);
  } catch (e) { showState('contacts', 'empty'); console.error('Contacts search error:', e); }
}

function renderContacts(records) {
  let rows = [...records].sort((a, b) => {
    const an = (a.fullname || '').trim(), bn = (b.fullname || '').trim();
    if (!an && bn) return 1;
    if (an && !bn) return -1;
    return an.localeCompare(bn);
  });
  $('contacts-count').textContent = rows.length;
  if (rows.length === 0) { showState('contacts', 'empty'); return; }
  $('contacts-body').innerHTML = rows.map(r => {
    const account = r['_parentcustomerid_value@OData.Community.Display.V1.FormattedValue'] || '—';
    return `<tr>
      <td><strong>${esc(r.fullname||'—')}</strong></td>
      <td class="muted">${esc(account)}</td>
      <td class="muted">${esc(r.mobilephone||r.telephone1||'—')}</td>
      <td class="muted">${esc(r.emailaddress1||'—')}</td>
      <td class="muted">${esc(r.jobtitle||'—')}</td>
    </tr>`;
  }).join('');
  showState('contacts', 'table');
}

$('contacts-search').addEventListener('input', e => {
  contactsSearch = e.target.value.trim();
  clearTimeout(contactsSearchTimer);
  contactsSearchTimer = setTimeout(() => {
    if (contactsSearch) searchContacts(contactsSearch);
    else if (cache.contacts) renderContacts(cache.contacts);
  }, 300);
});
$('contacts-refresh').addEventListener('click', () => loadContacts(true));

// ── TEAM ─────────────────────────────────────────────────────────────────────
async function loadTeam(force = false) {
  if (cache.team && !force) { renderTeam(cache.team); return; }
  showState('team', 'loading');
  try {
    const nowIso = new Date().toISOString();
    const filter = '(' + SCHEDULE_RESOURCES.map(n => `Resource/name eq '${n.replace(/'/g, "''")}'`).join(' or ') + ')' +
      ` and starttime le ${nowIso} and endtime ge ${nowIso}`;
    const bookings = await xrmFetch('bookableresourcebooking',
      `?$select=name&$expand=BookingStatus($select=name),Resource($select=name)&$filter=${filter}`
    );
    const priority = { 'Traveling': 3, 'In Progress': 2, 'Scheduled': 1 };
    const statusByName = {};
    bookings.forEach(b => {
      const name = b.Resource?.name;
      if (!name) return;
      const raw = (b.BookingStatus?.name || '').toLowerCase();
      let label = 'Scheduled';
      if (raw.includes('travel')) label = 'Traveling';
      else if (raw.includes('progress')) label = 'In Progress';
      const existing = statusByName[name];
      if (!existing || priority[label] > priority[existing]) statusByName[name] = label;
    });
    cache.team = SCHEDULE_RESOURCES.map(name => ({ name, status: statusByName[name] || 'Free' }));
    renderTeam(cache.team);
  } catch (e) {
    showState('team', 'empty');
    console.error('Team error:', e);
  }
}

function renderTeam(members) {
  $('team-count').textContent = members.length;
  if (members.length === 0) { showState('team', 'empty'); return; }
  $('team-body').innerHTML = members.map((m, i) => `<tr data-idx="${i}" class="row-clickable">
    <td><strong>${esc(m.name)}</strong></td>
    <td>${statusBadge(m.status === 'Free' ? 'Free' : m.status)}</td>
  </tr>`).join('');
  $('team-body').querySelectorAll('tr[data-idx]').forEach(tr => {
    tr.addEventListener('click', () => {
      const m = members[+tr.dataset.idx];
      window.api.openTeamMember(m.name, orgUrl, m.name);
    });
  });
  showState('team', 'table');
}

$('team-refresh').addEventListener('click', () => loadTeam(true));

// ── Team manager (add/remove members, persisted) ─────────────────────────────
let teamDraft = [];
function openTeamManager(firstRun = false) {
  teamDraft = [...SCHEDULE_RESOURCES];
  $('team-modal-sub').textContent = firstRun
    ? 'Welcome! Add your team members by name, exactly as they appear in Dynamics. These drive the Schedule Board, Team, and bookings filters.'
    : 'Add the resource names exactly as they appear in Dynamics. These drive the Schedule Board, Team, and bookings filters.';
  renderTeamDraft();
  $('team-modal').classList.remove('hidden');
  $('team-modal-input').focus();
}
function renderTeamDraft() {
  const list = $('team-modal-list');
  if (!teamDraft.length) {
    list.innerHTML = '<div class="team-mgr-empty">No team members yet — add some below.</div>';
    return;
  }
  list.innerHTML = teamDraft.map((name, i) =>
    `<div class="team-mgr-item"><span>${esc(name)}</span><button data-rm="${i}" title="Remove">&times;</button></div>`
  ).join('');
  list.querySelectorAll('[data-rm]').forEach(btn => {
    btn.addEventListener('click', () => { teamDraft.splice(+btn.dataset.rm, 1); renderTeamDraft(); });
  });
}
function addTeamDraftMember() {
  const input = $('team-modal-input');
  const name = input.value.trim();
  if (!name) return;
  if (teamDraft.some(n => n.toLowerCase() === name.toLowerCase())) { input.value = ''; return; }
  teamDraft.push(name);
  input.value = '';
  renderTeamDraft();
  input.focus();
}
$('team-edit').addEventListener('click', () => openTeamManager(false));
$('team-modal-add').addEventListener('click', addTeamDraftMember);
$('team-modal-input').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); addTeamDraftMember(); } });
$('team-modal-close').addEventListener('click', () => $('team-modal').classList.add('hidden'));
$('team-modal').addEventListener('click', e => { if (e.target.id === 'team-modal') $('team-modal').classList.add('hidden'); });
$('team-modal-save').addEventListener('click', async () => {
  appSettings.team = [...teamDraft];
  await persistSettings();
  // If the previously-selected bookings resource was removed, fall back to default
  const removedCurrent = bookingsResource !== ALL_MEMBERS_VALUE && !teamDraft.includes(bookingsResource);
  applyTeam(appSettings.team);
  if (removedCurrent) bookingsResource = DEFAULT_BOOKINGS_RESOURCE;
  // Team list affects schedule, team, and bookings — invalidate their caches
  cache.schedule = cache.scheduleWeek = cache.team = null;
  cache.bookingsAllByResource = {};
  $('team-modal').classList.add('hidden');
  if (xrmReady) loadTab(activeTab);
});

// ── ASSETS ───────────────────────────────────────────────────────────────────
let assetsSearch = '', assetsSearchTimer = null, assetsRowMap = {};

async function searchAssets(q) {
  showState('assets', 'loading');
  try {
    const safe = q.replace(/'/g, "''");
    const records = await xrmFetch('msdyn_customerasset',
      `?$select=msdyn_customerassetid,msdyn_name,wc_assettag,msdyn_assettag,wc_seriallotnumber,wc_knumber` +
      `&$filter=contains(wc_assettag,'${safe}') or contains(msdyn_assettag,'${safe}') or contains(wc_seriallotnumber,'${safe}')` +
      `&$orderby=msdyn_name asc&$top=200`
    );
    renderAssets(records);
  } catch (e) { showState('assets', 'empty'); console.error('Assets search error:', e); }
}

function renderAssets(records) {
  $('assets-count').textContent = records.length;
  if (records.length === 0) { showState('assets', 'empty'); return; }
  assetsRowMap = {};
  $('assets-body').innerHTML = records.map((r, i) => {
    assetsRowMap[i] = r;
    const tag = r.wc_assettag || r.msdyn_assettag || '—';
    return `<tr data-idx="${i}" class="row-clickable">
      <td><strong>${esc(r.msdyn_name || '—')}</strong></td>
      <td class="muted">${esc(tag)}</td>
      <td class="muted">${esc(r.wc_knumber || '—')}</td>
      <td class="muted">${esc(r.wc_seriallotnumber || '—')}</td>
    </tr>`;
  }).join('');
  $('assets-body').querySelectorAll('tr[data-idx]').forEach(tr => {
    tr.addEventListener('click', () => openAssetDetail(assetsRowMap[+tr.dataset.idx]));
  });
  showState('assets', 'table');
}

const ASSETS_DETAIL_SELECT = 'msdyn_customerassetid,msdyn_name,wc_assettag,msdyn_assettag,wc_seriallotnumber,' +
  'statuscode,wc_knumber,msdyn_manufacturingdate,_msdyn_parentasset_value,_wc_warrantyservicecontract_value,' +
  '_msdyn_masterasset_value,_msdyn_product_value,_wc_manufacturer_value,_msdyn_workorderproduct_value';

function fv(obj, field) {
  if (!obj) return '';
  return obj[`${field}@OData.Community.Display.V1.FormattedValue`] || obj[field] || '';
}

async function openAssetDetail(row) {
  const assetId = row?.msdyn_customerassetid;
  if (!assetId) return;
  const modal = $('asset-modal');
  const body  = $('asset-modal-body');
  body.innerHTML = '<div class="loading-state"><div class="spinner"></div></div>';
  modal.classList.remove('hidden');

  try {
    const records = await xrmFetch('msdyn_customerasset', `?$select=${ASSETS_DETAIL_SELECT}&$filter=msdyn_customerassetid eq ${assetId}`);
    const a = records[0];
    if (!a) { body.innerHTML = '<div class="am-empty">Asset not found.</div>'; return; }

    const tag = a.wc_assettag || a.msdyn_assettag || '—';

    body.innerHTML = `
      <div class="am-title">${esc(a.msdyn_name || '—')}</div>
      <div class="am-sub">Asset Tag: ${esc(tag)}</div>

      <div class="am-grid">
        <div><div class="am-field-label">Asset Status</div><div class="am-field-value">${esc(fv(a,'statuscode')||'—')}</div></div>
        <div><div class="am-field-label">K Number</div><div class="am-field-value">${esc(a.wc_knumber||'—')}</div></div>
        <div><div class="am-field-label">Parent Asset</div><div class="am-field-value">${esc(fv(a,'_msdyn_parentasset_value')||'—')}</div></div>
        <div><div class="am-field-label">Manufacturing Date</div><div class="am-field-value">${esc(fmtDate(a.msdyn_manufacturingdate)||'—')}</div></div>
        <div><div class="am-field-label">Warranty Service Contract</div><div class="am-field-value">${esc(fv(a,'_wc_warrantyservicecontract_value')||'—')}</div></div>
        <div><div class="am-field-label">Top-Level Asset</div><div class="am-field-value">${esc(fv(a,'_msdyn_masterasset_value')||'—')}</div></div>
        <div><div class="am-field-label">Product</div><div class="am-field-value">${esc(fv(a,'_msdyn_product_value')||'—')}</div></div>
        <div><div class="am-field-label">Manufacturer</div><div class="am-field-value">${esc(fv(a,'_wc_manufacturer_value')||'—')}</div></div>
        <div><div class="am-field-label">Work Order Product</div><div class="am-field-value">${esc(fv(a,'_msdyn_workorderproduct_value')||'—')}</div></div>
        <div><div class="am-field-label">Serial/Lot #</div><div class="am-field-value">${esc(a.wc_seriallotnumber||'—')}</div></div>
      </div>
    `;
  } catch (e) {
    body.innerHTML = '<div class="am-empty">Failed to load asset details.</div>';
    console.error('Asset detail error:', e);
  }
}

$('asset-modal-close').addEventListener('click', () => $('asset-modal').classList.add('hidden'));
$('asset-modal').addEventListener('click', e => { if (e.target.id === 'asset-modal') $('asset-modal').classList.add('hidden'); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') $('asset-modal').classList.add('hidden'); });

$('assets-search').addEventListener('input', e => {
  assetsSearch = e.target.value.trim();
  clearTimeout(assetsSearchTimer);
  assetsSearchTimer = setTimeout(() => {
    if (assetsSearch.length >= 2) searchAssets(assetsSearch);
    else showState('assets', 'empty');
  }, 300);
});
$('assets-refresh').addEventListener('click', () => {
  if (assetsSearch.length >= 2) searchAssets(assetsSearch);
});

// ── Show/hide states ───────────────────────────────────────────────────────
function showState(tab, state) {
  const loading = $(`${tab}-loading`);
  const empty   = $(`${tab}-empty`);
  const table   = $(`${tab}-table`);
  const grid    = $('schedule-grid');

  [loading, empty, table, grid].forEach(el => el?.classList.add('hidden'));

  if (state === 'loading') loading?.classList.remove('hidden');
  else if (state === 'empty') empty?.classList.remove('hidden');
  else if (state === 'table') table?.classList.remove('hidden');
  else if (state === 'grid')  grid?.classList.remove('hidden');
}

// ── Account detail modal ──────────────────────────────────────────────────
async function openAccountDetail(row) {
  const accountId = row?.accountid;
  if (!accountId) return;
  const modal = $('account-modal');
  const body  = $('account-modal-body');
  body.innerHTML = '<div class="loading-state"><div class="spinner"></div></div>';
  modal.classList.remove('hidden');

  try {
    const [accounts, contacts, workOrders] = await Promise.all([
      xrmFetch('account',
        `?$select=name,emailaddress1,telephone1,address1_line1,address1_city,address1_stateorprovince,address1_postalcode,websiteurl,description&$filter=accountid eq ${accountId}`
      ),
      xrmFetch('contact',
        `?$select=contactid,fullname,emailaddress1,mobilephone,telephone1,jobtitle&$filter=_parentcustomerid_value eq ${accountId}&$orderby=fullname asc`
      ),
      xrmFetch('msdyn_workorder',
        `?$select=msdyn_workorderid,msdyn_name,msdyn_systemstatus,createdon&$filter=_msdyn_serviceaccount_value eq ${accountId}&$orderby=createdon desc&$top=10`
      ).catch(() => [])
    ]);

    const a = accounts[0];
    if (!a) { body.innerHTML = '<div class="am-empty">Account not found.</div>'; return; }

    const address = [a.address1_line1, a.address1_city, a.address1_stateorprovince, a.address1_postalcode]
      .filter(Boolean).join(', ');

    body.innerHTML = `
      <div class="am-title">${esc(a.name || '—')}</div>
      <div class="am-sub">${esc(address || 'No address on file')}</div>

      <div class="am-grid">
        <div><div class="am-field-label">Phone</div><div class="am-field-value">${esc(a.telephone1||'—')}</div></div>
        <div><div class="am-field-label">Email</div><div class="am-field-value">${esc(a.emailaddress1||'—')}</div></div>
        <div><div class="am-field-label">Website</div><div class="am-field-value">${esc(a.websiteurl||'—')}</div></div>
      </div>

      <div class="am-section-title">Contacts (${contacts.length})</div>
      <div class="am-list">
        ${contacts.length ? contacts.map(c => `
          <div class="am-list-item am-clickable" data-contact-id="${esc(c.contactid)}">
            <div class="am-list-item-title">${esc(c.fullname||'—')}</div>
            <div class="am-list-item-sub">${esc(c.jobtitle || '')}${c.jobtitle && (c.mobilephone||c.telephone1||c.emailaddress1) ? ' · ' : ''}${esc(c.mobilephone||c.telephone1||'')}${(c.mobilephone||c.telephone1) && c.emailaddress1 ? ' · ' : ''}${esc(c.emailaddress1||'')}</div>
          </div>`).join('') : '<div class="am-empty">No contacts on file.</div>'}
      </div>

      <div class="am-section-title">Recent Work Orders</div>
      <div class="am-list">
        ${workOrders.length ? workOrders.map(w => `
          <div class="am-list-item am-clickable" data-wo-id="${esc(w.msdyn_workorderid)}">
            <div class="am-list-item-title">${esc(w.msdyn_name||'—')}</div>
            <div class="am-list-item-sub">${esc(w['msdyn_systemstatus@OData.Community.Display.V1.FormattedValue']||'')}${fmtDate(w.createdon) !== '—' ? ' · ' + fmtDate(w.createdon) : ''}</div>
          </div>`).join('') : '<div class="am-empty">No work orders on file.</div>'}
      </div>
    `;

    body.querySelectorAll('[data-contact-id]').forEach(el => {
      el.addEventListener('click', () => window.api.openContact(el.dataset.contactId, orgUrl, 'Contact'));
    });
    body.querySelectorAll('[data-wo-id]').forEach(el => {
      el.addEventListener('click', () => window.api.openWorkOrderDirect(el.dataset.woId, orgUrl, 'Work Order'));
    });
  } catch (e) {
    body.innerHTML = '<div class="am-empty">Failed to load account details.</div>';
    console.error('Account detail error:', e);
  }
}

$('account-modal-close').addEventListener('click', () => $('account-modal').classList.add('hidden'));
$('account-modal').addEventListener('click', e => { if (e.target.id === 'account-modal') $('account-modal').classList.add('hidden'); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') $('account-modal').classList.add('hidden'); });

// ── Schedule tooltip ──────────────────────────────────────────────────────
const schedTip = $('sched-tooltip');
document.addEventListener('mouseover', e => {
  const block = e.target.closest('.booking-block[data-key]');
  if (!block) return;
  const time    = block.dataset.time    || '';
  const wo      = block.dataset.woname  || '';
  const problem = block.dataset.problem || '';
  const status  = block.dataset.status  || '';
  schedTip.innerHTML =
    `<div class="tt-wo">${wo}</div>` +
    `<div class="tt-time">${time}</div>` +
    (status  ? `<div class="tt-status">${status}</div>`  : '') +
    (problem ? `<div class="tt-problem">${problem}</div>` : '');
  schedTip.classList.add('visible');
});
document.addEventListener('mousemove', e => {
  if (!schedTip.classList.contains('visible')) return;
  const x = e.clientX + 14, y = e.clientY + 14;
  const overRight = x + 270 > window.innerWidth;
  schedTip.style.left = (overRight ? e.clientX - 270 : x) + 'px';
  schedTip.style.top  = Math.min(y, window.innerHeight - schedTip.offsetHeight - 8) + 'px';
});
document.addEventListener('mouseout', e => {
  if (e.target.closest('.booking-block[data-key]')) schedTip.classList.remove('visible');
});

// ── Quick-create: internal "Travel to Home" work order ───────────────────────
// Template values captured from a real travel work order (WO-70403).
const TRAVEL_WO = {
  serviceaccount:  'YOUR-SERVICE-ACCOUNT-GUID', // service account
  incidenttype:    'YOUR-INCIDENT-TYPE-GUID', // incident type
  priority:        'YOUR-PRIORITY-GUID', // Low
  pricelist:       'YOUR-PRICE-LIST-GUID', // Mid
  serviceterritory:'YOUR-SERVICE-TERRITORY-GUID', // service territory
  problem:         'Travel to home.',
};
const TRAVEL_BOOKING_STATUS  = 'YOUR-SCHEDULED-BOOKING-STATUS-GUID'; // Completed
const TRAVEL_BOOKING_SETUPMD = 'YOUR-BOOKING-SETUP-METADATA-GUID'; // msdyn_workorder booking setup metadata
const TRAVEL_DURATION_MIN    = 60;

async function currentUserResourceId() {
  try {
    const raw = await apiWv.executeJavaScript('Xrm.Utility.getGlobalContext().userSettings.userId');
    const uid = String(raw || '').replace(/[{}]/g, '');
    if (!uid) return null;
    const res = await xrmFetch('bookableresource', `?$select=bookableresourceid&$filter=_userid_value eq ${uid}&$top=1`);
    return res[0]?.bookableresourceid || null;
  } catch (_) { return null; }
}

async function createTravelWorkOrder() {
  const btn = $('travel-wo-btn');
  if (!xrmReady) { toast('Not connected to Dynamics yet', true); return; }
  btn.disabled = true; const label = btn.textContent; btn.textContent = 'Creating…';
  try {
    // 1) Create the work order (entity-set names resolved from metadata to avoid guessing)
    const nav = await getNavPropMap('msdyn_workorder');
    const woPayload = { wc_workorderproblemdescription: TRAVEL_WO.problem };
    const bind = async (attr, id) => {
      const rel = nav[attr];
      if (!rel) return;
      const set = await entitySetOf(rel.target);
      woPayload[`${rel.nav}@odata.bind`] = `/${set}(${id})`;
    };
    await bind('msdyn_serviceaccount',     TRAVEL_WO.serviceaccount);
    await bind('msdyn_primaryincidenttype',TRAVEL_WO.incidenttype);
    await bind('msdyn_priority',           TRAVEL_WO.priority);
    await bind('msdyn_pricelist',          TRAVEL_WO.pricelist);
    await bind('msdyn_serviceterritory',   TRAVEL_WO.serviceterritory);
    // Reported By Contact = the signed-in user's contact record (e.g. "Your Name")
    try {
      const uname = String(await apiWv.executeJavaScript('Xrm.Utility.getGlobalContext().getUserName()') || '').trim();
      if (uname && nav.msdyn_reportedbycontact) {
        const contacts = await xrmFetch('contact', `?$select=contactid&$filter=fullname eq '${uname.replace(/'/g,"''")}'&$top=1`);
        if (contacts[0]) await bind('msdyn_reportedbycontact', contacts[0].contactid);
      }
    } catch (_) {}
    const woId = await xrmCreate('msdyn_workorder', woPayload);
    let woName = 'work order';
    try { woName = (await xrmRetrieve('msdyn_workorder', woId, '?$select=msdyn_name')).msdyn_name || woName; } catch (_) {}

    // 2) Create the booking — replicate the Dynamics booking form exactly:
    //    status "Scheduled", arrival = start, resource = current user, and let the
    //    platform assign Booking Setup Metadata (do NOT bind it, which triggers the
    //    append-permission check that blocked the earlier attempt).
    let bookingNote = '';
    try {
      const resourceId = await currentUserResourceId();
      if (!resourceId) throw new Error('Could not find your bookable resource');
      const scheduled = await xrmFetch('bookingstatus', `?$select=bookingstatusid&$filter=name eq 'Scheduled'&$orderby=createdon asc&$top=1`);
      const schedStatusId = scheduled[0]?.bookingstatusid;
      const bnav = await getNavPropMap('bookableresourcebooking');
      const start = new Date();
      const end   = new Date(start.getTime() + TRAVEL_DURATION_MIN * 60000);
      const bPayload = {
        starttime: start.toISOString(),
        endtime:   end.toISOString(),
        msdyn_actualarrivaltime: start.toISOString(),
        duration:  TRAVEL_DURATION_MIN,
      };
      const bbind = async (attr, id) => {
        const rel = bnav[attr];
        if (!rel || !id) return;
        const set = await entitySetOf(rel.target);
        bPayload[`${rel.nav}@odata.bind`] = `/${set}(${id})`;
      };
      await bbind('msdyn_workorder', woId);
      await bbind('resource',        resourceId);
      await bbind('bookingstatus',   schedStatusId);
      await xrmCreate('bookableresourcebooking', bPayload);
      bookingNote = ' + booking';
    } catch (be) {
      console.warn('Travel booking failed:', be.message);
      bookingNote = ' (work order only — booking must be added manually: ' + be.message.slice(0, 80) + ')';
    }

    toast(`Created ${woName}${bookingNote}`);
    if (activeTab === 'bookings') loadBookings(true);
    if (activeTab === 'schedule') renderSchedule(true);
  } catch (e) {
    toast('Failed to create travel work order: ' + e.message, true);
  } finally {
    btn.disabled = false; btn.textContent = label;
  }
}
$('travel-wo-btn').addEventListener('click', createTravelWorkOrder);

// ── Auto-update ──────────────────────────────────────────────────────────────
let pendingUpdateAsset = null;
async function checkForUpdates() {
  if (!window.api.checkForUpdate) return;
  try {
    const r = await window.api.checkForUpdate();
    if (r?.ok && r.updateAvailable && r.asset) {
      pendingUpdateAsset = r.asset;
      $('update-toast-sub').textContent = `Version ${r.latest.replace(/^v/,'')} is ready (you have ${r.current}).`;
      $('update-toast').classList.remove('hidden');
    }
  } catch (_) {}
}
window.api.onUpdateProgress?.(p => {
  $('update-toast-progress').classList.remove('hidden');
  $('update-toast-bar').style.width = p + '%';
});
$('update-toast-dismiss').addEventListener('click', () => $('update-toast').classList.add('hidden'));
$('update-toast-install').addEventListener('click', async () => {
  if (!pendingUpdateAsset) return;
  const btn = $('update-toast-install');
  btn.disabled = true; btn.textContent = 'Downloading…';
  $('update-toast-dismiss').disabled = true;
  $('update-toast-sub').textContent = 'Downloading update — the app will restart automatically.';
  $('update-toast-progress').classList.remove('hidden');
  try {
    await window.api.applyUpdate(pendingUpdateAsset);
    // App will relaunch via the updater; if we get here, it's still working.
    btn.textContent = 'Installing…';
  } catch (e) {
    $('update-toast-sub').textContent = 'Update failed: ' + e.message;
    btn.disabled = false; btn.textContent = 'Retry';
    $('update-toast-dismiss').disabled = false;
  }
});

// ── Boot ───────────────────────────────────────────────────────────────────
init();
checkForUpdates();
window.api.getVersion?.().then(v => { if (v) $('app-version').textContent = 'v' + v; });
