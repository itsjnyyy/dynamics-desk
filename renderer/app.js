'use strict';

// ── Globals ────────────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);
const apiWv   = $('api-wv');
const loginWv = $('login-wv');
let orgUrl = '';
let xrmReady = false;
let activeTab = 'bookings';
let weekOffset = 0;

const cache = { bookings: null, schedule: null, scheduleWeek: null, accounts: null, contacts: null };

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
async function init() {
  const s = await window.api.getSettings();
  if (s?.orgUrl) { orgUrl = s.orgUrl.replace(/\/$/, ''); startApp(); }
  else $('setup-overlay').classList.remove('hidden');
}

$('setup-btn').addEventListener('click', async () => {
  const val = $('setup-url').value.trim().replace(/\/$/, '');
  $('setup-error').textContent = '';
  if (!val.startsWith('http')) { $('setup-error').textContent = 'Enter a valid URL starting with https://'; return; }
  await window.api.saveSettings({ orgUrl: val });
  orgUrl = val;
  startApp();
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
  for (let i = 0; i < 40; i++) {
    try {
      const ready = await apiWv.executeJavaScript('typeof Xrm !== "undefined" && typeof Xrm.WebApi !== "undefined"');
      if (ready) {
        xrmReady = true;
        try {
          const uname = await apiWv.executeJavaScript('Xrm.Utility.getGlobalContext().getUserName()');
          if (uname) $('user-name').textContent = uname;
        } catch(_) {}
        return;
      }
    } catch (_) {}
    await sleep(500);
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
// Set this to your own resource name before running
const DEFAULT_BOOKINGS_RESOURCE = 'Your Name';
let bookingsResource = DEFAULT_BOOKINGS_RESOURCE;
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
  if (bookingsResource === ALL_MEMBERS_VALUE) {
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
    return `<tr class="clickable-row" data-idx="${i}">
      <td>${fmtDateTime(r.starttime)}</td>
      <td class="muted">${fmtDateTime(r.endtime)}</td>
      <td>${esc(customer)}</td>
      <td>${esc(r.name || '—')}</td>
      <td>${statusBadge(status)}</td>
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
    if (bookingsFilter === 'all') {
      loadAllBookingsForSearch();
    } else {
      const monthKey = `month:${bookingsMonthOffset}`;
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
        for (let i = 0; i < newWoIds.length; i += CHUNK) chunks.push(newWoIds.slice(i, i + CHUNK));
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

// List the technician/resource names you want shown on the schedule board and team tab
const SCHEDULE_RESOURCES = [
  'Resource One','Resource Two','Resource Three'
];

// ── Bookings team-member filter ──────────────────────────────────────────────
(function initBookingsResourceFilter() {
  const sel = $('bookings-resource-filter');
  sel.innerHTML = `<option value="${ALL_MEMBERS_VALUE}">All Members</option>` +
    SCHEDULE_RESOURCES.map(name =>
      `<option value="${esc(name)}">${esc(name)}</option>`
    ).join('');
  sel.value = bookingsResource;
  sel.addEventListener('change', e => {
    bookingsResource = e.target.value;
    if (bookingsFilter === 'all' || bookingsSearch) {
      loadAllBookingsForSearch();
    } else {
      loadBookings();
    }
  });
})();

// Time-based layout constants — horizontal timeline
const DAY_START_H    = 6;              // 6 AM
const DAY_END_H      = 21;             // 9 PM
const HOURS_PER_DAY  = DAY_END_H - DAY_START_H; // 15
const PX_PER_HOUR    = 200;
const PX_PER_MIN     = PX_PER_HOUR / 60;
const DAY_W          = HOURS_PER_DAY * PX_PER_HOUR;
const TOTAL_W        = 7 * DAY_W;
const ROW_H          = 80;

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

    let rowHtml = `<div class="sched-row">
      <div class="sched-resource-label" style="height:${rowHeight}px;">
        <div class="resource-dot" style="background:${color}"></div>${esc(res)}
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

// ── Boot ───────────────────────────────────────────────────────────────────
init();
