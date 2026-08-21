const params    = new URLSearchParams(window.location.hash.slice(1));
const bookingId = decodeURIComponent(params.get('bid') || '');
const directWoId = decodeURIComponent(params.get('wo') || '');
const orgUrl    = decodeURIComponent(params.get('org') || '');

// Route Web API calls through the main window's already-warm Dynamics session
// (see main.js) instead of loading the heavy shell in our own webview. The shim
// keeps every existing apiWv.executeJavaScript(...) call site unchanged.
const apiWv = { executeJavaScript: (script) => window.api.xrmExec(script) };

const $   = id => document.getElementById(id);
const esc = s  => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

function fmtDate(iso) {
  if (!iso) return '—';
  return new Date(iso).toLocaleString(undefined, { month:'short', day:'numeric', year:'numeric', hour:'2-digit', minute:'2-digit' });
}
function isoToLocal(iso) {
  if (!iso) return '';
  const d = new Date(iso), p = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}`;
}
function fmtDuration(mins) {
  if (mins == null) return '—';
  const h = Math.floor(mins/60), m = mins%60;
  return h ? `${h}h ${m}m` : `${m}m`;
}
function badgeClass(name) {
  const s = (name||'').toLowerCase();
  if (s.includes('progress'))  return 'badge-inprogress';
  if (s.includes('travel'))    return 'badge-traveling';
  if (s.includes('complet'))   return 'badge-completed';
  if (s.includes('cancel'))    return 'badge-canceled';
  if (s.includes('schedul'))   return 'badge-scheduled';
  return 'badge-default';
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// ── Xrm bridge ────────────────────────────────────────────────────────────────
async function waitForXrm() {
  // The main window's session is normally already warm, so this returns almost
  // immediately. ~30s timeout in case the main window is still connecting.
  for (let i = 0; i < 200; i++) {
    try { if (await window.api.xrmReady()) return; } catch(_) {}
    await sleep(150);
  }
  throw new Error('Could not connect to Dynamics — is your session active?');
}
async function xrmGet(entity, id, qs) {
  const q = qs.replace(/"/g,'\\"');
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{return JSON.stringify(await Xrm.WebApi.retrieveRecord("${entity}","${id}","${q}"));}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
  return r;
}
async function xrmList(entity, qs) {
  const q = qs.replace(/"/g,'\\"');
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{return JSON.stringify((await Xrm.WebApi.retrieveMultipleRecords("${entity}","${q}")).entities);}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (!Array.isArray(r)) throw new Error(r?.__err||'Unknown error');
  return r;
}
window.__dumpFields = async function(entityLogicalName, prefix) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes?$select=LogicalName,DisplayName,AttributeType`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  if (r?.__err) { console.error(r.__err); return; }
  let out = r.value.map(a => ({
    LogicalName: a.LogicalName,
    Label: a.DisplayName?.UserLocalizedLabel?.Label || '',
    Type: a.AttributeType
  })).sort((a,b) => a.LogicalName.localeCompare(b.LogicalName));
  if (prefix) out = out.filter(a => (a.LogicalName.startsWith(prefix) || a.LogicalName.includes(prefix) || a.Label.toLowerCase().includes(prefix.toLowerCase())) && a.Label);
  console.log(JSON.stringify(out));
  return out;
};
window.__dumpWOProductFields = async function(prefix) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='msdyn_workorderproduct')/Attributes?$select=LogicalName,DisplayName,AttributeType`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  if (r?.__err) { console.error(r.__err); return; }
  let out = r.value.map(a => ({
    LogicalName: a.LogicalName,
    Label: a.DisplayName?.UserLocalizedLabel?.Label || '',
    Type: a.AttributeType
  })).sort((a,b) => a.LogicalName.localeCompare(b.LogicalName));
  if (prefix) out = out.filter(a => (a.LogicalName.startsWith(prefix) || a.LogicalName.includes(prefix) || a.Label.toLowerCase().includes(prefix.toLowerCase())) && a.Label);
  console.log(JSON.stringify(out));
  return out;
};
window.__dumpOptionSet = async function(attributeLogicalName) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='msdyn_workorderproduct')/Attributes(LogicalName='${attributeLogicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  if (r?.__err) { console.error(r.__err); return; }
  const out = (r.OptionSet?.Options || []).map(o => ({ Value: o.Value, Label: o.Label?.UserLocalizedLabel?.Label || '' }));
  console.log(JSON.stringify(out, null, 2));
  return out;
};
// Diagnostics: find the target entity of a lookup, and dump fields of any entity
window.__lookupTargets = async function(entity, attr) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entity}')/Attributes(LogicalName='${attr}')/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=Targets`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  console.log('Targets:', JSON.stringify(r.Targets || r));
  return r.Targets || r;
};
window.__dumpEntity = async function(entity, prefix) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entity}')/Attributes?$select=LogicalName,DisplayName,AttributeType`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  if (!Array.isArray(r?.value)) { console.error('err', json.slice(0,300)); return; }
  let out = r.value.map(a => ({ LogicalName: a.LogicalName, Label: a.DisplayName?.UserLocalizedLabel?.Label || '', Type: a.AttributeType }))
    .sort((a,b)=>a.LogicalName.localeCompare(b.LogicalName));
  if (prefix) out = out.filter(a => (a.LogicalName.includes(prefix) || a.Label.toLowerCase().includes(prefix.toLowerCase())) && a.Label);
  console.log(JSON.stringify(out));
  return out;
};
window.__dumpOptionSetsOf = async function(entity, ...attrs) {
  const result = {};
  for (const attr of attrs) {
    const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entity}')/Attributes(LogicalName='${attr}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet`;
    const json = await apiWv.executeJavaScript(
      `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
    );
    const r = JSON.parse(json);
    result[attr] = (r.OptionSet?.Options || []).map(o => ({ Value: o.Value, Label: o.Label?.UserLocalizedLabel?.Label || '' }));
  }
  console.log(JSON.stringify(result, null, 2));
  return result;
};
async function getLookupNavProperty(entityLogicalName, lookupLogicalName) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/ManyToOneRelationships?$filter=ReferencingAttribute eq '${lookupLogicalName}'&$select=ReferencingEntityNavigationPropertyName`;
  const json = await apiWv.executeJavaScript(
    `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
  );
  const r = JSON.parse(json);
  if (r?.__err) throw new Error(r.__err);
  const name = r.value?.[0]?.ReferencingEntityNavigationPropertyName;
  if (!name) throw new Error(`Could not resolve navigation property for ${lookupLogicalName}`);
  return name;
}
async function xrmUpdate(entity, id, data) {
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{const __d=${JSON.stringify(data)};await Xrm.WebApi.updateRecord("${entity}","${id}",__d);return JSON.stringify({ok:1});}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
}

// The "Summary" box now maps to the org's new "Booking Summary" field (which
// replaced "Work Order Summary" in Dynamics). Its logical name / host entity
// isn't hard-coded — we resolve it from metadata so a wrong guess can never
// silently drop a tech's notes. Resolved to { entity, logical } and cached.
let summaryField;          // undefined = not tried, null = not found, obj = resolved
async function resolveSummaryField() {
  if (summaryField !== undefined) return summaryField;
  const norm = s => (s || '').toLowerCase().replace(/[^a-z]/g, '');
  const scan = async (entity) => {
    const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entity}')/Attributes?$select=LogicalName,DisplayName,AttributeType`;
    try {
      const json = await apiWv.executeJavaScript(
        `fetch(${JSON.stringify(url)}, {headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`
      );
      const r = JSON.parse(json);
      if (!Array.isArray(r?.value)) return null;
      const textFields = r.value.filter(a => a.AttributeType === 'Memo' || a.AttributeType === 'String');
      const label = a => norm(a.DisplayName?.UserLocalizedLabel?.Label);
      // The editable field is labeled "Booking Summary Notes" in this org. Prefer
      // that exact label, then a plain "Booking Summary", then anything containing it.
      const hit = textFields.find(a => label(a) === 'bookingsummarynotes')
               || textFields.find(a => label(a) === 'bookingsummary')
               || textFields.find(a => label(a).includes('bookingsummarynotes'))
               || textFields.find(a => label(a).includes('bookingsummary'));
      return hit ? { entity, logical: hit.LogicalName } : null;
    } catch (_) { return null; }
  };
  // Booking Summary lives on the work order form; check the booking entity too.
  summaryField = (await scan('msdyn_workorder')) || (await scan('bookableresourcebooking')) || null;
  return summaryField;
}

// Fetch the current Booking Summary value and drop it into its box.
async function loadBookingSummary() {
  const f = await resolveSummaryField();
  if (!f) return;
  try {
    const id = f.entity === 'bookableresourcebooking' ? bookingId : woId;
    if (!id) return;
    const rec = await xrmGet(f.entity, id, `?$select=${f.logical}`);
    if ($('f-booking-summary')) $('f-booking-summary').value = rec?.[f.logical] || '';
  } catch (_) {}
}

// Fetch the current Work Order Summary value (msdyn_workordersummary). Resilient:
// if the field no longer exists on this org, the box just stays empty.
async function loadWoSummary() {
  if (!woId) return;
  try {
    const rec = await xrmGet('msdyn_workorder', woId, '?$select=msdyn_workordersummary');
    if ($('f-wo-summary')) $('f-wo-summary').value = rec?.msdyn_workordersummary || '';
  } catch (_) {}
}
async function xrmCreate(entity, data) {
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{const __d=${JSON.stringify(data)};const r=await Xrm.WebApi.createRecord("${entity}",__d);return JSON.stringify({id:r.id});}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
  return r.id;
}
async function xrmDelete(entity, id) {
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{await Xrm.WebApi.deleteRecord("${entity}","${id}");return JSON.stringify({ok:1});}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
}

// ── State ─────────────────────────────────────────────────────────────────────
let booking = null, wo = null, woId = null, incident = null, contact = null, customerAsset = null, bookingStatuses = [], resources = [], subStatuses = [], workOrderTypes = [], dirty = {}, substatusNavProp = null, workordertypeNav = null;
let productsLoaded = false, notesLoaded = false, prodSearchInited = false;

// ── Boot ──────────────────────────────────────────────────────────────────────
async function loadData() {
  if (directWoId) {
    woId = directWoId;
    bookingStatuses = [];
    resources = [];
    $('booking-card').classList.add('hidden');
    $('wo-status-row').classList.add('hidden');
  } else {
    [booking, bookingStatuses, resources] = await Promise.all([
      xrmGet('bookableresourcebooking', bookingId, ''),
      xrmList('bookingstatus', '?$select=bookingstatusid,name,statuscode&$orderby=name asc'),
      xrmList('bookableresource', '?$select=bookableresourceid,name&$orderby=name asc'),
    ]);
    woId = booking._msdyn_workorder_value;
  }

  if (woId) {
    // The sub-status reference list doesn't depend on the work order, so start
    // it immediately and await it later (runs concurrently with everything else).
    const subStatusesP = xrmList('msdyn_workordersubstatus',
      '?$select=msdyn_workordersubstatusid,msdyn_name&$orderby=msdyn_name asc').catch(() => []);
    const woTypesP = xrmList('msdyn_workordertype',
      '?$select=msdyn_workordertypeid,msdyn_name&$orderby=msdyn_name asc').catch(() => []);

    wo = await xrmGet('msdyn_workorder', woId,
      '?$select=msdyn_name,msdyn_systemstatus,msdyn_instructions,' +
      'msdyn_address1,msdyn_address2,msdyn_city,msdyn_stateorprovince,msdyn_postalcode,msdyn_country,' +
      '_msdyn_serviceaccount_value,_msdyn_billingaccount_value,_msdyn_workordertype_value,' +
      '_msdyn_serviceterritory_value,_msdyn_substatus_value,_msdyn_priority_value,_msdyn_customerasset_value,' +
      'msdyn_datewindowstart,msdyn_datewindowend,msdyn_timetopromised,msdyn_timefrompromised,' +
      'wc_workorderproblemdescription,_msdyn_reportedbycontact_value');

    // Everything below only needs woId / values already on `wo`, so fire them all
    // off at once instead of one round-trip at a time.
    const contactId = wo._msdyn_reportedbycontact_value;
    const assetId = wo._msdyn_customerasset_value;
    const [contactR, assetR, incidents, subStatusesR, woTypesR] = await Promise.all([
      contactId
        ? xrmGet('contact', contactId, '?$select=fullname,telephone1,mobilephone,emailaddress1,jobtitle').catch(() => null)
        : Promise.resolve(null),
      assetId
        ? xrmGet('msdyn_customerasset', assetId, '?$select=msdyn_name,wc_assettag,msdyn_assettag,wc_seriallotnumber,_msdyn_product_value').catch(() => null)
        : Promise.resolve(null),
      xrmList('msdyn_workorderincident', `?$filter=_msdyn_workorder_value eq ${woId}&$top=1`).catch(() => []),
      subStatusesP,
      woTypesP,
      loadEngineers().catch(() => {}),
      loadAccountContract().catch(() => {}),
    ]);
    workOrderTypes = woTypesR || [];

    contact = contactR;
    customerAsset = assetR;
    if (customerAsset) {
      customerAsset.__tag    = customerAsset.wc_assettag || customerAsset.msdyn_assettag;
      customerAsset.__serial = customerAsset.wc_seriallotnumber;
    }
    incident = incidents[0] || null;
    subStatuses = subStatusesR || [];
  }
}

async function init() {
  try {
    await waitForXrm();
    await loadData();

    if (booking) { buildStatusDropdown(); buildResourceDropdown(); }
    if (wo) { buildSubstatusDropdown(); buildWorkOrderTypeDropdown(); }
    renderAll();
    listenEdits();
    wireOpenDynamics();
    wireRefresh();

    $('wo-loading').style.display = 'none';
    $('wo-content').style.display = 'flex';
  } catch(e) {
    $('wo-loading').innerHTML = `<div style="color:var(--danger);text-align:center;padding:20px;max-width:400px;">${esc(e.message)}</div>`;
  }
}

// Whether the work order's service account has an active service agreement (contract)
// covering today. Drives the parts-form "Warranty/Contract" default + banner.
let accountContract = null;
async function loadAccountContract() {
  accountContract = { active: false };
  const acctId = wo && (wo._msdyn_serviceaccount_value || wo._msdyn_billingaccount_value);
  if (!acctId) return;
  try {
    const ags = await xrmList('msdyn_agreement',
      `?$select=msdyn_name,msdyn_startdate,msdyn_enddate,statecode,c5_hasexclusions,c5_coveragetype,_c5_slaterms_value` +
      `&$filter=_msdyn_serviceaccount_value eq ${acctId} and statecode eq 0&$orderby=msdyn_enddate desc`);
    const now = new Date();
    const active = ags.filter(a => {
      const s = a.msdyn_startdate ? new Date(a.msdyn_startdate) : null;
      const e = a.msdyn_enddate   ? new Date(a.msdyn_enddate)   : null;
      return (!s || s <= now) && (!e || e >= now);
    });
    if (active.length) {
      const best = active[0]; // latest end date
      accountContract = {
        active: true,
        count: active.length,
        name: best.msdyn_name || '',
        coverage: best['_c5_slaterms_value@OData.Community.Display.V1.FormattedValue']
                  || best['c5_coveragetype@OData.Community.Display.V1.FormattedValue'] || '',
        endDate: best.msdyn_enddate || null,
        exclusions: !!best.c5_hasexclusions,
      };
    }
  } catch(_) { accountContract = { active: false }; }
}

// Default the Warranty/Contract field from the account's contract status and show a banner.
function applyContractToPartsForm() {
  const sel = $('product-warranty');
  if (!sel || !accountContract) return;
  sel.value = accountContract.active ? 'true' : 'false';

  let banner = $('contract-banner');
  if (!banner) {
    const grid = $('parts-form-card')?.querySelector('div[style*="grid-template-columns"]');
    if (!grid) return;
    banner = document.createElement('div');
    banner.id = 'contract-banner';
    banner.style.cssText = 'grid-column:1/4;font-size:12px;padding:8px 10px;border-radius:6px;';
    grid.insertBefore(banner, grid.firstChild);
  }
  if (accountContract.active) {
    banner.style.background = 'rgba(18,183,106,.15)'; banner.style.color = '#12b76a';
    const parts = [`✓ Active contract`];
    if (accountContract.coverage) parts.push(accountContract.coverage);
    if (accountContract.endDate)  parts.push(`exp ${new Date(accountContract.endDate).toLocaleDateString(undefined, { year:'numeric', month:'short', day:'numeric' })}`);
    banner.textContent = parts.join(' · ') + (accountContract.exclusions ? ' — has exclusions, verify parts coverage' : ' — parts covered');
  } else {
    banner.style.background = 'rgba(122,131,154,.15)'; banner.style.color = '#c4c9d2';
    banner.textContent = 'No active service contract on this account — parts likely billable';
  }
}

function wireRefresh() {
  $('refresh-btn')?.addEventListener('click', async () => {
    const btn = $('refresh-btn');
    btn.disabled = true;
    const original = btn.textContent;
    btn.textContent = 'Refreshing…';
    try {
      await loadData();
      dirty = {};
      if (booking) { buildStatusDropdown(); buildResourceDropdown(); }
      if (wo) { buildSubstatusDropdown(); buildWorkOrderTypeDropdown(); }
      renderAll();
      $('save-btn').classList.add('hidden');
      $('discard-btn').classList.add('hidden');
      toast('Refreshed');
    } catch(e) {
      toast('Refresh failed: ' + e.message, true);
    } finally {
      btn.disabled = false;
      btn.textContent = original;
    }
  });
}

function fv(obj, field) {
  if (!obj) return '';
  return obj[`${field}@OData.Community.Display.V1.FormattedValue`] || obj[field] || '';
}

function buildResourceDropdown() {
  const sel = $('f-resource');
  sel.innerHTML = resources.map(r =>
    `<option value="${r.bookableresourceid}">${esc(r.name)}</option>`
  ).join('');
  sel.value = booking._resource_value || '';
}

const ALLOWED_SUBSTATUSES = ['5 Day Monitoring', 'Completed', 'Unscheduled', 'Follow-up Required', 'Parts Required'];

function buildSubstatusDropdown() {
  const sel = $('f-substatus');
  const allowed = ALLOWED_SUBSTATUSES
    .map(name => subStatuses.find(s => s.msdyn_name?.toLowerCase() === name.toLowerCase()))
    .filter(Boolean);
  sel.innerHTML = '<option value="">—</option>' + allowed.map(s =>
    `<option value="${s.msdyn_workordersubstatusid}">${esc(s.msdyn_name)}</option>`
  ).join('');
  sel.value = wo._msdyn_substatus_value || '';
}

function buildWorkOrderTypeDropdown() {
  const sel = $('f-workordertype');
  if (!sel) return;
  const current = wo._msdyn_workordertype_value || '';
  const currentName = wo['_msdyn_workordertype_value@OData.Community.Display.V1.FormattedValue'] || '';
  const opts = workOrderTypes.slice();
  // Make sure the work order's current type is always an option, even if the list
  // hasn't loaded (so the field never appears blank for an already-typed WO).
  if (current && !opts.some(t => t.msdyn_workordertypeid === current)) {
    opts.unshift({ msdyn_workordertypeid: current, msdyn_name: currentName || '(current)' });
  }
  sel.innerHTML = '<option value="">—</option>' + opts.map(t =>
    `<option value="${t.msdyn_workordertypeid}">${esc(t.msdyn_name)}</option>`
  ).join('');
  sel.value = current;
}

function wireOpenDynamics() {
  const APP_ID = '5f751dd8-1b58-eb11-bb23-000d3a3b3842';
  const url = woId
    ? `${orgUrl}/main.aspx?appid=${APP_ID}&pagetype=entityrecord&etn=msdyn_workorder&id=${woId}`
    : `${orgUrl}/main.aspx?appid=${APP_ID}&pagetype=entityrecord&etn=bookableresourcebooking&id=${bookingId}`;
  $('open-dynamics-btn')?.addEventListener('click', () => window.api.openExternal(url));
}

function buildStatusDropdown() {
  const sel = $('f-booking-status');
  const seen = new Map();
  for (const s of bookingStatuses) {
    if (!seen.has(s.name) || s.statuscode < seen.get(s.name).statuscode) seen.set(s.name, s);
  }
  const deduped = [...seen.values()].sort((a,b) => a.name.localeCompare(b.name));
  sel.innerHTML = deduped.map(s => `<option value="${s.bookingstatusid}">${esc(s.name)}</option>`).join('');
  const current = bookingStatuses.find(s => s.bookingstatusid === booking._bookingstatus_value);
  sel.value = (current ? seen.get(current.name)?.bookingstatusid : null) || booking._bookingstatus_value || '';
}

function openAccount(accountId, name) {
  if (!accountId) return;
  openAccountDetail(accountId);
}

// Account detail modal shown inside this window — mirrors the Accounts-tab modal in
// the main window (openAccountDetail in app.js).
async function openAccountDetail(accountId) {
  const modal = $('account-modal'), body = $('account-modal-body');
  if (!modal || !body) return;
  body.innerHTML = '<div class="inline-loading"><div class="spinner"></div></div>';
  modal.classList.remove('hidden');
  try {
    const [accounts, contacts, workOrders, assets] = await Promise.all([
      xrmList('account', `?$select=name,emailaddress1,telephone1,address1_line1,address1_city,address1_stateorprovince,address1_postalcode,websiteurl,description&$filter=accountid eq ${accountId}`),
      xrmList('contact', `?$select=contactid,fullname,emailaddress1,mobilephone,telephone1,jobtitle&$filter=_parentcustomerid_value eq ${accountId}&$orderby=fullname asc`),
      xrmList('msdyn_workorder', `?$select=msdyn_workorderid,msdyn_name,msdyn_systemstatus,createdon&$filter=_msdyn_serviceaccount_value eq ${accountId}&$orderby=createdon desc&$top=10`).catch(() => []),
      xrmList('msdyn_customerasset', `?$select=msdyn_customerassetid,msdyn_name,wc_assettag,msdyn_assettag,wc_knumber,wc_seriallotnumber,statuscode&$filter=_msdyn_account_value eq ${accountId}&$orderby=msdyn_name asc&$top=200`).catch(() => []),
    ]);
    const a = accounts[0];
    if (!a) { body.innerHTML = '<div class="am-empty">Account not found.</div>'; return; }
    const address = [a.address1_line1, a.address1_city, a.address1_stateorprovince, a.address1_postalcode].filter(Boolean).join(', ');
    body.innerHTML = `
      <div class="am-title">${esc(a.name || '—')}</div>
      <div class="am-sub">${esc(address || 'No address on file')}</div>
      <div class="am-grid">
        <div><div class="am-field-label">Phone</div><div class="am-field-value">${esc(a.telephone1||'—')}</div></div>
        <div><div class="am-field-label">Email</div><div class="am-field-value">${esc(a.emailaddress1||'—')}</div></div>
        <div><div class="am-field-label">Website</div><div class="am-field-value">${esc(a.websiteurl||'—')}</div></div>
      </div>
      <div class="am-section am-collapsed">
        <div class="am-section-title am-toggle">Contacts (${contacts.length})<span class="am-caret">&#9656;</span></div>
        <div class="am-list">
          ${contacts.length ? contacts.map(c => `
            <div class="am-list-item am-clickable" data-contact-id="${esc(c.contactid)}">
              <div class="am-list-item-title">${esc(c.fullname||'—')}</div>
              <div class="am-list-item-sub">${esc(c.jobtitle || '')}${c.jobtitle && (c.mobilephone||c.telephone1||c.emailaddress1) ? ' · ' : ''}${esc(c.mobilephone||c.telephone1||'')}${(c.mobilephone||c.telephone1) && c.emailaddress1 ? ' · ' : ''}${esc(c.emailaddress1||'')}</div>
            </div>`).join('') : '<div class="am-empty">No contacts on file.</div>'}
        </div>
      </div>
      <div class="am-section am-collapsed">
        <div class="am-section-title am-toggle">Systems (${assets.length})<span class="am-caret">&#9656;</span></div>
        <div class="am-list">
          ${assets.length ? assets.map(s => {
            const tag = s.wc_assettag || s.msdyn_assettag || '';
            const sub = [tag ? 'Tag ' + tag : '', s.wc_knumber ? 'K ' + s.wc_knumber : '', s.wc_seriallotnumber ? 'S/N ' + s.wc_seriallotnumber : '', fv(s,'statuscode')].filter(Boolean).join(' · ');
            return `<div class="am-list-item am-clickable" data-asset-id="${esc(s.msdyn_customerassetid)}">
              <div class="am-list-item-title">${esc(s.msdyn_name||'—')}</div>
              <div class="am-list-item-sub">${esc(sub || '—')}</div>
            </div>`;
          }).join('') : '<div class="am-empty">No systems on file.</div>'}
        </div>
      </div>
      <div class="am-section am-collapsed">
        <div class="am-section-title am-toggle">Recent Work Orders (${workOrders.length})<span class="am-caret">&#9656;</span></div>
        <div class="am-list">
          ${workOrders.length ? workOrders.map(w => `
            <div class="am-list-item am-clickable" data-wo-id="${esc(w.msdyn_workorderid)}">
              <div class="am-list-item-title">${esc(w.msdyn_name||'—')}</div>
              <div class="am-list-item-sub">${esc(w['msdyn_systemstatus@OData.Community.Display.V1.FormattedValue']||'')}${fmtDate(w.createdon) !== '—' ? ' · ' + fmtDate(w.createdon) : ''}</div>
            </div>`).join('') : '<div class="am-empty">No work orders on file.</div>'}
        </div>
      </div>`;
    body.querySelectorAll('.am-toggle').forEach(t =>
      t.addEventListener('click', () => t.closest('.am-section').classList.toggle('am-collapsed')));
    body.querySelectorAll('[data-contact-id]').forEach(el =>
      el.addEventListener('click', () => window.api.openContact(el.dataset.contactId, orgUrl, 'Contact')));
    body.querySelectorAll('[data-asset-id]').forEach(el =>
      el.addEventListener('click', () => openAssetDetail(el.dataset.assetId)));
    body.querySelectorAll('[data-wo-id]').forEach(el =>
      el.addEventListener('click', () => window.api.openWorkOrderDirect(el.dataset.woId, orgUrl, 'Work Order')));
  } catch (e) {
    body.innerHTML = '<div class="am-empty">Failed to load account details.</div>';
    console.error('Account detail error:', e);
  }
}
$('account-modal-close')?.addEventListener('click', () => $('account-modal').classList.add('hidden'));
$('account-modal')?.addEventListener('click', e => { if (e.target.id === 'account-modal') $('account-modal').classList.add('hidden'); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') $('account-modal')?.classList.add('hidden'); });

// Customer asset detail modal — mirrors the Assets-tab modal in the main window.
const ASSETS_DETAIL_SELECT = 'msdyn_customerassetid,msdyn_name,wc_assettag,msdyn_assettag,wc_seriallotnumber,' +
  'statuscode,wc_knumber,msdyn_manufacturingdate,_msdyn_parentasset_value,_wc_warrantyservicecontract_value,' +
  '_msdyn_masterasset_value,_msdyn_product_value,_wc_manufacturer_value,_msdyn_workorderproduct_value,' +
  '_msdyn_account_value,_msdyn_functionallocation_value';

async function openAssetDetail(assetId) {
  const modal = $('asset-modal'), body = $('asset-modal-body');
  if (!modal || !body || !assetId) return;
  body.innerHTML = '<div class="inline-loading"><div class="spinner"></div></div>';
  modal.classList.remove('hidden');
  try {
    const records = await xrmList('msdyn_customerasset', `?$select=${ASSETS_DETAIL_SELECT}&$filter=msdyn_customerassetid eq ${assetId}`);
    const a = records[0];
    if (!a) { body.innerHTML = '<div class="am-empty">Asset not found.</div>'; return; }
    const tag = a.wc_assettag || a.msdyn_assettag || '—';
    body.innerHTML = `
      <div class="am-title">${esc(a.msdyn_name || '—')}</div>
      <div class="am-sub">Asset Tag: ${esc(tag)}</div>
      <div class="am-grid">
        <div><div class="am-field-label">Site</div><div class="am-field-value">${esc(fv(a,'_msdyn_functionallocation_value') || fv(a,'_msdyn_account_value') || '—')}</div></div>
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
      </div>`;
  } catch (e) {
    body.innerHTML = '<div class="am-empty">Failed to load asset details.</div>';
    console.error('Asset detail error:', e);
  }
}
$('asset-modal-close')?.addEventListener('click', () => $('asset-modal').classList.add('hidden'));
$('asset-modal')?.addEventListener('click', e => { if (e.target.id === 'asset-modal') $('asset-modal').classList.add('hidden'); });
document.addEventListener('keydown', e => { if (e.key === 'Escape') $('asset-modal')?.classList.add('hidden'); });
// Set an element's text and, when an account id is present, make it a clickable link
// that opens the account detail modal.
function setAccountLink(id, accountId, name) {
  const el = $(id); if (!el) return;
  el.textContent = name || '—';
  if (accountId) {
    el.style.cursor = 'pointer';
    el.style.color = 'var(--accent)';
    el.title = 'Open account';
    el.onclick = () => openAccount(accountId, name);
  } else {
    el.style.cursor = ''; el.style.color = ''; el.title = ''; el.onclick = null;
  }
}
// Same, for a customer asset -> opens the asset detail modal.
function setAssetLink(id, assetId, name) {
  const el = $(id); if (!el) return;
  el.textContent = name || '—';
  if (assetId) {
    el.style.cursor = 'pointer';
    el.style.color = 'var(--accent)';
    el.title = 'Open asset';
    el.onclick = () => openAssetDetail(assetId);
  } else {
    el.style.cursor = ''; el.style.color = ''; el.title = ''; el.onclick = null;
  }
}

function renderAll() {
  const woNum      = wo?.msdyn_name || '';
  const account    = wo ? fv(wo,'_msdyn_serviceaccount_value') : '';

  $('titlebar-label').textContent = woNum || booking?.name || 'Work Order';
  $('wo-number').textContent      = woNum || '—';
  setAccountLink('wo-account', wo?._msdyn_serviceaccount_value, account);
  $('wo-booking-ref').textContent = booking?.name ? `Booking: ${booking.name}` : '';
  document.title = woNum || booking?.name || 'Work Order';

  if (booking) {
    // Booking fields
    $('f-start').value   = isoToLocal(booking.starttime);
    $('f-end').value     = isoToLocal(booking.endtime);
    $('f-arrival').value = isoToLocal(booking.msdyn_actualarrivaltime || '');
    set('d-duration',  fmtDuration(booking.duration));
  }

  // WO fields
  const WO_STATUS = {690970000:'Unscheduled',690970001:'Scheduled',690970002:'In Progress',690970003:'Completed',690970004:'Posted',690970005:'Canceled'};
  set('d-wo-status', wo ? (WO_STATUS[wo.msdyn_systemstatus] || fv(wo,'msdyn_systemstatus')) : '—');
  set('d-priority',   wo ? fv(wo,'_msdyn_priority_value')        : '—');
  setAccountLink('d-account', wo?._msdyn_serviceaccount_value, wo ? fv(wo,'_msdyn_serviceaccount_value') : '—');
  set('d-contact-name',  contact?.fullname      || '—');
  set('d-contact-phone', contact?.telephone1 || contact?.mobilephone || '—');
  set('d-contact-email', contact?.emailaddress1 || '—');
  set('d-contact-title', contact?.jobtitle      || '');
  setAccountLink('d-billing', wo?._msdyn_billingaccount_value, wo ? fv(wo,'_msdyn_billingaccount_value') : '—');
  set('d-territory',  wo ? fv(wo,'_msdyn_serviceterritory_value'): '—');
  setAssetLink('d-asset-tag', wo?._msdyn_customerasset_value, wo ? fv(wo,'_msdyn_customerasset_value') : '—');
  set('d-asset-tagnum',  customerAsset?.__tag    || '—');
  set('d-asset-serial',  customerAsset?.__serial || '—');
  set('d-win-start',  fmtDate(wo?.msdyn_datewindowstart));
  set('d-win-end',    fmtDate(wo?.msdyn_datewindowend));
  set('d-time-from',  fmtDate(wo?.msdyn_timefrompromised));
  set('d-time-to',    fmtDate(wo?.msdyn_timetopromised));

  // Editable WO fields
  $('f-addr1').value        = wo?.msdyn_address1         || '';
  $('f-addr2').value        = wo?.msdyn_address2         || '';
  $('f-city').value         = wo?.msdyn_city             || '';
  $('f-state').value        = wo?.msdyn_stateorprovince  || '';
  $('f-zip').value          = wo?.msdyn_postalcode       || '';
  $('f-country').value      = wo?.msdyn_country          || '';
  $('f-booking-summary').value = '';  // filled async from the Booking Summary field
  $('f-wo-summary').value      = '';  // filled async from msdyn_workordersummary
  loadBookingSummary();
  loadWoSummary();
  $('f-instructions').value = wo?.msdyn_instructions     || '';
  $('f-problem').value = wo?.wc_workorderproblemdescription || '';
}

function set(id, val) { const el=$(id); if(el) el.textContent = val||'—'; }

// ── Edit listeners ────────────────────────────────────────────────────────────
function listenEdits() {
  $('f-booking-status').addEventListener('change', e => {
    dirty._bookingStatus = e.target.value;
    // Auto-fill arrival time when switching to In Progress and it's not already set
    const selected = bookingStatuses.find(s => s.bookingstatusid === e.target.value);
    if (selected && selected.name.toLowerCase().includes('progress') && !$('f-arrival').value) {
      const now = isoToLocal(new Date().toISOString());
      $('f-arrival').value = now;
      dirty._actualarrival = now;
    }
    showSave();
  });
  $('f-substatus').addEventListener('change', e => { dirty._substatus = e.target.value; showSave(); });
  $('f-workordertype')?.addEventListener('change', e => { dirty._workordertype = e.target.value; showSave(); });
  $('f-start').addEventListener('input',   e => { dirty._starttime = e.target.value; showSave(); });
  $('f-end').addEventListener('input',     e => { dirty._endtime   = e.target.value; showSave(); });
  $('f-arrival').addEventListener('input', e => { dirty._actualarrival = e.target.value; showSave(); });
  $('f-resource').addEventListener('change', e => { dirty._resource = e.target.value; showSave(); });

  [['f-addr1','msdyn_address1'],['f-addr2','msdyn_address2'],['f-city','msdyn_city'],
   ['f-state','msdyn_stateorprovince'],['f-zip','msdyn_postalcode'],['f-country','msdyn_country'],
   ['f-booking-summary','_bookingSummary'],['f-wo-summary','msdyn_workordersummary'],
   ['f-instructions','msdyn_instructions'],
   ['f-problem','wc_workorderproblemdescription']
  ].forEach(([id,key]) => $(id)?.addEventListener('input', () => { dirty[key]=$(id).value; showSave(); }));


  $('save-btn').addEventListener('click',    save);
  $('discard-btn').addEventListener('click', discard);
}

function showSave() { $('save-btn').classList.remove('hidden'); $('discard-btn').classList.remove('hidden'); }

async function save() {
  const btn = $('save-btn');
  btn.textContent = 'Saving…'; btn.disabled = true;
  const snap = {...dirty};
  try {
    // Resolve the Booking Summary field first. Dynamics has a rule that blocks
    // completing a booking unless its notes are filled, so the summary must be
    // written BEFORE (or together with) the booking-status change — never after.
    let summaryF = null;
    if (snap._bookingSummary !== undefined) {
      summaryF = await resolveSummaryField();
      if (!summaryF) throw new Error('Could not find the Booking Summary field in Dynamics');
    }

    const bPatch = {};
    if (snap._bookingStatus) bPatch['BookingStatus@odata.bind'] = `/bookingstatuses(${snap._bookingStatus})`;
    if (snap._starttime)      bPatch.starttime          = new Date(snap._starttime).toISOString();
    if (snap._endtime)        bPatch.endtime            = new Date(snap._endtime).toISOString();
    if (snap._actualarrival)  bPatch.msdyn_actualarrivaltime = new Date(snap._actualarrival).toISOString();
    if (snap._resource)       bPatch['Resource@odata.bind'] = `/bookableresources(${snap._resource})`;
    // If Booking Summary lives on the booking, set it in the SAME update as the
    // status so the notes are present when the completion rule evaluates.
    if (summaryF && summaryF.entity === 'bookableresourcebooking') {
      bPatch[summaryF.logical] = snap._bookingSummary;
    } else if (summaryF) {
      // Otherwise it's on the work order — write it before touching the booking.
      if (woId) await xrmUpdate('msdyn_workorder', woId, { [summaryF.logical]: snap._bookingSummary });
    }
    if (Object.keys(bPatch).length) await xrmUpdate('bookableresourcebooking', bookingId, bPatch);

    const wPatch = {};
    ['msdyn_address1','msdyn_address2','msdyn_city','msdyn_stateorprovince',
     'msdyn_postalcode','msdyn_country','msdyn_instructions','msdyn_workordersummary','wc_workorderproblemdescription'
    ].forEach(k => { if (snap[k] !== undefined) wPatch[k] = snap[k]; });
    if (woId && Object.keys(wPatch).length) await xrmUpdate('msdyn_workorder', woId, wPatch);

    if (woId && snap._substatus) {
      if (!substatusNavProp) substatusNavProp = await getLookupNavProperty('msdyn_workorder', 'msdyn_substatus');
      await xrmUpdate('msdyn_workorder', woId, { [`${substatusNavProp}@odata.bind`]: `/msdyn_workordersubstatuses(${snap._substatus})` });
    }

    if (woId && snap._workordertype) {
      if (!workordertypeNav) workordertypeNav = await getLookupNavProperty('msdyn_workorder', 'msdyn_workordertype');
      await xrmUpdate('msdyn_workorder', woId, { [`${workordertypeNav}@odata.bind`]: `/msdyn_workordertypes(${snap._workordertype})` });
    }


    // Sync local state
    if (snap._bookingStatus) {
      booking._bookingstatus_value = snap._bookingStatus;
      const found = bookingStatuses.find(s => s.bookingstatusid === snap._bookingStatus);
      if (found && booking.BookingStatus) booking.BookingStatus.name = found.name;
    }
    if (snap._starttime)     booking.starttime           = new Date(snap._starttime).toISOString();
    if (snap._endtime)       booking.endtime             = new Date(snap._endtime).toISOString();
    if (snap._actualarrival) booking.msdyn_actualarrivaltime = new Date(snap._actualarrival).toISOString();
    if (snap._resource)     booking._resource_value = snap._resource;
    if (wo) Object.assign(wo, wPatch);
    if (wo && snap._substatus) wo._msdyn_substatus_value = snap._substatus;
    if (wo && snap._workordertype) {
      wo._msdyn_workordertype_value = snap._workordertype;
      const t = workOrderTypes.find(t => t.msdyn_workordertypeid === snap._workordertype);
      if (t) wo['_msdyn_workordertype_value@OData.Community.Display.V1.FormattedValue'] = t.msdyn_name;
    }

    dirty = {};
    $('save-btn').classList.add('hidden');
    $('discard-btn').classList.add('hidden');
    toast('Saved');
  } catch(e) { toast('Save failed: '+e.message, true); }
  finally { btn.textContent = 'Save'; btn.disabled = false; }
}

function discard() {
  dirty = {};
  if (booking) { buildStatusDropdown(); buildResourceDropdown(); }
  if (wo) { buildSubstatusDropdown(); buildWorkOrderTypeDropdown(); }
  renderAll();
  $('save-btn').classList.add('hidden');
  $('discard-btn').classList.add('hidden');
}

// ── Assigned Engineers ────────────────────────────────────────────────────────
function resourceOptions(selectedId) {
  return resources.map(r =>
    `<option value="${r.bookableresourceid}"${r.bookableresourceid === selectedId ? ' selected' : ''}>${esc(r.name)}</option>`
  ).join('');
}

async function loadEngineers() {
  const el = $('d-engineers-list');
  if (!el) return;
  if (!resources.length) {
    try { resources = await xrmList('bookableresource', '?$select=bookableresourceid,name&$orderby=name asc'); } catch (_) {}
  }
  const bookings = await xrmList('bookableresourcebooking',
    `?$select=bookableresourcebookingid,name,starttime,endtime,_resource_value&$expand=Resource($select=name)&$filter=_msdyn_workorder_value eq ${woId}&$orderby=starttime asc`);

  const rowsHtml = bookings.length ? bookings.map((b, i) => {
    const name = b.Resource?.name || 'Unassigned';
    const time = `${fmtDate(b.starttime)} – ${fmtDate(b.endtime)}`;
    return `
      <div class="engineer-row" data-idx="${i}" data-bid="${b.bookableresourcebookingid}" data-res="${b._resource_value || ''}"
        style="background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:8px 12px;display:flex;justify-content:space-between;align-items:center;gap:8px;">
        <span class="eng-open" style="font-size:13px;cursor:pointer;flex:1;">${esc(name)}</span>
        <span style="font-size:11px;color:var(--muted2);">${esc(time)}</span>
        <button class="btn btn-ghost btn-sm eng-transfer" title="Transfer to another engineer" style="padding:2px 8px;font-size:11px;">Transfer</button>
        <button class="btn btn-ghost btn-sm eng-remove" title="Remove engineer (deletes this booking)" data-name="${esc(name)}" style="padding:2px 8px;font-size:11px;color:var(--danger);">Remove</button>
      </div>`;
  }).join('') : `<div class="field-value dim">No engineers assigned</div>`;

  el.innerHTML = rowsHtml + `
    <div id="add-engineer-bar" style="display:flex;gap:8px;align-items:center;margin-top:4px;">
      <select id="add-engineer-select" class="field-input" style="flex:1;"><option value="">Add another engineer…</option>${resourceOptions('')}</select>
      <button class="btn btn-ghost btn-sm" id="add-engineer-btn" disabled>Add</button>
    </div>`;

  // Open a booking in its own window
  el.querySelectorAll('.engineer-row .eng-open').forEach(span => {
    span.addEventListener('click', () => {
      const b = bookings[+span.closest('.engineer-row').dataset.idx];
      window.api.openWorkOrder(b.bookableresourcebookingid, orgUrl, b.Resource?.name || b.name || 'Booking');
    });
  });

  // Inline transfer: swap the row for a resource picker
  el.querySelectorAll('.eng-transfer').forEach(btn => {
    btn.addEventListener('click', () => {
      const row = btn.closest('.engineer-row');
      const bid = row.dataset.bid, curRes = row.dataset.res;
      row.innerHTML = `
        <select class="field-input eng-transfer-sel" style="flex:1;">${resourceOptions(curRes)}</select>
        <button class="btn btn-primary btn-sm eng-transfer-confirm" style="padding:2px 8px;font-size:11px;">Save</button>
        <button class="btn btn-ghost btn-sm eng-transfer-cancel" style="padding:2px 8px;font-size:11px;">Cancel</button>`;
      row.querySelector('.eng-transfer-cancel').addEventListener('click', loadEngineers);
      row.querySelector('.eng-transfer-confirm').addEventListener('click', async () => {
        const newRes = row.querySelector('.eng-transfer-sel').value;
        if (!newRes || newRes === curRes) { loadEngineers(); return; }
        try {
          await xrmUpdate('bookableresourcebooking', bid, { 'Resource@odata.bind': `/bookableresources(${newRes})` });
          toast('Engineer transferred');
          await loadEngineers();
        } catch (e) { toast('Transfer failed: ' + e.message, true); loadEngineers(); }
      });
    });
  });

  // Remove engineer = delete their booking off this work order
  el.querySelectorAll('.eng-remove').forEach(btn => {
    btn.addEventListener('click', async () => {
      const row = btn.closest('.engineer-row');
      const bid = row.dataset.bid, name = btn.dataset.name || 'this engineer';
      if (!bid) return;
      if (!confirm(`Remove ${name} from this work order? This deletes their booking.`)) return;
      btn.disabled = true; btn.textContent = 'Removing…';
      try {
        await xrmDelete('bookableresourcebooking', bid);
        toast('Engineer removed');
        await loadEngineers();
      } catch (e) {
        toast('Remove failed: ' + e.message, true);
        btn.disabled = false; btn.textContent = 'Remove';
      }
    });
  });

  // Add another engineer = duplicate this booking onto another resource (same work order)
  const addSel = $('add-engineer-select'), addBtn = $('add-engineer-btn');
  addSel.addEventListener('change', () => { addBtn.disabled = !addSel.value; });
  addBtn.addEventListener('click', async () => {
    const resId = addSel.value;
    if (!resId || !woId) return;
    addBtn.disabled = true; addBtn.textContent = 'Adding…';
    try {
      const payload = {
        'msdyn_workorder@odata.bind': `/msdyn_workorders(${woId})`,
        'Resource@odata.bind': `/bookableresources(${resId})`,
      };
      if (booking?.starttime) payload.starttime = booking.starttime;
      if (booking?.endtime)   payload.endtime   = booking.endtime;
      if (booking?.duration != null) payload.duration = booking.duration;
      if (booking?._bookingstatus_value) payload['BookingStatus@odata.bind'] = `/bookingstatuses(${booking._bookingstatus_value})`;
      await xrmCreate('bookableresourcebooking', payload);
      toast('Engineer added to work order');
      await loadEngineers();
    } catch (e) {
      toast('Add failed: ' + e.message, true);
      addBtn.disabled = false; addBtn.textContent = 'Add';
    }
  });
}

// ── Tabs ──────────────────────────────────────────────────────────────────────
document.querySelectorAll('.wo-tab').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.wo-tab').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.wo-panel').forEach(p => p.classList.add('hidden'));
    btn.classList.add('active');
    $(`panel-${btn.dataset.tab}`).classList.remove('hidden');
    if (btn.dataset.tab==='products') { initProdSearch(); loadProducts(); }
    if (btn.dataset.tab==='timeline') loadTimeline();
    if (btn.dataset.tab==='notes'    && !notesLoaded)    loadNotes();
    if (btn.dataset.tab==='details') { /* always loaded */ }
  });
});

// ── Parts Request (Work Order Products) ──────────────────────────────────────
// Parts are ordered exactly like the Dynamics UI: each part is a msdyn_workorderproduct
// record on the work order, then flagged submitted (the "Submit Parts" ribbon action).
let partsRequestRows = [];
let draftParts = [];
let partsOptionsLoaded = false;
const partsOptions = { shipping: [], shiptolocation: [], systemstatus: [] };
const partsOptionLabel = { shipping: {}, shiptolocation: {}, systemstatus: {} };

async function fetchWopOptionSet(attr) {
  const url = `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='msdyn_workorderproduct')/Attributes(LogicalName='${attr}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet`;
  const json = await apiWv.executeJavaScript(`fetch(${JSON.stringify(url)},{headers:{Accept:'application/json'}}).then(r=>r.json()).then(d=>JSON.stringify(d)).catch(e=>JSON.stringify({__err:e.message}))`);
  const r = JSON.parse(json);
  return (r.OptionSet?.Options || []).map(o => ({ value: o.Value, label: o.Label?.UserLocalizedLabel?.Label || String(o.Value) }));
}
function fillOptionSelect(id, opts) {
  const sel = $(id); if (!sel) return;
  sel.innerHTML = '<option value="">—</option>' + opts.map(o => `<option value="${o.value}">${esc(o.label)}</option>`).join('');
}
// Populate the Shipping / Ship-to / System-status dropdowns from Dynamics metadata
// so the option values are always correct for this org.
async function loadPartsOptions() {
  if (partsOptionsLoaded) return;
  partsOptionsLoaded = true;
  const map = { shipping:'cr217_shipping', shiptolocation:'cr217_shiptolocation', systemstatus:'cr217_currentsystemstatus' };
  await Promise.all(Object.entries(map).map(async ([key, attr]) => {
    try {
      const opts = await fetchWopOptionSet(attr);
      partsOptions[key] = opts;
      opts.forEach(o => { partsOptionLabel[key][o.value] = o.label; });
    } catch(_) {}
  }));
  fillOptionSelect('product-shipping', partsOptions.shipping);
  fillOptionSelect('product-shiptolocation', partsOptions.shiptolocation);
  fillOptionSelect('product-systemstatus', partsOptions.systemstatus);
}

async function loadProducts() {
  if (!woId) { $('products-body').innerHTML=`<tr><td colspan="5"><div class="empty-msg">No work order linked</div></td></tr>`; return; }
  try {
    // Pull parts requests (linked either by regardingobjectid or the cr217_workorder lookup)
    // AND work order products, so anything ordered on this work order shows up regardless of
    // how it was created. Separate queries avoid a fragile OR on the polymorphic regarding field.
    const prSelect = '?$select=activityid,subject,wc_partnumber,wc_quantity,_wc_partname_value,wc_shiptolocation,cr217_partsrequeststatus,createdon';
    const [byReg, byWo, prods] = await Promise.all([
      xrmList('wc_partsrequest', `${prSelect}&$filter=_regardingobjectid_value eq ${woId}&$orderby=createdon desc`).catch(() => []),
      xrmList('wc_partsrequest', `${prSelect}&$filter=_cr217_workorder_value eq ${woId}&$orderby=createdon desc`).catch(() => []),
      xrmList('msdyn_workorderproduct',
        `?$select=msdyn_name,_msdyn_product_value,msdyn_quantity,msdyn_linestatus,cr217_newpartnumbernotinsystem,cr217_vendor,cr217_shiptolocation,pmich_new_partrequestsubmitted,createdon&$filter=_msdyn_workorder_value eq ${woId}&$orderby=createdon desc`).catch(() => []),
    ]);
    const seen = new Set();
    const requests = [];
    for (const r of [...byReg, ...byWo]) { if (!seen.has(r.activityid)) { seen.add(r.activityid); requests.push(r); } }
    const products = prods;
    partsRequestRows = [
      ...requests.map(p => ({
        part: p['_wc_partname_value@OData.Community.Display.V1.FormattedValue'] || p.wc_partnumber || p.subject || '—',
        qty: p.wc_quantity ?? '—',
        shipTo: p['wc_shiptolocation@OData.Community.Display.V1.FormattedValue'] || '—',
        status: p['cr217_partsrequeststatus@OData.Community.Display.V1.FormattedValue'] || 'Open',
        when: p.createdon,
      })),
      ...products.map(p => ({
        part: p.cr217_newpartnumbernotinsystem || p['_msdyn_product_value@OData.Community.Display.V1.FormattedValue'] || p.msdyn_name || '—',
        qty: p.msdyn_quantity ?? '—',
        shipTo: p['cr217_shiptolocation@OData.Community.Display.V1.FormattedValue'] || '—',
        status: p.pmich_new_partrequestsubmitted ? 'Submitted' : (p['msdyn_linestatus@OData.Community.Display.V1.FormattedValue'] || 'Open'),
        when: p.createdon,
      })),
    ].sort((a, b) => new Date(b.when) - new Date(a.when));
    productsLoaded = true;
    renderPartsTable();
  } catch(e) { $('products-body').innerHTML=`<tr><td colspan="5"><div class="empty-msg">Error: ${esc(e.message)}</div></td></tr>`; }
}

function renderPartsTable() {
  const rows = partsRequestRows.map(p => `<tr>
      <td>${esc(p.part)}</td><td class="col-muted">${esc(p.qty)}</td><td class="col-muted">${esc(p.shipTo)}</td>
      <td><span class="status-badge badge-scheduled" style="font-size:10px;padding:2px 8px;">${esc(p.status)}</span></td><td></td>
    </tr>`);
  const draftRows = draftParts.map((d, i) => `<tr>
    <td>${esc(d.displayName)}</td><td class="col-muted">${esc(d.quantity)}</td><td class="col-muted">${esc(partsOptionLabel.shiptolocation[d.shipToLocation]||'—')}</td>
    <td><span class="status-badge badge-cancelled" style="font-size:10px;padding:2px 8px;">Draft</span></td>
    <td><button class="btn btn-ghost btn-sm" data-remove-draft="${i}" style="padding:2px 8px;font-size:11px;">Remove</button></td>
  </tr>`);
  const all = [...draftRows, ...rows];
  $('products-body').innerHTML = all.length ? all.join('') : `<tr><td colspan="5"><div class="empty-msg">No parts requests</div></td></tr>`;
  $('products-body').querySelectorAll('[data-remove-draft]').forEach(btn => {
    btn.addEventListener('click', () => { draftParts.splice(+btn.dataset.removeDraft,1); renderPartsTable(); });
  });
  $('submit-parts-request-btn').disabled = !draftParts.length;
}

let selProduct=null, searchTimer=null;
function initProdSearch() {
  if (prodSearchInited) return; prodSearchInited=true;

  // Collapsible "new parts request" form
  const formCard = $('parts-form-card'), toggleBtn = $('toggle-parts-form');
  toggleBtn.addEventListener('click', () => {
    const hidden = formCard.classList.toggle('hidden');
    toggleBtn.textContent = hidden ? '+ New Parts Request' : '– Hide Form';
    if (!hidden) { loadPartsOptions(); applyContractToPartsForm(); $('product-search').focus(); }
  });
  loadPartsOptions();
  applyContractToPartsForm();

  const sEl=$('product-search'), rEl=$('product-results'), selEl=$('product-selected');
  sEl.addEventListener('input', () => {
    clearTimeout(searchTimer); selProduct=null; selEl.style.display='none';
    const q=sEl.value.trim();
    if (q.length<2) { rEl.style.display='none'; rEl.innerHTML=''; return; }
    searchTimer = setTimeout(async () => {
      try {
        const esc_q = q.replace(/'/g,"''");
        const res = await xrmList('product', `?$select=productid,name,productnumber&$filter=(contains(name,'${esc_q}') or contains(productnumber,'${esc_q}')) and statecode eq 0&$top=10&$orderby=name asc`);
        if (!res.length) { rEl.innerHTML='<div class="prod-option col-muted">No results</div>'; rEl.style.display='block'; return; }
        rEl.innerHTML = res.map(p=>`<div class="prod-option" data-id="${p.productid}" data-name="${esc(p.name)}" data-num="${esc(p.productnumber||'')}">${esc(p.name)}${p.productnumber?` <span class="col-muted">(${esc(p.productnumber)})</span>`:''}</div>`).join('');
        rEl.style.display='block';
        rEl.querySelectorAll('.prod-option').forEach(el => el.addEventListener('click', () => {
          selProduct={id:el.dataset.id,name:el.dataset.name,number:el.dataset.num};
          sEl.value=el.dataset.name; rEl.style.display='none'; rEl.innerHTML='';
          selEl.textContent=`✓ ${el.dataset.name}`; selEl.style.display='block';
        }));
      } catch(_) { rEl.innerHTML='<div class="prod-option col-muted">Search failed</div>'; rEl.style.display='block'; }
    }, 350);
  });

  $('add-product-btn').addEventListener('click', () => {
    const partNumber = $('product-partnumber').value.trim();
    const qty = parseFloat($('product-qty').value) || 1;
    if (!selProduct && !partNumber) { toast('Select a product or enter a part number', true); return; }
    // Dynamics requires a Product on every Work Order Product. If the part isn't in
    // the catalog (New Part # only), we fall back to the work order asset's product
    // at submit time — but if there's no asset product to borrow, we can't proceed.
    if (!selProduct && !(customerAsset && customerAsset._msdyn_product_value)) {
      toast('This part isn’t in the catalog and the work order has no asset product to attach it to — please select a product.', true);
      return;
    }
    const boolOf = v => v === '' ? null : (v === 'true');
    const intOf  = v => v ? parseInt(v, 10) : null;
    draftParts.push({
      product: selProduct ? { id: selProduct.id, name: selProduct.name } : null,
      partNumber: partNumber || (selProduct && selProduct.number) || '',
      displayName: selProduct ? selProduct.name : partNumber,
      quantity: qty,
      vendor: $('product-vendor').value.trim() || null,
      shipping: intOf($('product-shipping').value),
      shipToLocation: intOf($('product-shiptolocation').value),
      shipToName: $('product-shiptoname').value.trim() || null,
      installMinutes: intOf($('product-installmins').value),
      systemStatus: intOf($('product-systemstatus').value),
      fromStock: boolOf($('product-fromstock').value),
      partUsed: boolOf($('product-partused').value),
      warranty: boolOf($('product-warranty').value),
      rma: $('product-rma').value.trim() || null,
      additionalInfo: $('product-additionalinfo').value.trim() || null,
    });
    sEl.value=''; selEl.style.display='none'; selProduct=null;
    $('product-partnumber').value=''; $('product-qty').value='1'; $('product-vendor').value='';
    $('product-shipping').value=''; $('product-shiptolocation').value=''; $('product-shiptoname').value='';
    $('product-installmins').value=''; $('product-systemstatus').value=''; $('product-fromstock').value='false';
    $('product-partused').value='true'; $('product-rma').value='';
    $('product-additionalinfo').value='';
    applyContractToPartsForm(); // restore contract-based Warranty/Contract default
    renderPartsTable();
    toast('Part added to list');
  });

  $('submit-parts-request-btn').addEventListener('click', async () => {
    if (!draftParts.length || !woId) return;
    const btn = $('submit-parts-request-btn');
    btn.disabled = true; btn.textContent = 'Submitting…';
    try {
      // Customer Asset is copied from the work order (same as the Dynamics form default).
      const assetId = wo && wo._msdyn_customerasset_value;
      let failed = 0;
      let firstError = '';
      for (const d of draftParts) {
        const payload = {
          msdyn_name: d.product?.name || d.partNumber || 'Part',
          msdyn_quantity: d.quantity,
          'msdyn_workorder@odata.bind': `/msdyn_workorders(${woId})`,
        };
        // Product is required by Dynamics. Use the picked catalog product, or fall
        // back to the work order asset's product for out-of-catalog (New Part #) items.
        const productId = d.product?.id || (customerAsset && customerAsset._msdyn_product_value) || null;
        if (productId) payload['msdyn_product@odata.bind'] = `/products(${productId})`;
        if (assetId)   payload['msdyn_customerasset@odata.bind'] = `/msdyn_customerassets(${assetId})`;
        if (bookingId)  payload['msdyn_booking@odata.bind']        = `/bookableresourcebookings(${bookingId})`;
        if (d.partNumber)         payload.cr217_newpartnumbernotinsystem  = d.partNumber;
        if (d.vendor)             payload.cr217_vendor                    = d.vendor;
        if (d.shipToName)         payload.cr217_shiptoname                = d.shipToName;
        if (d.shipping != null)   payload.cr217_shipping                  = d.shipping;
        if (d.shipToLocation != null) payload.cr217_shiptolocation        = d.shipToLocation;
        if (d.installMinutes != null) payload.cr217_estimatedserviceinstalltime = d.installMinutes;
        if (d.systemStatus != null)   payload.cr217_currentsystemstatus   = d.systemStatus;
        if (d.fromStock != null)  payload.cr217_fromstock                 = d.fromStock;
        if (d.partUsed != null)   payload.cr217_partused                  = d.partUsed;
        if (d.warranty != null)   payload.cr217_warrantycontract          = d.warranty;
        if (d.rma)                payload.cr217_rma                       = d.rma;
        if (d.additionalInfo)     payload.cr217_additionalinformation     = d.additionalInfo;

        try {
          // Create the part as UNsubmitted; the submit step below flags it (matches
          // the Dynamics flow of "add parts, then Submit Parts").
          await xrmCreate('msdyn_workorderproduct', payload);
        } catch (e) { failed++; if (!firstError) firstError = e.message; console.warn('WOP create failed:', e.message); }
      }

      // Replicate the Dynamics "Submit Parts" button exactly (PartRequest.submitPartOrder):
      // flag unsubmitted work order products AND create the wc_partsrequest activity that
      // posts to the timeline and notifies the parts team.
      let submitNote = '';
      try {
        const n = await submitPartsForWorkOrder();
        if (!n) submitNote = ' (nothing to submit)';
      } catch (se) {
        submitNote = ' — but the SUBMIT step failed: ' + se.message + ' (parts team NOT notified)';
        console.error('Submit-parts step failed:', se);
      }

      draftParts = [];
      productsLoaded = false;
      timelineLoaded = false;
      await loadProducts();
      if (failed) { toast(`Created with ${failed} error(s): ${firstError}`, true); console.error('Parts create error:', firstError); }
      else toast('Parts order submitted' + submitNote, !!submitNote);
    } catch(e) {
      toast('Failed: '+e.message, true);
    } finally {
      btn.textContent = 'Submit Parts Order';
      btn.disabled = !draftParts.length;
    }
  });
}

// Build the HTML parts table for the parts-request activity — identical format to the
// Dynamics "Submit Parts" web resource (buildPartsDetailsHTML in PartRequest.js).
function buildPartsDetailsHTML(records) {
  let body = "The following parts have been requested:<br><br>";
  body += "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse; width: 100%;'>";
  body += "<tr style='background-color: #0078d4; color: white;'>";
  body += "<th style='padding: 8px; text-align: left;'>Product</th>";
  body += "<th style='padding: 8px; text-align: center;'>Quantity</th>";
  body += "</tr>";
  records.forEach(rec => {
    const productName = rec.msdyn_name || "Unknown Product";
    const quantity = rec.msdyn_quantity || 0;
    body += "<tr>";
    body += "<td style='padding: 8px;'>" + productName + "</td>";
    body += "<td style='padding: 8px; text-align: center;'>" + quantity + "</td>";
    body += "</tr>";
  });
  body += "</table>";
  body += "<br><em>Submitted: " + new Date().toLocaleString() + "</em>";
  return body;
}

// Replicate the Dynamics "Submit Parts" button (PartRequest.submitPartOrder): flag all
// unsubmitted work order products as submitted and create the wc_partsrequest activity
// (timeline entry + parts-team notification). Returns how many parts were submitted.
async function submitPartsForWorkOrder() {
  if (!woId) return 0;
  const woName = wo?.msdyn_name || 'Work Order';
  const records = await xrmList('msdyn_workorderproduct',
    `?$select=msdyn_workorderproductid,msdyn_name,msdyn_quantity,msdyn_lineorder,msdyn_description` +
    `&$filter=_msdyn_workorder_value eq ${woId} and pmich_new_partrequestsubmitted eq false&$orderby=msdyn_lineorder asc`);
  if (!records.length) return 0;

  const partsDetails = buildPartsDetailsHTML(records);
  // Mark every unsubmitted part as submitted (in parallel, like the web resource).
  await Promise.all(records.map(rec =>
    xrmUpdate('msdyn_workorderproduct', rec.msdyn_workorderproductid, { pmich_new_partrequestsubmitted: true })
  ));
  // Create the Parts Request activity — this is what hits the timeline and notifies the team.
  await xrmCreate('wc_partsrequest', {
    subject: 'Part Request - ' + woName,
    description: partsDetails,
    statecode: 0,
    'regardingobjectid_msdyn_workorder@odata.bind': `/msdyn_workorders(${woId})`,
  });
  return records.length;
}

// ── Timeline ───────────────────────────────────────────────────────────────────
let timelineLoaded = false;
function prettyActivityType(code) {
  if (!code) return 'Activity';
  const map = { wc_partsrequest:'Parts Request', task:'Task', email:'Email', phonecall:'Phone Call', appointment:'Appointment', fax:'Fax', letter:'Letter' };
  return map[code] || String(code).replace(/_/g,' ').replace(/\b\w/g, ch => ch.toUpperCase());
}
function stripHtml(s) { return String(s||'').replace(/<[^>]*>/g,' ').replace(/&nbsp;/g,' ').replace(/&amp;/g,'&').replace(/\s+/g,' ').trim(); }
// Like stripHtml but keeps line breaks so multi-line notes stay readable.
function htmlToText(s) {
  return String(s||'')
    .replace(/<\s*br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|tr|h[1-6])\s*>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/[ \t]+/g, ' ')
    .replace(/[ \t]*\n[ \t]*/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
// If an annotation is an image attachment, return a data: URL for it (else null).
function noteImageSrc(n) {
  if (n && n.isdocument && n.documentbody && n.mimetype && n.mimetype.toLowerCase().startsWith('image/')) {
    return `data:${n.mimetype};base64,${n.documentbody}`;
  }
  return null;
}
// Non-image attachment filename (for a small paperclip line), else null.
function noteFileName(n) {
  if (n && n.isdocument && n.documentbody && !(n.mimetype && n.mimetype.toLowerCase().startsWith('image/'))) {
    return n.filename || 'attachment';
  }
  return null;
}
// Move <img src> to data-src so the image doesn't try to load before we can fetch
// it through the authenticated Dynamics session.
function deferImgSrc(html) {
  return String(html || '').replace(/<img([^>]*?)\ssrc=(["'])(.*?)\2/gi, (m, pre, q, src) => `<img${pre} data-src=${q}${src}${q}`);
}
// Rich-text notes embed images as Dynamics URLs that require auth. Fetch them via
// the warm session (same origin) and return a data: URL.
// Fetch a note/rich-text image through the authenticated Dynamics session and return a
// downscaled JPEG data URL (small enough to cross IPC). Returns null on any failure
// (e.g. no ReadAccess on the Rich Text Attachment entity).
async function fetchAuthedDataUrl(src) {
  if (!src) return null;
  const url = /^https?:/i.test(src) ? src : orgUrl + (src.startsWith('/') ? '' : '/') + src;
  const script = `(async()=>{try{
    const r=await fetch(${JSON.stringify(url)},{credentials:'include'});
    if(!r.ok) return null;
    const b=await r.blob();
    if(!b||!b.size||!/^image\\//.test(b.type||'')) return null;
    let out;
    try{
      const bmp=await createImageBitmap(b);
      const max=1400; let w=bmp.width,h=bmp.height;
      if(w>max||h>max){const s=Math.min(max/w,max/h);w=Math.round(w*s);h=Math.round(h*s);}
      const c=new OffscreenCanvas(w,h); c.getContext('2d').drawImage(bmp,0,0,w,h);
      const jb=await c.convertToBlob({type:'image/jpeg',quality:0.85});
      out=await new Promise(res=>{const fr=new FileReader();fr.onloadend=()=>res(String(fr.result));fr.readAsDataURL(jb);});
    }catch(_){
      out=await new Promise(res=>{const fr=new FileReader();fr.onloadend=()=>res(String(fr.result));fr.readAsDataURL(b);});
    }
    return out;
  }catch(e){return null;}})()`;
  try {
    const res = await Promise.race([
      apiWv.executeJavaScript(script),
      new Promise(r => setTimeout(() => r(null), 12000)),
    ]);
    return (typeof res === 'string' && res.startsWith('data:')) ? res : null;
  } catch (_) { return null; }
}
// Fetch every embedded image through the authenticated session and swap it in.
async function hydrateNoteImages(root) {
  if (!root) return;
  for (const img of [...root.querySelectorAll('img')]) {
    // getAttribute returns the literal url even after the browser mangled a relative src.
    const raw = img.getAttribute('data-src') || img.getAttribute('src') || '';
    img.removeAttribute('data-src');
    if (!raw || raw.startsWith('data:')) continue;
    img.removeAttribute('src'); // stop the failing (cross-origin) load
    img.classList.add('note-img');
    const dataUrl = await fetchAuthedDataUrl(raw);
    if (dataUrl) {
      img.src = dataUrl;
    } else {
      const ph = document.createElement('div');
      ph.className = 'note-photo-missing';
      ph.textContent = '🔒 Photo not viewable here (needs Read access to Rich Text Attachment in Dynamics)';
      img.replaceWith(ph);
    }
  }
}
// Render a note's body: rich HTML (with images) if it's HTML, else plain text.
function renderNoteBodyHtml(notetext) {
  if (!notetext) return '';
  if (/<[a-z][\s\S]*>/i.test(notetext)) return `<div class="note-body note-rich">${deferImgSrc(notetext)}</div>`;
  return `<div class="note-body">${esc(notetext).replace(/\n/g,'<br>')}</div>`;
}

async function loadTimeline() {
  const el = $('timeline-list');
  if (!woId) { el.innerHTML = `<div class="empty-msg">No work order linked</div>`; return; }
  el.innerHTML = `<div class="inline-loading"><div class="spinner"></div></div>`;
  try {
    const [activities, notes] = await Promise.all([
      xrmList('activitypointer', `?$select=activityid,subject,activitytypecode,createdon,description&$filter=_regardingobjectid_value eq ${woId}&$orderby=createdon desc&$top=100`).catch(()=>[]),
      xrmList('annotation', `?$select=subject,notetext,createdon,_createdby_value,isdocument,mimetype,filename,documentbody&$filter=_objectid_value eq ${woId}&$orderby=createdon desc&$top=100`).catch(()=>[]),
    ]);
    const items = [];
    activities.forEach(a => items.push({ when:a.createdon, type:prettyActivityType(a.activitytypecode), title:a.subject||prettyActivityType(a.activitytypecode), body:a.description||'' }));
    notes.forEach(n => items.push({ when:n.createdon, type:'Note', title:n.subject||'Note', body:n.notetext||'', author:n['_createdby_value@OData.Community.Display.V1.FormattedValue']||'', img:noteImageSrc(n), file:noteFileName(n) }));
    items.sort((a,b) => new Date(b.when) - new Date(a.when));
    timelineLoaded = true;
    if (!items.length) { el.innerHTML = `<div class="empty-msg">No timeline activity yet</div>`; return; }
    el.innerHTML = items.map((it, i) => {
      const hasEmbeddedImg = /<img/i.test(it.body || '');
      let bodyBlock;
      if (hasEmbeddedImg) {
        // Rich note with embedded photo(s) — render the HTML and hydrate images.
        bodyBlock = `<div class="note-rich tl-rich">${deferImgSrc(it.body)}</div>`;
      } else {
        const body = it.body ? htmlToText(it.body) : '';
        const long = body.length > 260 || body.split('\n').length > 5;
        bodyBlock = body ? `<div class="tl-body${long ? ' clamp' : ''}" data-i="${i}">${esc(body)}</div>${long ? `<button class="tl-toggle" data-i="${i}">Show more</button>` : ''}` : '';
      }
      return `
      <div style="background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:10px 12px;">
        <div style="display:flex;justify-content:space-between;gap:10px;align-items:baseline;">
          <span style="font-size:12px;font-weight:600;">${esc(it.title)}</span>
          <span style="font-size:11px;color:var(--muted2);white-space:nowrap;">${esc(fmtDate(it.when))}</span>
        </div>
        <div style="font-size:11px;color:var(--accent);margin-top:2px;">${esc(it.type)}${it.author?` · ${esc(it.author)}`:''}</div>
        ${bodyBlock}
        ${it.img ? `<img class="tl-img" src="${it.img}" alt="${esc(it.file||'')}"/>` : (it.file ? `<div class="tl-file">📎 ${esc(it.file)}</div>` : '')}
      </div>`;
    }).join('');

    // Expand/collapse each event's full text
    el.querySelectorAll('.tl-toggle').forEach(btn => {
      btn.addEventListener('click', () => {
        const body = el.querySelector(`.tl-body[data-i="${btn.dataset.i}"]`);
        const collapsed = body.classList.toggle('clamp');
        btn.textContent = collapsed ? 'Show more' : 'Show less';
      });
    });
    hydrateNoteImages(el);
  } catch(e) { el.innerHTML = `<div class="empty-msg">Error: ${esc(e.message)}</div>`; }
}

// ── Notes ─────────────────────────────────────────────────────────────────────
async function loadNotes() {
  $('notes-list').innerHTML=`<div class="inline-loading"><div class="spinner"></div></div>`;
  try {
    const target = woId||bookingId;
    const rows = await xrmList('annotation',
      `?$select=subject,notetext,createdon,_createdby_value,isdocument,mimetype,filename,documentbody&$filter=_objectid_value eq ${target}&$orderby=createdon desc&$top=50`);
    notesLoaded=true;
    if (!rows.length) { $('notes-list').innerHTML=`<div class="empty-msg">No notes yet</div>`; return; }
    $('notes-list').innerHTML = rows.map(n => {
      const author = n['_createdby_value@OData.Community.Display.V1.FormattedValue']||'Unknown';
      const img = noteImageSrc(n), file = noteFileName(n);
      return `<div class="note-card">
        <div class="note-head"><span class="note-subject">${esc(n.subject||'Note')}</span><span class="note-meta">${esc(author)} · ${fmtDate(n.createdon)}</span></div>
        ${renderNoteBodyHtml(n.notetext)}
        ${img ? `<img class="note-img" src="${img}" alt="${esc(file||'')}"/>` : (file ? `<div class="tl-file">📎 ${esc(file)}</div>` : '')}
      </div>`;
    }).join('');
    hydrateNoteImages($('notes-list'));
  } catch(e) { $('notes-list').innerHTML=`<div class="empty-msg">Failed to load notes</div>`; }
}

$('add-note-btn').addEventListener('click', async () => {
  const text = $('note-text').value.trim(); if (!text) { $('note-text').focus(); return; }
  const btn=$('add-note-btn'); btn.disabled=true; btn.textContent='Adding…';
  const bind = woId
    ? {'objectid_msdyn_workorder@odata.bind':`/msdyn_workorders(${woId})`}
    : {'objectid_bookableresourcebooking@odata.bind':`/bookableresourcebookings(${bookingId})`};
  try {
    await xrmCreate('annotation', {subject:$('note-subject').value.trim()||'Note', notetext:text, ...bind});
    $('note-subject').value=''; $('note-text').value='';
    notesLoaded=false; await loadNotes(); toast('Note added');
  } catch(e) { toast('Failed: '+e.message, true); }
  finally { btn.disabled=false; btn.textContent='Add Note'; }
});

// ── Toast ─────────────────────────────────────────────────────────────────────
function toast(msg, err=false) {
  const el=$('toast'); el.textContent=msg; el.className='show'+(err?' err':'');
  clearTimeout(el._t); el._t=setTimeout(()=>el.className='',3000);
}

init();
