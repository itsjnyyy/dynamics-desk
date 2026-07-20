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
let booking = null, wo = null, woId = null, incident = null, contact = null, customerAsset = null, bookingStatuses = [], resources = [], subStatuses = [], dirty = {}, substatusNavProp = null;
let tasksLoaded = false, productsLoaded = false, notesLoaded = false, prodSearchInited = false;

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

    wo = await xrmGet('msdyn_workorder', woId,
      '?$select=msdyn_name,msdyn_systemstatus,msdyn_workordersummary,msdyn_instructions,' +
      'msdyn_address1,msdyn_address2,msdyn_city,msdyn_stateorprovince,msdyn_postalcode,msdyn_country,' +
      '_msdyn_serviceaccount_value,_msdyn_billingaccount_value,_msdyn_workordertype_value,' +
      '_msdyn_serviceterritory_value,_msdyn_substatus_value,_msdyn_priority_value,_msdyn_customerasset_value,' +
      'msdyn_datewindowstart,msdyn_datewindowend,msdyn_timetopromised,msdyn_timefrompromised,' +
      'wc_workorderproblemdescription,_msdyn_reportedbycontact_value');

    // Everything below only needs woId / values already on `wo`, so fire them all
    // off at once instead of one round-trip at a time.
    const contactId = wo._msdyn_reportedbycontact_value;
    const assetId = wo._msdyn_customerasset_value;
    const [contactR, assetR, incidents, subStatusesR] = await Promise.all([
      contactId
        ? xrmGet('contact', contactId, '?$select=fullname,telephone1,mobilephone,emailaddress1,jobtitle').catch(() => null)
        : Promise.resolve(null),
      assetId
        ? xrmGet('msdyn_customerasset', assetId, '?$select=msdyn_name,wc_assettag,msdyn_assettag,wc_seriallotnumber').catch(() => null)
        : Promise.resolve(null),
      xrmList('msdyn_workorderincident', `?$filter=_msdyn_workorder_value eq ${woId}&$top=1`).catch(() => []),
      subStatusesP,
      loadEngineers().catch(() => {}),
    ]);

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
    if (wo) buildSubstatusDropdown();
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
      if (wo) buildSubstatusDropdown();
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

function wireOpenDynamics() {
  const APP_ID = 'YOUR-MODEL-DRIVEN-APP-ID';
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

function renderAll() {
  const woNum      = wo?.msdyn_name || '';
  const account    = wo ? fv(wo,'_msdyn_serviceaccount_value') : '';

  $('titlebar-label').textContent = woNum || booking?.name || 'Work Order';
  $('wo-number').textContent      = woNum || '—';
  $('wo-account').textContent     = account;
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
  set('d-type',       wo ? fv(wo,'_msdyn_workordertype_value')   : '—');
  set('d-priority',   wo ? fv(wo,'_msdyn_priority_value')        : '—');
  set('d-account',    wo ? fv(wo,'_msdyn_serviceaccount_value')  : '—');
  set('d-contact-name',  contact?.fullname      || '—');
  set('d-contact-phone', contact?.telephone1 || contact?.mobilephone || '—');
  set('d-contact-email', contact?.emailaddress1 || '—');
  set('d-contact-title', contact?.jobtitle      || '');
  set('d-billing',    wo ? fv(wo,'_msdyn_billingaccount_value')  : '—');
  set('d-territory',  wo ? fv(wo,'_msdyn_serviceterritory_value'): '—');
  set('d-asset-tag',     wo ? fv(wo,'_msdyn_customerasset_value') : '—');
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
  $('f-summary').value      = wo?.msdyn_workordersummary || '';
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
  $('f-start').addEventListener('input',   e => { dirty._starttime = e.target.value; showSave(); });
  $('f-end').addEventListener('input',     e => { dirty._endtime   = e.target.value; showSave(); });
  $('f-arrival').addEventListener('input', e => { dirty._actualarrival = e.target.value; showSave(); });
  $('f-resource').addEventListener('change', e => { dirty._resource = e.target.value; showSave(); });

  [['f-addr1','msdyn_address1'],['f-addr2','msdyn_address2'],['f-city','msdyn_city'],
   ['f-state','msdyn_stateorprovince'],['f-zip','msdyn_postalcode'],['f-country','msdyn_country'],
   ['f-summary','msdyn_workordersummary'],['f-instructions','msdyn_instructions'],
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
    const bPatch = {};
    if (snap._bookingStatus) bPatch['BookingStatus@odata.bind'] = `/bookingstatuses(${snap._bookingStatus})`;
    if (snap._starttime)      bPatch.starttime          = new Date(snap._starttime).toISOString();
    if (snap._endtime)        bPatch.endtime            = new Date(snap._endtime).toISOString();
    if (snap._actualarrival)  bPatch.msdyn_actualarrivaltime = new Date(snap._actualarrival).toISOString();
    if (snap._resource)       bPatch['Resource@odata.bind'] = `/bookableresources(${snap._resource})`;
    if (Object.keys(bPatch).length) await xrmUpdate('bookableresourcebooking', bookingId, bPatch);

    const wPatch = {};
    ['msdyn_address1','msdyn_address2','msdyn_city','msdyn_stateorprovince',
     'msdyn_postalcode','msdyn_country','msdyn_workordersummary','msdyn_instructions','wc_workorderproblemdescription'
    ].forEach(k => { if (snap[k] !== undefined) wPatch[k] = snap[k]; });
    if (woId && Object.keys(wPatch).length) await xrmUpdate('msdyn_workorder', woId, wPatch);

    if (woId && snap._substatus) {
      if (!substatusNavProp) substatusNavProp = await getLookupNavProperty('msdyn_workorder', 'msdyn_substatus');
      await xrmUpdate('msdyn_workorder', woId, { [`${substatusNavProp}@odata.bind`]: `/msdyn_workordersubstatuses(${snap._substatus})` });
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
  if (wo) buildSubstatusDropdown();
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
    if (btn.dataset.tab==='tasks'    && !tasksLoaded)    loadTasks();
    if (btn.dataset.tab==='products') { initProdSearch(); loadProducts(); }
    if (btn.dataset.tab==='timeline') loadTimeline();
    if (btn.dataset.tab==='notes'    && !notesLoaded)    loadNotes();
    if (btn.dataset.tab==='details') { /* always loaded */ }
  });
});

// ── Service Tasks ─────────────────────────────────────────────────────────────
async function loadTasks() {
  if (!woId) { $('tasks-body').innerHTML=`<tr><td colspan="5"><div class="empty-msg">No work order linked</div></td></tr>`; return; }
  try {
    const rows = await xrmList('msdyn_workorderservicetask',
      `?$select=msdyn_workorderservicetaskid,msdyn_name,_msdyn_tasktype_value,msdyn_description,msdyn_percentcomplete,msdyn_estimatedduration&$filter=_msdyn_workorder_value eq ${woId}&$orderby=msdyn_name asc`);
    tasksLoaded = true;
    if (!rows.length) { $('tasks-body').innerHTML=`<tr><td colspan="5"><div class="empty-msg">No service tasks</div></td></tr>`; return; }
    $('tasks-body').innerHTML = rows.map(t => {
      const type = t['_msdyn_tasktype_value@OData.Community.Display.V1.FormattedValue']||'—';
      const pct  = t.msdyn_percentcomplete ?? 0;
      const done = pct === 100;
      const dur  = t.msdyn_estimatedduration!=null ? `${t.msdyn_estimatedduration} min` : '—';
      return `<tr class="${done?'task-done':''}" data-tid="${t.msdyn_workorderservicetaskid}">
        <td><label class="task-label"><input type="checkbox" class="task-cb" ${done?'checked':''}><span>${esc(t.msdyn_name||'—')}</span></label></td>
        <td class="col-muted">${esc(type)}</td>
        <td><div class="pct-wrap"><div class="pct-bar"><div class="pct-fill" style="width:${pct}%"></div></div><span class="pct-lbl">${pct}%</span></div></td>
        <td class="col-muted">${esc(dur)}</td>
        <td class="col-muted">${esc(t.msdyn_description||'—')}</td>
      </tr>`;
    }).join('');
    $('tasks-body').querySelectorAll('.task-cb').forEach(cb => {
      cb.addEventListener('change', async () => {
        const row = cb.closest('tr'); cb.disabled = true;
        const done = cb.checked;
        try {
          await xrmUpdate('msdyn_workorderservicetask', row.dataset.tid, {msdyn_percentcomplete:done?100:0});
          row.classList.toggle('task-done', done);
          row.querySelector('.pct-fill').style.width = (done?100:0)+'%';
          row.querySelector('.pct-lbl').textContent  = (done?100:0)+'%';
          toast(done?'Marked complete':'Marked incomplete');
        } catch(e) { cb.checked=!done; toast('Failed: '+e.message, true); }
        finally { cb.disabled=false; }
      });
    });
  } catch(e) { $('tasks-body').innerHTML=`<tr><td colspan="5"><div class="empty-msg">Error: ${esc(e.message)}</div></td></tr>`; }
}

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
    if (!hidden) { loadPartsOptions(); $('product-search').focus(); }
  });
  loadPartsOptions();

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
    $('product-installmins').value=''; $('product-systemstatus').value=''; $('product-fromstock').value='';
    $('product-partused').value=''; $('product-warranty').value=''; $('product-rma').value='';
    $('product-additionalinfo').value='';
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
      for (const d of draftParts) {
        const payload = {
          msdyn_name: d.product?.name || d.partNumber || 'Part',
          msdyn_quantity: d.quantity,
          'msdyn_workorder@odata.bind': `/msdyn_workorders(${woId})`,
        };
        if (d.product)  payload['msdyn_product@odata.bind']       = `/products(${d.product.id})`;
        if (assetId)    payload['msdyn_customerasset@odata.bind']  = `/msdyn_customerassets(${assetId})`;
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
          const id = await xrmCreate('msdyn_workorderproduct', payload);
          // "Submit Parts" ribbon action = flag the part request as submitted.
          try { await xrmUpdate('msdyn_workorderproduct', id, { pmich_new_partrequestsubmitted: true }); } catch(_) {}
        } catch (e) { failed++; console.warn('WOP create failed:', e.message); }
      }
      draftParts = [];
      productsLoaded = false;
      timelineLoaded = false;
      await loadProducts();
      toast(failed ? `Submitted with ${failed} error(s)` : 'Parts order submitted', !!failed);
    } catch(e) {
      toast('Failed: '+e.message, true);
    } finally {
      btn.textContent = 'Submit Parts Order';
      btn.disabled = !draftParts.length;
    }
  });
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
async function loadTimeline() {
  const el = $('timeline-list');
  if (!woId) { el.innerHTML = `<div class="empty-msg">No work order linked</div>`; return; }
  el.innerHTML = `<div class="inline-loading"><div class="spinner"></div></div>`;
  try {
    const [activities, notes] = await Promise.all([
      xrmList('activitypointer', `?$select=activityid,subject,activitytypecode,createdon,description&$filter=_regardingobjectid_value eq ${woId}&$orderby=createdon desc&$top=100`).catch(()=>[]),
      xrmList('annotation', `?$select=subject,notetext,createdon,_createdby_value&$filter=_objectid_value eq ${woId}&$orderby=createdon desc&$top=100`).catch(()=>[]),
    ]);
    const items = [];
    activities.forEach(a => items.push({ when:a.createdon, type:prettyActivityType(a.activitytypecode), title:a.subject||prettyActivityType(a.activitytypecode), body:a.description||'' }));
    notes.forEach(n => items.push({ when:n.createdon, type:'Note', title:n.subject||'Note', body:n.notetext||'', author:n['_createdby_value@OData.Community.Display.V1.FormattedValue']||'' }));
    items.sort((a,b) => new Date(b.when) - new Date(a.when));
    timelineLoaded = true;
    if (!items.length) { el.innerHTML = `<div class="empty-msg">No timeline activity yet</div>`; return; }
    el.innerHTML = items.map((it, i) => {
      const body = it.body ? htmlToText(it.body) : '';
      // "Long" = worth collapsing behind a toggle so the list stays scannable.
      const long = body.length > 260 || body.split('\n').length > 5;
      return `
      <div style="background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:10px 12px;">
        <div style="display:flex;justify-content:space-between;gap:10px;align-items:baseline;">
          <span style="font-size:12px;font-weight:600;">${esc(it.title)}</span>
          <span style="font-size:11px;color:var(--muted2);white-space:nowrap;">${esc(fmtDate(it.when))}</span>
        </div>
        <div style="font-size:11px;color:var(--accent);margin-top:2px;">${esc(it.type)}${it.author?` · ${esc(it.author)}`:''}</div>
        ${body ? `<div class="tl-body${long ? ' clamp' : ''}" data-i="${i}">${esc(body)}</div>${long ? `<button class="tl-toggle" data-i="${i}">Show more</button>` : ''}` : ''}
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
  } catch(e) { el.innerHTML = `<div class="empty-msg">Error: ${esc(e.message)}</div>`; }
}

// ── Notes ─────────────────────────────────────────────────────────────────────
async function loadNotes() {
  $('notes-list').innerHTML=`<div class="inline-loading"><div class="spinner"></div></div>`;
  try {
    const target = woId||bookingId;
    const rows = await xrmList('annotation',
      `?$select=subject,notetext,createdon,_createdby_value&$filter=_objectid_value eq ${target}&$orderby=createdon desc&$top=50`);
    notesLoaded=true;
    if (!rows.length) { $('notes-list').innerHTML=`<div class="empty-msg">No notes yet</div>`; return; }
    $('notes-list').innerHTML = rows.map(n => {
      const author = n['_createdby_value@OData.Community.Display.V1.FormattedValue']||'Unknown';
      return `<div class="note-card">
        <div class="note-head"><span class="note-subject">${esc(n.subject||'Note')}</span><span class="note-meta">${esc(author)} · ${fmtDate(n.createdon)}</span></div>
        <div class="note-body">${esc(n.notetext||'').replace(/\n/g,'<br>')}</div>
      </div>`;
    }).join('');
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
