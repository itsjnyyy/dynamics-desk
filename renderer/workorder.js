const params    = new URLSearchParams(window.location.hash.slice(1));
const bookingId = decodeURIComponent(params.get('bid') || '');
const directWoId = decodeURIComponent(params.get('wo') || '');
const orgUrl    = decodeURIComponent(params.get('org') || '');

const apiWv = document.getElementById('api-wv');
apiWv.src   = `${orgUrl}/main.aspx`;

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
  for (let i = 0; i < 60; i++) {
    await sleep(500);
    try { if (await apiWv.executeJavaScript('typeof Xrm!=="undefined"&&!!Xrm.WebApi')) return; } catch(_) {}
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
  await apiWv.executeJavaScript(`window.__xd=${JSON.stringify(data)}`);
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{await Xrm.WebApi.updateRecord("${entity}","${id}",window.__xd);return JSON.stringify({ok:1});}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
}
async function xrmCreate(entity, data) {
  await apiWv.executeJavaScript(`window.__xd=${JSON.stringify(data)}`);
  const r = JSON.parse(await apiWv.executeJavaScript(
    `(async()=>{try{const r=await Xrm.WebApi.createRecord("${entity}",window.__xd);return JSON.stringify({id:r.id});}catch(e){return JSON.stringify({__err:e.message})}})()`
  ));
  if (r?.__err) throw new Error(r.__err);
  return r.id;
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
    wo = await xrmGet('msdyn_workorder', woId,
      '?$select=msdyn_name,msdyn_systemstatus,msdyn_workordersummary,msdyn_instructions,' +
      'msdyn_address1,msdyn_address2,msdyn_city,msdyn_stateorprovince,msdyn_postalcode,msdyn_country,' +
      '_msdyn_serviceaccount_value,_msdyn_billingaccount_value,_msdyn_workordertype_value,' +
      '_msdyn_serviceterritory_value,_msdyn_substatus_value,_msdyn_priority_value,_msdyn_customerasset_value,' +
      'msdyn_datewindowstart,msdyn_datewindowend,msdyn_timetopromised,msdyn_timefrompromised,' +
      'wc_workorderproblemdescription,_msdyn_reportedbycontact_value');
    const contactId = wo._msdyn_reportedbycontact_value;
    contact = null;
    if (contactId) {
      try {
        contact = await xrmGet('contact', contactId,
          '?$select=fullname,telephone1,mobilephone,emailaddress1,jobtitle');
      } catch(_) {}
    }

    const assetId = wo._msdyn_customerasset_value;
    customerAsset = null;
    if (assetId) {
      try {
        customerAsset = await xrmGet('msdyn_customerasset', assetId,
          '?$select=msdyn_name,wc_assettag,msdyn_assettag,wc_seriallotnumber');
        customerAsset.__tag    = customerAsset.wc_assettag || customerAsset.msdyn_assettag;
        customerAsset.__serial = customerAsset.wc_seriallotnumber;
      } catch(_) {}
    }

    const incidents = await xrmList('msdyn_workorderincident',
      `?$filter=_msdyn_workorder_value eq ${woId}&$top=1`);
    incident = incidents[0] || null;

    try {
      subStatuses = await xrmList('msdyn_workordersubstatus',
        '?$select=msdyn_workordersubstatusid,msdyn_name&$orderby=msdyn_name asc');
    } catch(_) { subStatuses = []; }

    try {
      await loadEngineers();
    } catch(_) {}
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
  const APP_ID = 'YOUR-MODEL-DRIVEN-APP-ID'; // find this in your app's URL in Dynamics
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
    $('f-arrival').value = isoToLocal(
      Object.entries(booking).find(([k,v]) => !k.includes('@') && (k.toLowerCase().includes('arriv') || k.toLowerCase().includes('actual')) && v)?.[1] || ''
    );
    // debug: show candidate fields in placeholder
    const candidates = Object.entries(booking).filter(([k,v]) => !k.includes('@') && typeof v === 'string' && v.includes('T') && v.includes('Z')).map(([k,v])=>`${k}: ${v}`).join(' | ');
    if (!$('f-arrival').value) $('f-arrival').title = candidates;
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
    if (snap._actualarrival)  bPatch.msdyn_actualarrival = new Date(snap._actualarrival).toISOString();
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
    if (snap._actualarrival) booking.msdyn_actualarrival = new Date(snap._actualarrival).toISOString();
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
async function loadEngineers() {
  const el = $('d-engineers-list');
  if (!el) return;
  const bookings = await xrmList('bookableresourcebooking',
    `?$select=bookableresourcebookingid,name,starttime,endtime&$expand=Resource($select=name)&$filter=_msdyn_workorder_value eq ${woId}&$orderby=starttime asc`);

  if (!bookings.length) { el.innerHTML = `<div class="field-value dim">No engineers assigned</div>`; return; }

  el.innerHTML = bookings.map((b, i) => {
    const name = b.Resource?.name || 'Unassigned';
    const time = `${fmtDate(b.starttime)} – ${fmtDate(b.endtime)}`;
    return `
      <div class="engineer-row" data-idx="${i}" style="background:var(--bg3);border:1px solid var(--border);border-radius:8px;padding:8px 12px;display:flex;justify-content:space-between;align-items:center;cursor:pointer;">
        <span style="font-size:13px;">${esc(name)}</span>
        <span style="font-size:11px;color:var(--muted2);">${esc(time)}</span>
      </div>`;
  }).join('');

  el.querySelectorAll('.engineer-row').forEach(row => {
    row.addEventListener('mouseenter', () => row.style.borderColor = 'var(--accent)');
    row.addEventListener('mouseleave', () => row.style.borderColor = 'var(--border)');
    row.addEventListener('click', () => {
      const b = bookings[+row.dataset.idx];
      window.api.openWorkOrder(b.bookableresourcebookingid, orgUrl, b.Resource?.name || b.name || 'Booking');
    });
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
    if (btn.dataset.tab==='products' && !productsLoaded) { initProdSearch(); loadProducts(); }
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

// ── Products ──────────────────────────────────────────────────────────────────
async function loadProducts() {
  if (!woId) { $('products-body').innerHTML=`<tr><td colspan="4"><div class="empty-msg">No work order linked</div></td></tr>`; return; }
  try {
    const rows = await xrmList('msdyn_workorderproduct',
      `?$select=msdyn_name,_msdyn_product_value,msdyn_quantity,msdyn_linestatus,_msdyn_unit_value&$filter=_msdyn_workorder_value eq ${woId}`);
    productsLoaded = true;
    if (!rows.length) { $('products-body').innerHTML=`<tr><td colspan="4"><div class="empty-msg">No products</div></td></tr>`; return; }
    $('products-body').innerHTML = rows.map(p => {
      const name   = p['_msdyn_product_value@OData.Community.Display.V1.FormattedValue']||p.msdyn_name||'—';
      const unit   = p['_msdyn_unit_value@OData.Community.Display.V1.FormattedValue']||'—';
      const qty    = p.msdyn_quantity??'—';
      const used   = p.msdyn_linestatus===690970001;
      const cls    = used?'badge-inprogress':'badge-scheduled';
      return `<tr>
        <td>${esc(name)}</td><td class="col-muted">${esc(qty)}</td><td class="col-muted">${esc(unit)}</td>
        <td><span class="status-badge ${cls}" style="font-size:10px;padding:2px 8px;">${used?'Used':'Estimated'}</span></td>
      </tr>`;
    }).join('');
  } catch(e) { $('products-body').innerHTML=`<tr><td colspan="4"><div class="empty-msg">Error: ${esc(e.message)}</div></td></tr>`; }
}

let selProduct=null, searchTimer=null;
function initProdSearch() {
  if (prodSearchInited) return; prodSearchInited=true;
  const sEl=$('product-search'), rEl=$('product-results'), selEl=$('product-selected'), addBtn=$('add-product-btn');
  sEl.addEventListener('input', () => {
    clearTimeout(searchTimer); selProduct=null; addBtn.disabled=true; selEl.style.display='none';
    const q=sEl.value.trim();
    if (q.length<2) { rEl.style.display='none'; rEl.innerHTML=''; return; }
    searchTimer = setTimeout(async () => {
      try {
        const res = await xrmList('product', `?$select=productid,name&$filter=contains(name,'${q.replace(/'/g,"''")}') and statecode eq 0&$top=10&$orderby=name asc`);
        if (!res.length) { rEl.innerHTML='<div class="prod-option col-muted">No results</div>'; rEl.style.display='block'; return; }
        rEl.innerHTML = res.map(p=>`<div class="prod-option" data-id="${p.productid}" data-name="${esc(p.name)}">${esc(p.name)}</div>`).join('');
        rEl.style.display='block';
        rEl.querySelectorAll('.prod-option').forEach(el => el.addEventListener('click', () => {
          selProduct={id:el.dataset.id,name:el.dataset.name};
          sEl.value=el.dataset.name; rEl.style.display='none'; rEl.innerHTML='';
          selEl.textContent=`✓ ${el.dataset.name}`; selEl.style.display='block';
          addBtn.disabled=false;
        }));
      } catch(_) { rEl.innerHTML='<div class="prod-option col-muted">Search failed</div>'; rEl.style.display='block'; }
    }, 350);
  });
  addBtn.addEventListener('click', async () => {
    if (!selProduct||!woId) return;
    addBtn.disabled=true; addBtn.textContent='Adding…';
    try {
      await xrmCreate('msdyn_workorderproduct', {
        'msdyn_workorder@odata.bind':`/msdyn_workorders(${woId})`,
        'msdyn_product@odata.bind':`/products(${selProduct.id})`,
        msdyn_quantity: parseFloat($('product-qty').value)||1,
        msdyn_linestatus: 690970000,
      });
      sEl.value=''; selEl.style.display='none'; selProduct=null; addBtn.disabled=true;
      productsLoaded=false; await loadProducts(); toast('Part added');
    } catch(e) { toast('Failed: '+e.message, true); }
    finally { addBtn.textContent='Add'; addBtn.disabled=!selProduct; }
  });
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
