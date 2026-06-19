const params  = new URLSearchParams(window.location.hash.slice(1));
const memberName = decodeURIComponent(params.get('name') || '');
const orgUrl  = decodeURIComponent(params.get('org') || '');

const apiWv = document.getElementById('api-wv');
apiWv.src   = `${orgUrl}/main.aspx`;

const $   = id => document.getElementById(id);
const esc = s  => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

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

function fv(obj, field) {
  if (!obj) return '';
  return obj[`${field}@OData.Community.Display.V1.FormattedValue`] || obj[field] || '';
}

function badgeClass(name) {
  const s = (name || '').toLowerCase();
  if (s === 'free') return 'badge-free';
  if (s.includes('travel')) return 'badge-traveling';
  if (s.includes('progress')) return 'badge-inprogress';
  if (s.includes('scheduled')) return 'badge-scheduled';
  return 'badge-default';
}

function displayLabel(rawStatusName) {
  const s = (rawStatusName || '').toLowerCase();
  if (s.includes('travel')) return 'Traveling';
  if (s.includes('progress')) return 'In Progress';
  return 'Scheduled';
}

async function loadData() {
  const safeName = memberName.replace(/'/g, "''");

  $('titlebar-label').textContent = memberName || 'Team Member';
  document.title = memberName || 'Team Member';
  const initials = (memberName || '?').trim().split(/\s+/).map(p => p[0]).slice(0,2).join('').toUpperCase();
  $('t-avatar').textContent = initials || '?';
  $('t-name').textContent   = memberName || '—';

  // Contact info — try matching a system user by name first, then fall back to a contact record
  let user = null, contact = null;
  try {
    const users = await xrmList('systemuser',
      `?$select=fullname,internalemailaddress,mobilephone,telephone1,title&$filter=fullname eq '${safeName}'&$top=1`);
    user = users[0] || null;
  } catch(_) {}
  if (!user) {
    try {
      const contacts = await xrmList('contact',
        `?$select=fullname,emailaddress1,mobilephone,telephone1,jobtitle&$filter=fullname eq '${safeName}'&$top=1`);
      contact = contacts[0] || null;
    } catch(_) {}
  }
  $('t-title').textContent  = user?.title || contact?.jobtitle || '';
  $('t-phone').textContent  = user?.telephone1   || contact?.telephone1   || '—';
  $('t-mobile').textContent = user?.mobilephone  || contact?.mobilephone  || '—';
  $('t-email').textContent  = user?.internalemailaddress || contact?.emailaddress1 || '—';

  // Current status — any booking scheduled for right now, regardless of its status
  const nowIso = new Date().toISOString();
  const bookings = await xrmList('bookableresourcebooking',
    `?$select=name,starttime,endtime,_msdyn_workorder_value&$expand=BookingStatus($select=name)` +
    `&$filter=Resource/name eq '${safeName}' and starttime le ${nowIso} and endtime ge ${nowIso}` +
    `&$orderby=starttime desc`);

  const statusBody = $('t-status-body');
  if (!bookings.length) {
    statusBody.innerHTML = `<span class="status-badge badge-free">Free</span>`;
    return;
  }

  // Fetch work order names + problem descriptions for the active bookings
  const woIds = [...new Set(bookings.map(b => b._msdyn_workorder_value).filter(Boolean))];
  let woMap = {};
  if (woIds.length) {
    try {
      const filter = woIds.map(id => `msdyn_workorderid eq ${id}`).join(' or ');
      const wos = await xrmList('msdyn_workorder',
        `?$select=msdyn_workorderid,msdyn_name,wc_workorderproblemdescription&$filter=${filter}`);
      wos.forEach(w => { woMap[w.msdyn_workorderid] = w; });
    } catch(_) {}
  }

  statusBody.innerHTML = `<div class="wo-list">` + bookings.map(b => {
    const label   = displayLabel(b.BookingStatus?.name);
    const wo      = woMap[b._msdyn_workorder_value];
    const woName  = wo?.msdyn_name || b.name || 'Work Order';
    const problem = wo?.wc_workorderproblemdescription || '';
    return `<div class="wo-item" data-problem="${esc(problem)}">
      <span class="wo-item-title">${esc(woName)}</span>
      <span class="status-badge ${badgeClass(label)}">${esc(label)}</span>
    </div>`;
  }).join('') + `</div>`;

  const tooltip = $('t-tooltip');
  statusBody.querySelectorAll('.wo-item').forEach(el => {
    el.addEventListener('mouseenter', () => {
      const problem = el.dataset.problem;
      if (!problem) return;
      tooltip.textContent = problem;
      tooltip.classList.add('visible');
    });
    el.addEventListener('mousemove', e => {
      if (!tooltip.classList.contains('visible')) return;
      const x = e.clientX + 14, y = e.clientY + 14;
      const overRight = x + 270 > window.innerWidth;
      tooltip.style.left = (overRight ? e.clientX - 270 : x) + 'px';
      tooltip.style.top  = Math.min(y, window.innerHeight - tooltip.offsetHeight - 8) + 'px';
    });
    el.addEventListener('mouseleave', () => tooltip.classList.remove('visible'));
  });
}

async function init() {
  try {
    await waitForXrm();
    await loadData();
    $('t-loading').style.display = 'none';
    $('t-content').style.display = 'flex';
  } catch (e) {
    $('t-loading').innerHTML = `<div style="color:var(--danger);text-align:center;padding:20px;max-width:400px;">${esc(e.message)}</div>`;
  }
}

$('refresh-btn')?.addEventListener('click', async () => {
  const btn = $('refresh-btn');
  btn.disabled = true;
  const original = btn.textContent;
  btn.textContent = 'Refreshing…';
  try {
    await loadData();
  } catch (e) {
    $('t-loading').innerHTML = `<div style="color:var(--danger);text-align:center;padding:20px;max-width:400px;">${esc(e.message)}</div>`;
    $('t-loading').style.display = 'flex';
    $('t-content').style.display = 'none';
  } finally {
    btn.disabled = false;
    btn.textContent = original;
  }
});

init();
