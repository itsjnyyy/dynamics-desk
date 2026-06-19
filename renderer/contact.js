const params    = new URLSearchParams(window.location.hash.slice(1));
const contactId = decodeURIComponent(params.get('cid') || '');
const orgUrl    = decodeURIComponent(params.get('org') || '');

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

function fv(obj, field) {
  if (!obj) return '';
  return obj[`${field}@OData.Community.Display.V1.FormattedValue`] || obj[field] || '';
}

async function init() {
  try {
    await waitForXrm();
    const c = await xrmGet('contact', contactId,
      '?$select=fullname,jobtitle,telephone1,mobilephone,emailaddress1,_parentcustomerid_value');

    const initials = (c.fullname || '?').trim().split(/\s+/).map(p => p[0]).slice(0,2).join('').toUpperCase();

    $('titlebar-label').textContent = c.fullname || 'Contact';
    document.title = c.fullname || 'Contact';
    $('c-avatar').textContent = initials || '?';
    $('c-name').textContent   = c.fullname || '—';
    $('c-title').textContent  = c.jobtitle || '';
    $('c-phone').textContent  = c.telephone1   || '—';
    $('c-mobile').textContent = c.mobilephone  || '—';
    $('c-email').textContent  = c.emailaddress1 || '—';
    $('c-account').textContent = fv(c, '_parentcustomerid_value') || '—';

    const APP_ID = 'YOUR-MODEL-DRIVEN-APP-ID'; // find this in your app's URL in Dynamics
    const url = `${orgUrl}/main.aspx?appid=${APP_ID}&pagetype=entityrecord&etn=contact&id=${contactId}`;
    $('open-dynamics-btn').addEventListener('click', () => window.api.openExternal(url));

    $('c-loading').style.display = 'none';
    $('c-content').style.display = 'flex';
  } catch (e) {
    $('c-loading').innerHTML = `<div style="color:var(--danger);text-align:center;padding:20px;max-width:400px;">${esc(e.message)}</div>`;
  }
}

init();
