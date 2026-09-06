/**
 * MRTH JOB MANAGER — mobile web app
 * ---------------------------------------------------------------
 * ONE source of truth: the "Bookings" sheet.
 * The app is the only thing you touch. No double entry, ever.
 *
 * This is a SEPARATE, STANDALONE project — it does not touch your
 * working form backend at all. Nothing you already have can break.
 *
 * ── SETUP (about 5 minutes) ────────────────────────────────────
 * 1. Open your bookings Google Sheet and copy its ID from the URL:
 *      docs.google.com/spreadsheets/d/THIS_LONG_BIT_HERE/edit
 *    Paste it into SHEET_ID below.
 *
 * 2. Go to script.google.com > New project.
 *    Delete the sample code, paste this whole file in, save.
 *    Name the project "MRTH Job Manager".
 *
 * 3. Run the function testSetup once (pick it from the dropdown
 *    next to the Run button). Approve the permissions prompt.
 *    You should see a success message in the log.
 *
 * 4. Deploy > New deployment > Web app
 *      Execute as:     Me
 *      Who has access: ONLY MYSELF     <-- keeps customer data private
 *    Copy the URL.
 *
 * 5. Open that URL on your phone > Share > Add to Home Screen.
 *    That is your app icon. Opens like any other app.
 *
 * ── AFTER ANY CODE EDIT ────────────────────────────────────────
 * Deploy > Manage deployments > pencil icon > Version: New version
 * > Deploy. Same URL, new code. Skipping this is why apps "stop
 * working" after an edit.
 * ───────────────────────────────────────────────────────────────
 */

// PASTE YOUR SHEET ID HERE (from the spreadsheet URL)
const SHEET_ID = "1DWyrxMULrVglhRFD1PG8jpzx3LX0So-YtG1pYEuOPSc";

const SHEET_NAME = 'Enquiries';

// Columns the app needs. Missing ones get added automatically.
const REQUIRED_COLS = [
  'Received','Status','Type','Name','Phone','Email','Material','Quantity',
  'Suburb','Timeframe','Access notes','Quoted $','Notes','Carrier','Load Type',
  'Delivery address','Paid date','Delivered date','Invoice No','Invoice Link',
  'Scheduled date','Start Date','End Date','Calendar Event ID'
];

const STATUSES = ['New','Quoted','Paid','Booked','Delivered','Closed'];

/** Serve the app. */
function doGet() {
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('MRTH Jobs')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ============ SHEET HELPERS ============ */

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(REQUIRED_COLS);
    sh.getRange(1, 1, 1, REQUIRED_COLS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // make sure every column we need exists
  const lastCol = Math.max(sh.getLastColumn(), 1);
  let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  REQUIRED_COLS.forEach(function (c) {
    if (headers.indexOf(c) === -1) {
      sh.getRange(1, headers.length + 1).setValue(c).setFontWeight('bold');
      headers.push(c);
    }
  });
  return sh;
}

function headerMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const map = {};
  headers.forEach(function (h, i) { map[h] = i; });
  return map;
}

/* ============ API CALLED FROM THE PHONE ============ */

/** Every job, newest first. */
function getJobs() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const map = headerMap_(sh);
  const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const get = function (k) {
      const v = map[k] === undefined ? '' : r[map[k]];
      return v === null || v === undefined ? '' : String(v).replace(/^'/, '');
    };
    let received = map['Received'] !== undefined ? r[map['Received']] : '';
    if (received instanceof Date) received = Utilities.formatDate(received, Session.getScriptTimeZone(), 'd MMM h:mm a');
    const fmtDay = function (v) {
      if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'd MMM yyyy');
      return v ? String(v) : '';
    };
    out.push({
      row: i + 2,
      received: String(received),
      status: get('Status') || 'New',
      type: get('Type'),
      name: get('Name'),
      phone: get('Phone'),
      email: get('Email'),
      material: get('Material'),
      quantity: get('Quantity'),
      suburb: get('Suburb'),
      timeframe: get('Timeframe'),
      access: get('Access notes'),
      quoted: get('Quoted $'),
      carrier: get('Carrier'),
      notes: get('Notes'),
      loadType: get('Load Type'),
      address: get('Delivery address'),
      invoiceNo: get('Invoice No'),
      invoiceLink: get('Invoice Link'),
      scheduled: fmtDay(get('Scheduled date')),
      startDate: fmtDay(get('Start Date')),
      endDate: fmtDay(get('End Date'))
    });
  }
  return out.reverse();
}

/** Update one field on one job. */
function updateJob(row, field, value) {
  const sh = sheet_();
  const map = headerMap_(sh);
  if (map[field] === undefined) throw new Error('Unknown field: ' + field);
  if (field === 'Phone') value = "'" + value;
  sh.getRange(row, map[field] + 1).setValue(value);
  if (field === 'Status') stampStatusDate_(sh, map, row, value);
  syncJobCalendarEvent_(row);
  return true;
}

/** Paid/Delivered dates are stamped automatically the first time a chip is
 *  tapped — never overwritten, and never asked for, since this is filled in
 *  from a phone in the field. */
function stampStatusDate_(sh, map, row, status) {
  const col = status === 'Paid' ? 'Paid date' : status === 'Delivered' ? 'Delivered date' : null;
  if (!col || map[col] === undefined) return;
  const cell = sh.getRange(row, map[col] + 1);
  if (!cell.getValue()) cell.setValue(new Date());
}

/** Add an enquiry that came in by Marketplace, text or phone. */
function addJob(obj) {
  const sh = sheet_();
  const map = headerMap_(sh);
  const width = sh.getLastColumn();
  const row = new Array(width).fill('');
  const set = function (k, v) { if (map[k] !== undefined) row[map[k]] = v; };
  set('Received', new Date());
  set('Status', 'New');
  set('Type', obj.type || 'Materials delivery');
  set('Name', obj.name || '');
  set('Phone', obj.phone ? "'" + obj.phone : '');
  set('Email', obj.email || '');
  set('Material', obj.material || '');
  set('Quantity', obj.quantity || '');
  set('Suburb', obj.suburb || '');
  set('Timeframe', obj.timeframe || '');
  set('Access notes', obj.access || '');
  set('Notes', obj.notes || '');
  sh.appendRow(row);
  return true;
}

/** Price list for the in-app quoter. Mirrors the July 2026 guide. */
function getPrices() {
  return {
    'Mulch (per m3)': [
      ['Marri Chips', 150, 'm3', 0], ['Pine Bark', 150, 'm3', 0],
      ['Bushland', 130, 'm3', 0], ['Woodland Brown/Black', 130, 'm3', 0],
      ['Karri & Peat', 130, 'm3', 0], ['Derby Brown', 130, 'm3', 0],
      ['Plantation (budget)', 110, 'm3', 0]
    ],
    'Soils (per m3)': [
      ['Vegie / Garden Blend', 130, 'm3', 0], ['Soil Conditioner', 125, 'm3', 0],
      ['Landscape Blend', 125, 'm3', 0], ['Lawn Blend', 95, 'm3', 0]
    ],
    'Sands (per tonne)': [
      ['Screened Sand', 50, 't', 1.5], ['Fill / House Pad', 55, 't', 1.5],
      ['Brickies Sand', 80, 't', 1.5], ['Grit Sand', 85, 't', 1.5],
      ['Plasterers Sand', 130, 't', 1.5], ['Playground Sand', 130, 't', 1.5]
    ],
    'Stones & Pebbles (per m3)': [
      ['Fine River Gravel', 500, 'm3', 0], ['Fine Rainbow Quartz', 500, 'm3', 0],
      ['Rainbow Quartz', 480, 'm3', 0], ['Autumn Stone', 460, 'm3', 0],
      ['River Pebbles 20-40mm', 460, 'm3', 0], ['River Pebbles 40-90mm', 460, 'm3', 0],
      ['Cream Brick Rubble', 230, 'm3', 0], ['Limestone Rubble', 140, 'm3', 0]
    ],
    'Gravel & Road Base': [
      ['Cracked Gravel 5/12/18mm', 260, 'm3', 0], ['Filter Rock 19-25mm', 260, 'm3', 0],
      ['Laterite Fines', 240, 'm3', 0], ['MRD Gravel Road Base', 54, 't', 1.8],
      ['MRD Limestone Road Base', 56, 't', 1.8]
    ],
    'Blue Metal (per tonne)': [
      ['Blue Metal 10/14/20/50mm', 120, 't', 1.5],
      ['Blue Metal Road Base', 90, 't', 1.8], ['Cracker Dust / Rip Rap', 90, 't', 1.5]
    ],
    'Rock (per tonne)': [
      ['Granite Rock (various)', 320, 't', 1.5], ['Ironstone Rock', 80, 't', 1.5]
    ]
  };
}

/** Run this once after pasting, to check the sheet connection + permissions. */
function testSetup() {
  if (SHEET_ID === 'PASTE_YOUR_SHEET_ID_HERE') {
    throw new Error('Paste your Sheet ID into SHEET_ID at the top of this file first.');
  }
  const sh = sheet_();
  const jobs = getJobs();
  Logger.log('Connected to: ' + sh.getParent().getName());
  Logger.log('Sheet tab: ' + sh.getName());
  Logger.log('Jobs found: ' + jobs.length);
  Logger.log('SUCCESS — now deploy as a web app (Access: Only myself).');
  return 'OK';
}

/* ============ THE APP ============ */

const APP_HTML = `
<!DOCTYPE html><html><head><meta charset="utf-8">
<style>
:root{--bg:#1A1A1A;--org:#E8651A;--mid:#2E2E2E;--drk:#262626;--div:#3A3A3A;--sof:#CCCCCC}
*{margin:0;padding:0;box-sizing:border-box;-webkit-tap-highlight-color:transparent}
body{background:var(--bg);color:#fff;font:16px/1.45 -apple-system,system-ui,sans-serif;padding-bottom:80px}
header{background:var(--drk);border-bottom:3px solid var(--org);padding:12px 14px;position:sticky;top:0;z-index:10}
h1{font-size:1.1rem;letter-spacing:.05em;text-transform:uppercase}
h1 span{color:var(--org)}
.tabs{display:flex;background:var(--drk);border-bottom:1px solid var(--div);position:sticky;top:52px;z-index:9}
.tab{flex:1;padding:13px 4px;text-align:center;font-weight:600;font-size:.9rem;color:var(--sof);border-bottom:3px solid transparent}
.tab.on{color:var(--org);border-bottom-color:var(--org)}
.wrap{padding:12px}
.filters{display:flex;gap:6px;overflow-x:auto;padding:0 0 10px}
.chip{padding:7px 13px;border:1px solid var(--div);border-radius:20px;background:var(--drk);color:var(--sof);font-size:.82rem;white-space:nowrap}
.chip.on{background:var(--org);color:#1A1A1A;border-color:var(--org);font-weight:600}
.card{background:var(--drk);border:1px solid var(--div);border-left:4px solid var(--org);border-radius:5px;padding:12px;margin-bottom:9px}
.card .top{display:flex;justify-content:space-between;gap:8px;align-items:flex-start}
.card .nm{font-weight:600}
.card .mt{color:var(--sof);font-size:.9rem;margin-top:2px}
.card .dt{color:#8d8d8d;font-size:.76rem;margin-top:4px}
.badge{font-size:.68rem;text-transform:uppercase;letter-spacing:.06em;padding:3px 8px;border-radius:3px;font-weight:700;white-space:nowrap}
.b-New{background:var(--org);color:#1A1A1A}
.b-Quoted{background:#3d6fb5;color:#fff}
.b-Paid{background:#2f7d4f;color:#fff}
.b-Booked{background:#7a4fb5;color:#fff}
.b-Delivered{background:#4a4a4a;color:#ddd}
.b-Closed{background:#333;color:#888}
.detail{margin-top:11px;padding-top:11px;border-top:1px solid var(--div);display:none}
.detail.open{display:block}
.kv{font-size:.88rem;color:var(--sof);margin-bottom:4px}
.kv b{color:#fff;font-weight:600}
.acts{display:flex;gap:7px;margin-top:11px;flex-wrap:wrap}
.btn{flex:1;min-width:88px;padding:11px 8px;border:0;border-radius:4px;background:var(--mid);color:#fff;font:inherit;font-size:.85rem;font-weight:600;text-align:center;text-decoration:none}
.btn.org{background:var(--org);color:#1A1A1A}
.btn.gh{background:transparent;border:1px solid var(--div);color:var(--sof)}
.stat{display:flex;gap:6px;flex-wrap:wrap;margin-top:11px}
.sbtn{padding:8px 11px;border:1px solid var(--div);border-radius:4px;background:var(--drk);color:var(--sof);font-size:.8rem;font-weight:600}
.sbtn.on{background:var(--org);color:#1A1A1A;border-color:var(--org)}
label{display:block;font-size:.85rem;font-weight:600;margin:12px 0 4px}
input,select,textarea{width:100%;padding:11px;border:1px solid var(--div);border-radius:4px;background:var(--mid);color:#fff;font:inherit}
textarea{min-height:64px}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.row3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px}
.out{margin-top:14px;padding:14px;background:var(--drk);border:1px solid var(--div);border-left:4px solid var(--org);border-radius:4px;display:none}
.out.on{display:block}
.out .big{font-size:1.4rem;font-weight:700;color:var(--org)}
.out .ln{font-size:.9rem;color:var(--sof);margin-top:5px}
.out .tot{font-size:1.15rem;font-weight:700;margin-top:9px}
.msg{position:fixed;left:12px;right:12px;bottom:16px;background:var(--org);color:#1A1A1A;padding:12px;border-radius:5px;font-weight:600;text-align:center;display:none;z-index:50}
.empty{text-align:center;color:#8d8d8d;padding:40px 12px;font-size:.9rem}
.spin{text-align:center;color:var(--org);padding:30px;font-weight:600}
</style></head><body>

<header><h1>MRTH <span>Jobs</span></h1></header>
<div class="tabs">
  <div class="tab on" data-t="jobs">Jobs</div>
  <div class="tab" data-t="new">+ Enquiry</div>
  <div class="tab" data-t="quote">Quote</div>
  <div class="tab" data-t="settings">Settings</div>
</div>

<div class="wrap">

  <div id="v-jobs">
    <div class="filters" id="filters"></div>
    <div id="list"><div class="spin">Loading…</div></div>
  </div>

  <div id="v-new" style="display:none">
    <label>Name</label><input id="n-name" autocomplete="off">
    <label>Mobile</label><input id="n-phone" type="tel" inputmode="tel">
    <label>Type</label>
    <select id="n-type"><option>Materials delivery</option><option>Trailer hire</option></select>
    <label>Material</label><input id="n-material" placeholder="e.g. Pine bark mulch">
    <div class="row2">
      <div><label>Quantity</label><input id="n-qty" placeholder="4 m3"></div>
      <div><label>Suburb</label><input id="n-suburb" placeholder="Witchcliffe"></div>
    </div>
    <label>When</label>
    <select id="n-when"><option>This week</option><option>Next week</option><option>Within a month</option><option>Just pricing</option></select>
    <label>Access / notes</label><textarea id="n-notes"></textarea>
    <div class="acts"><button class="btn org" id="saveNew">Save enquiry</button></div>
  </div>

  <div id="v-quote" style="display:none">
    <label>Material</label><select id="q-mat"></select>
    <div class="row3">
      <div><label>Length m</label><input id="q-l" type="number" inputmode="decimal"></div>
      <div><label>Width m</label><input id="q-w" type="number" inputmode="decimal"></div>
      <div><label>Depth mm</label><input id="q-d" type="number" inputmode="numeric"></div>
    </div>
    <label>…or enter amount directly</label>
    <div class="row2">
      <input id="q-amt" type="number" inputmode="decimal" placeholder="amount">
      <select id="q-unit"><option value="m3">m3</option><option value="t">tonnes</option></select>
    </div>
    <label>Delivery fee to add ($)</label>
    <input id="q-del" type="number" inputmode="decimal" placeholder="e.g. 140">
    <div class="out" id="q-out"></div>
    <div class="acts" style="margin-top:12px">
      <button class="btn org" id="q-copy">Copy quote message</button>
    </div>
  </div>

  <div id="v-settings" style="display:none">
    <label>GST start date</label>
    <input id="s-gst" type="date">
    <div class="acts">
      <button class="btn org" id="s-gst-save">Save GST date</button>
      <button class="btn gh" id="s-gst-clear">Not registered</button>
    </div>
    <label>Next invoice number</label>
    <input id="s-seq" type="number" inputmode="numeric" min="1">
    <div class="acts"><button class="btn org" id="s-seq-save">Save number</button></div>
    <div class="kv" id="s-current" style="margin-top:16px;white-space:pre-line"></div>
  </div>

</div>
<div class="msg" id="toast"></div>

<script>
var JOBS = [], PRICES = {}, FILTER = 'Active', OPEN = null;
var STATUSES = ['New','Quoted','Paid','Booked','Delivered','Closed'];

function toast(t){var e=document.getElementById('toast');e.textContent=t;e.style.display='block';
  setTimeout(function(){e.style.display='none';},2200);}

/* tabs */
document.querySelectorAll('.tab').forEach(function(tb){
  tb.onclick=function(){
    document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on');});
    tb.classList.add('on');
    ['jobs','new','quote','settings'].forEach(function(v){
      document.getElementById('v-'+v).style.display = (v===tb.dataset.t)?'block':'none';
    });
    if(tb.dataset.t==='settings') loadSettings();
  };
});

/* ---------- JOBS ---------- */
function drawFilters(){
  var f=['Active'].concat(STATUSES).concat(['All']);
  document.getElementById('filters').innerHTML = f.map(function(s){
    return '<div class="chip'+(s===FILTER?' on':'')+'" data-f="'+s+'">'+s+'</div>';
  }).join('');
  document.querySelectorAll('.chip').forEach(function(c){
    c.onclick=function(){FILTER=c.dataset.f;drawFilters();drawList();};
  });
}

function visible(){
  if(FILTER==='All') return JOBS;
  if(FILTER==='Active') return JOBS.filter(function(j){
    return j.status!=='Delivered' && j.status!=='Closed';});
  return JOBS.filter(function(j){return j.status===FILTER;});
}

function esc(s){return String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}

function drawList(){
  var v=visible();
  if(!v.length){document.getElementById('list').innerHTML='<div class="empty">Nothing here yet.</div>';return;}
  document.getElementById('list').innerHTML = v.map(function(j){
    var open = (OPEN===j.row);
    var sms = j.phone ? 'sms:'+j.phone.replace(/\\s/g,'') : '';
    var tel = j.phone ? 'tel:'+j.phone.replace(/\\s/g,'') : '';
    var d = '<div class="detail'+(open?' open':'')+'" id="d'+j.row+'">'
      + (j.material?'<div class="kv">Material: <b>'+esc(j.material)+'</b></div>':'')
      + (j.quantity?'<div class="kv">Qty: <b>'+esc(j.quantity)+'</b></div>':'')
      + (j.loadType?'<div class="kv">Load type: <b>'+esc(j.loadType)+'</b></div>':'')
      + (j.suburb?'<div class="kv">Suburb: <b>'+esc(j.suburb)+'</b></div>':'')
      + (j.address?'<div class="kv">Address: <b>'+esc(j.address)+'</b></div>':'')
      + (j.timeframe?'<div class="kv">When: <b>'+esc(j.timeframe)+'</b></div>':'')
      + (j.access?'<div class="kv">Access: <b>'+esc(j.access)+'</b></div>':'')
      + (j.email?'<div class="kv">Email: <b>'+esc(j.email)+'</b></div>':'')
      + (j.quoted?'<div class="kv">Quoted: <b>$'+esc(j.quoted)+'</b></div>':'')
      + (j.carrier?'<div class="kv">Carrier: <b>'+esc(j.carrier)+'</b></div>':'')
      + (j.notes?'<div class="kv">Notes: <b>'+esc(j.notes)+'</b></div>':'')
      + (j.invoiceNo?'<div class="kv">Invoice: <b>'+esc(j.invoiceNo)+'</b>'+(j.invoiceLink?' — <a href="'+j.invoiceLink+'" target="_blank" style="color:inherit">PDF</a>':'')+'</div>':'')
      + (j.startDate?'<div class="kv">Hire dates: <b>'+esc(j.startDate)+(j.endDate?' to '+esc(j.endDate):'')+'</b></div>':'')
      + (j.scheduled?'<div class="kv">Scheduled: <b>'+esc(j.scheduled)+'</b></div>':'')
      + '<div class="acts">'
      + (tel?'<a class="btn" href="'+tel+'">Call</a>':'')
      + (sms?'<a class="btn" href="'+sms+'">Text</a>':'')
      + '<button class="btn gh" onclick="setQuoted('+j.row+')">Set $</button>'
      + '<button class="btn gh" onclick="setCarrier('+j.row+')">Carrier</button>'
      + '<button class="btn gh" onclick="setAddress('+j.row+')">Address</button>'
      + '<button class="btn gh" onclick="setJobDate('+j.row+')">Date</button>'
      + '<button class="btn gh" onclick="setNote('+j.row+')">Note</button>'
      + '<button class="btn org" onclick="genInvoice('+j.row+')">Invoice</button>'
      + '</div>'
      + '<div class="stat">' + STATUSES.map(function(s){
          return '<div class="sbtn'+(j.status===s?' on':'')+'" onclick="setStatus('+j.row+',\\''+s+'\\')">'+s+'</div>';
        }).join('') + '</div>'
      + '</div>';
    return '<div class="card" onclick="tog(event,'+j.row+')">'
      + '<div class="top"><div><div class="nm">'+esc(j.name||'(no name)')+'</div>'
      + '<div class="mt">'+esc([j.material,j.quantity,j.loadType,j.suburb].filter(Boolean).join(' · ')||j.type)+'</div>'
      + '<div class="dt">'+esc(j.received)+'</div></div>'
      + '<span class="badge b-'+j.status+'">'+j.status+'</span></div>'
      + d + '</div>';
  }).join('');
}

function tog(e,row){
  if(e.target.closest('.acts')||e.target.closest('.stat')) return;
  OPEN = (OPEN===row)?null:row; drawList();
}

function save(row,field,val,label){
  google.script.run.withSuccessHandler(function(){
    var j=JOBS.filter(function(x){return x.row===row;})[0];
    if(j){
      if(field==='Status')j.status=val;
      if(field==='Quoted $')j.quoted=val;
      if(field==='Carrier')j.carrier=val;
      if(field==='Notes')j.notes=val;
    }
    drawList(); toast(label||'Saved');
  }).withFailureHandler(function(e){toast('Failed: '+e.message);})
   .updateJob(row,field,val);
}
function setStatus(row,s){ save(row,'Status',s,s); }
function setQuoted(row){ var v=prompt('Quoted amount ($)'); if(v!==null) save(row,'Quoted $',v,'Quote saved'); }
function setCarrier(row){ var v=prompt('Carrier / driver'); if(v!==null) save(row,'Carrier',v,'Carrier saved'); }
function setAddress(row){
  var j=JOBS.filter(function(x){return x.row===row;})[0];
  var v=prompt('Delivery address', j?j.address:''); if(v!==null) save(row,'Delivery address',v,'Address saved');
}
function setNote(row){
  var j=JOBS.filter(function(x){return x.row===row;})[0];
  var v=prompt('Notes', j?j.notes:''); if(v!==null) save(row,'Notes',v,'Note saved');
}
function setJobDate(row){
  var j=JOBS.filter(function(x){return x.row===row;})[0];
  var isTrailer = j && j.type && j.type.toLowerCase().indexOf('trailer')!==-1;
  if(isTrailer){
    var s=prompt('Start date (YYYY-MM-DD)', j?j.startDate:''); if(s===null) return;
    save(row,'Start Date',s,'Start date saved');
    var e=prompt('End date (YYYY-MM-DD)', (j?j.endDate:'')||s); if(e===null) return;
    save(row,'End Date',e,'End date saved');
  } else {
    var d=prompt('Scheduled date (YYYY-MM-DD)', j?j.scheduled:''); if(d===null) return;
    save(row,'Scheduled date',d,'Scheduled date saved');
  }
}

/* ---------- INVOICE ---------- */
function genInvoice(row){
  google.script.run.withSuccessHandler(function(p){
    var lines=[];
    if(p.alreadyInvoiced) lines.push('Already invoiced as '+p.alreadyInvoiced+'. Generate another?','');
    lines.push(p.title+' '+p.invoiceNumber);
    lines.push('Customer: '+(p.customerName||'(no name on this job)'));
    lines.push('Supply date: '+p.supplyDateFormatted+' (from '+p.supplyDateSource+')');
    lines.push('Quoted $ field: "'+p.rawQuoted+'"');
    lines.push('Parsed amount: '+p.totalFormatted+(p.gst?' (incl. GST '+p.gstAmountFormatted+')':' — no GST'));
    lines.push('Payment link: '+(p.pricingMatch?'matched to '+p.pricingMatch:'generic (no trailer rate matched)'));
    if(p.warnings.length) lines.push('','Check:','- '+p.warnings.join('\\n- '));
    lines.push('','Generate this invoice?');
    if(!confirm(lines.join('\\n'))) return;
    toast('Generating…');
    google.script.run.withSuccessHandler(function(r){
      var msg=r.title+' '+r.invoiceNumber+' created for '+r.totalFormatted+'.\\nSaved to Drive.';
      if(r.customerEmail){
        if(confirm(msg+'\\n\\nCreate a Gmail draft to '+r.customerEmail+'?')){
          google.script.run.withSuccessHandler(function(){toast('Draft saved in Gmail');})
            .withFailureHandler(function(e){toast('Draft failed: '+e.message);})
            .createInvoiceEmailDraft(row);
          load();
          return;
        }
      }
      toast(msg); load();
    }).withFailureHandler(function(e){toast('Failed: '+e.message);}).confirmGenerateInvoice(row);
  }).withFailureHandler(function(e){toast('Failed: '+e.message);}).previewInvoice(row);
}

/* ---------- SETTINGS ---------- */
function loadSettings(){
  google.script.run.withSuccessHandler(function(s){
    document.getElementById('s-gst').value = s.gstStart || '';
    document.getElementById('s-seq').value = s.nextSeq;
    document.getElementById('s-current').textContent =
      'ABN '+s.abn+'\\nGST: '+(s.gstStart?'registered from '+s.gstStart:'not registered')+
      '\\nNext invoice: '+s.nextInvoiceNumber+'\\nEmail mode: '+s.emailMode+
      '\\nDrive folder: '+s.driveFolder;
  }).getInvoiceSettings();
}
document.getElementById('s-gst-save').onclick=function(){
  var v=document.getElementById('s-gst').value;
  if(!v){toast('Pick a date, or use "Not registered"');return;}
  google.script.run.withSuccessHandler(function(){toast('GST start date saved');loadSettings();})
    .withFailureHandler(function(e){toast('Failed: '+e.message);}).setGstStartDate(v);
};
document.getElementById('s-gst-clear').onclick=function(){
  if(!confirm('Clear GST registration? Invoices will go back to plain, no-GST.')) return;
  google.script.run.withSuccessHandler(function(){toast('GST cleared');loadSettings();})
    .withFailureHandler(function(e){toast('Failed: '+e.message);}).setGstStartDate('');
};
document.getElementById('s-seq-save').onclick=function(){
  var n=+document.getElementById('s-seq').value;
  if(!n||n<1){toast('Enter a whole number of 1 or more');return;}
  if(!confirm('Set next invoice number to '+n+'? Only do this once, when you start invoicing.')) return;
  google.script.run.withSuccessHandler(function(){toast('Saved');loadSettings();})
    .withFailureHandler(function(e){toast('Failed: '+e.message);}).setInvoiceCounter(n);
};

function load(){
  google.script.run.withSuccessHandler(function(d){
    JOBS=d; drawFilters(); drawList();
  }).withFailureHandler(function(e){
    document.getElementById('list').innerHTML='<div class="empty">Could not load: '+e.message+'</div>';
  }).getJobs();
}

/* ---------- NEW ENQUIRY ---------- */
document.getElementById('saveNew').onclick=function(){
  var o={
    name:document.getElementById('n-name').value,
    phone:document.getElementById('n-phone').value,
    type:document.getElementById('n-type').value,
    material:document.getElementById('n-material').value,
    quantity:document.getElementById('n-qty').value,
    suburb:document.getElementById('n-suburb').value,
    timeframe:document.getElementById('n-when').value,
    notes:document.getElementById('n-notes').value
  };
  if(!o.name && !o.phone){toast('Add a name or number');return;}
  this.textContent='Saving…'; this.disabled=true;
  var btn=this;
  google.script.run.withSuccessHandler(function(){
    ['n-name','n-phone','n-material','n-qty','n-suburb','n-notes'].forEach(function(id){
      document.getElementById(id).value='';});
    btn.textContent='Save enquiry'; btn.disabled=false;
    toast('Enquiry added'); load();
    document.querySelector('.tab[data-t="jobs"]').click();
  }).withFailureHandler(function(e){
    btn.textContent='Save enquiry'; btn.disabled=false; toast('Failed: '+e.message);
  }).addJob(o);
};

/* ---------- QUOTE ---------- */
function buildMats(){
  google.script.run.withSuccessHandler(function(p){
    PRICES=p; var sel=document.getElementById('q-mat'); sel.innerHTML='';
    Object.keys(p).forEach(function(g){
      var og=document.createElement('optgroup'); og.label=g;
      p[g].forEach(function(m){
        var o=document.createElement('option');
        o.value=m[0]; o.textContent=m[0]+' — $'+m[1]+'/'+(m[2]==='t'?'t':'m3');
        o.dataset.price=m[1]; o.dataset.unit=m[2]; o.dataset.dens=m[3];
        og.appendChild(o);
      });
      sel.appendChild(og);
    });
    quote();
  }).getPrices();
}

function quote(){
  var sel=document.getElementById('q-mat'); var o=sel.options[sel.selectedIndex];
  if(!o) return;
  var price=+o.dataset.price, unit=o.dataset.unit, dens=+o.dataset.dens;
  var L=+document.getElementById('q-l').value, W=+document.getElementById('q-w').value,
      D=+document.getElementById('q-d').value, amt=+document.getElementById('q-amt').value,
      au=document.getElementById('q-unit').value, del=+document.getElementById('q-del').value;
  var m3=0,t=0;
  if(amt>0){ if(au==='m3'){m3=amt; t=dens?amt*dens:0;} else {t=amt; m3=dens?amt/dens:amt;} }
  else if(L>0&&W>0&&D>0){ m3=L*W*(D/1000); t=dens?m3*dens:0; }
  var out=document.getElementById('q-out');
  if(m3<=0){ out.classList.remove('on'); return; }
  var bill = (unit==='t')? t : m3;
  var mat = bill*price;
  var tot = mat + (del>0?del:0);
  var qtyTxt = (unit==='t')? (t.toFixed(1)+' tonnes') : (m3.toFixed(1)+' m3');
  out.innerHTML = '<div class="big">'+m3.toFixed(1)+' m3'+(dens?' ≈ '+t.toFixed(1)+' t':'')+'</div>'
    + '<div class="ln">'+o.value+'</div>'
    + '<div class="ln">Materials: $'+Math.round(mat)+'</div>'
    + (del>0?'<div class="ln">Delivery: $'+Math.round(del)+'</div>':'<div class="ln">Delivery: not added yet</div>')
    + '<div class="tot">Total: $'+Math.round(tot)+'</div>'
    + (m3>5.5?'<div class="ln">⚠ Over a trailer load — truck job.</div>':'')
    + (unit==='t'&&t<1?'<div class="ln">⚠ Under 1 t minimum.</div>':'');
  out.classList.add('on');
  out.dataset.msg = 'All up that\\'s $'+Math.round(tot)+' delivered — '+qtyTxt+' of '
    + o.value + ', tipped where you want it.\\n\\nTo lock it in, payment first please: https://buy.stripe.com/6oUbIUdbo9APbiobRW7AI0B\\n'
    + 'Just enter $'+Math.round(tot)+' when it asks for the amount 👍\\n\\n'
    + 'Once that\\'s through I\\'ll confirm your delivery window.';
}
['q-mat','q-l','q-w','q-d','q-amt','q-unit','q-del'].forEach(function(id){
  document.getElementById(id).addEventListener('input',quote);
  document.getElementById(id).addEventListener('change',quote);
});
document.getElementById('q-copy').onclick=function(){
  var m=document.getElementById('q-out').dataset.msg;
  if(!m){toast('Enter a quantity first');return;}
  var ta=document.createElement('textarea'); ta.value=m; document.body.appendChild(ta);
  ta.select(); try{document.execCommand('copy');toast('Quote copied');}catch(e){toast('Copy failed');}
  document.body.removeChild(ta);
};

load(); buildMats();
</script>
</body></html>
`;