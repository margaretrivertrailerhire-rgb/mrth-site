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
  'Scheduled date','Start Date','End Date','Calendar Event ID',
  'Source','Customer type','Lane','Qty','Unit','Cost ex GST','Sell ex GST',
  'Margin ex GST','Lost reason','Closed date','Month',
  'Supplier cost','Payment method'
];

const STATUSES = ['New','Quoted','Paid','Booked','Delivered','Closed','Declined','Lost'];
const SOURCES = ['FB ad','FB group post','Marketplace','Google','Website','Repeat customer','Referral','Word of mouth','Other'];
const CUSTOMER_TYPES = ['Trade','Repeat','One-off'];
const LANES = ['Trailer single','Trailer double','Truck 10t (NCJ)','Truck 10t (other carrier)','Truck 20t+'];
const UNITS = ['t','m3'];
const LOST_REASONS = ['Price','No reply','Away on swing','No carrier','Other'];
const PAYMENT_METHODS = ['Stripe','PayID','Cash','Other'];

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
      endDate: fmtDay(get('End Date')),
      source: get('Source'),
      custType: get('Customer type'),
      lane: get('Lane'),
      pqty: get('Qty'),
      unit: get('Unit'),
      cost: get('Cost ex GST'),
      sell: get('Sell ex GST'),
      margin: get('Margin ex GST'),
      lostReason: get('Lost reason'),
      closedDate: fmtDay(get('Closed date')),
      supplierCost: get('Supplier cost'),
      payMethod: get('Payment method'),
      paidDate: fmtDay(get('Paid date'))
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
  if (field === 'Qty' || field === 'Unit' || field === 'Type') inferLaneIfBlank_(sh, map, row);
  syncJobFinance_(row);
  syncJobCalendarEvent_(row);
  return true;
}

/**
 * Closes a job as Declined (turned away) or Lost (quoted, then lost) —
 * one compound action: sets Status, stamps Closed date, records the
 * reason (defaults to 'Other' if the picker was skipped). Everything on
 * the row — quoted amount, qty, lane, cost — is left exactly as it was;
 * only Status/Lost reason/Closed date change. A tentative calendar event
 * (if any) is removed the same way any other non-Booked status removes
 * one, via syncJobCalendarEvent_.
 */
function closeJob(row, status, reason) {
  const sh = sheet_();
  const map = headerMap_(sh);
  if (map['Status'] === undefined) throw new Error('Missing Status column.');
  sh.getRange(row, map['Status'] + 1).setValue(status);
  if (map['Lost reason'] !== undefined) sh.getRange(row, map['Lost reason'] + 1).setValue(reason || 'Other');
  if (map['Closed date'] !== undefined) sh.getRange(row, map['Closed date'] + 1).setValue(new Date());
  syncJobFinance_(row);
  syncJobCalendarEvent_(row);
  return true;
}

/** Defaults a blank Lane to Trailer single once a job has some real
 *  content (a Type or a Qty) — never overwrites a Lane that's already
 *  set. Truck lanes are never guessed; they're only ever written
 *  explicitly by the truck-load tab's save, for its six NCJ products. */
function inferLaneIfBlank_(sh, map, row) {
  if (map['Lane'] === undefined) return;
  const laneCell = sh.getRange(row, map['Lane'] + 1);
  if (laneCell.getValue()) return;

  const get = function (field) {
    if (map[field] === undefined) return '';
    const v = sh.getRange(row, map[field] + 1).getValue();
    return v === null || v === undefined ? '' : v;
  };
  const qty = parseFloat(get('Qty')) || 0;
  if (qty <= 0 && !get('Type')) return;
  // Truck lanes are only ever set explicitly, by the truck-load tab's
  // save (six NCJ products only) — every other material, whatever the
  // tonnage, defaults to a trailer lane. Tonnage alone is not a signal
  // that NCJ is involved.
  laneCell.setValue('Trailer single');
}

/**
 * Sell ex GST and Margin ex GST for one row, recalculated on every
 * relevant edit so both stay correct across the 1 Oct 2026 GST
 * changeover — including for a job whose Paid/Delivered date is set
 * later, well after it was first quoted. Reuses the exact same
 * gstAppliesOn()/resolveSupplyDate_() the invoice generator uses, so
 * "is this job GST-inclusive" only has one answer anywhere in the app.
 * Sell is written as a plain value (it depends on a Script Property,
 * which a sheet formula can't read); Margin is a real formula, written
 * once, that keeps recalculating itself against Sell/Cost on the row.
 */
function syncJobFinance_(row) {
  const sh = sheet_();
  const map = headerMap_(sh);
  if (map['Sell ex GST'] === undefined) return;

  const ctx = readJobRow_(row);
  const rawQuoted = ctx.get('Quoted $');
  const total = parseAmount_(rawQuoted);
  const supply = resolveSupplyDate_(ctx);
  const gst = gstAppliesOn(supply.date);
  const sell = total > 0 ? (gst ? Math.round((total / 1.1) * 100) / 100 : total) : '';
  sh.getRange(row, map['Sell ex GST'] + 1).setValue(sell);

  if (map['Margin ex GST'] !== undefined) {
    const sellCol = columnLetter_(map['Sell ex GST'] + 1);
    const costCol = columnLetter_(map['Cost ex GST'] + 1);
    const formula = '=IF(OR(' + sellCol + row + '="",' + costCol + row + '=""),"",' + sellCol + row + '-' + costCol + row + ')';
    const cell = sh.getRange(row, map['Margin ex GST'] + 1);
    if (cell.getFormula() !== formula) cell.setFormula(formula);
  }
  if (map['Month'] !== undefined && map['Received'] !== undefined) {
    const recCol = columnLetter_(map['Received'] + 1);
    const formula = '=IF(' + recCol + row + '="","",TEXT(EOMONTH(' + recCol + row + ',-1)+1,"mmm-yy"))';
    const cell = sh.getRange(row, map['Month'] + 1);
    if (cell.getFormula() !== formula) cell.setFormula(formula);
  }
}

function columnLetter_(col) {
  let s = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
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
  set('Source', obj.source || '');
  set('Customer type', obj.custType || '');
  sh.appendRow(row);
  syncJobFinance_(sh.getLastRow());
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

/** Flat, all-in 10 t truck-load prices — material + delivery, no fuel
 *  levy, no extras. Mirrors the "Full truck loads" table on the
 *  homepage. Each material row is [name, MR & south, Cowaramup,
 *  Gracetown]; anything outside those three zones is quoted manually. */
function getTruckLoadPrices() {
  return {
    zones: ['MR & south', 'Cowaramup', 'Gracetown'],
    materials: [
      ['Screened yellow / white sand', 750, 810, 860],
      ['Screened topsoil', 800, 860, 910],
      ['A-grade gravel', 850, 910, 960],
      ['Crushed 50mm red stone', 875, 935, 985],
      ['Limestone road base', 900, 960, 1010],
      ['Main Roads spec gravel', 925, 985, 1035]
    ]
  };
}

/* ============ PIPELINE: cost lookup ============ */

const PRICING_SHEET_NAME = 'Pricing';
const COST_TABLE_COL = 6; // column F — a second table on the Pricing tab: Product | Lane | Cost ex GST | Cost basis

/** A few getPrices() material names map cleanly onto the seeded cost
 *  figures under a different label. This is the only guessing this does
 *  — anything not listed here needs its own cost row added by name in
 *  the sheet (see setupPricingCosts). */
const PRODUCT_COST_ALIASES = {
  'bushland': 'bushland mulch',
  'mrd gravel road base': 'main roads gravel',
  'mrd limestone road base': 'limestone road base',
  'screened yellow / white sand': 'white sand', // same cost either colour (both 38.50/t)
  'main roads spec gravel': 'main roads gravel' // truck-load tab's name for this NCJ product
};

function costTableSheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(PRICING_SHEET_NAME);
}

/** Last row actually used by the F:I cost table specifically — the A:D
 *  trailer-rate table on the same tab has its own, unrelated height. */
function costTableLastRow_(sh) {
  const col = sh.getRange(1, COST_TABLE_COL, sh.getMaxRows(), 1).getValues();
  let last = 0;
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() !== '') last = i + 1;
  }
  return last;
}

function costRows_() {
  const sh = costTableSheet_();
  if (!sh) return [];
  const last = costTableLastRow_(sh);
  if (last < 2) return [];
  return sh.getRange(2, COST_TABLE_COL, last - 1, 4).getValues()
    .map(function (r) {
      return { product: String(r[0] || '').trim(), lane: String(r[1] || '').trim(), cost: r[2], basis: String(r[3] || '').trim() };
    })
    .filter(function (r) { return r.product; });
}

/** Looks up Cost ex GST for a product + lane: exact product name, then
 *  the small alias list above, then falls back to the product's only
 *  cost row if it has just one (lane doesn't matter). Returns null (not
 *  0) when nothing is found, so a missing cost can be flagged rather
 *  than silently costed at zero. */
function lookupCost_(product, lane) {
  const key = String(product || '').trim().toLowerCase();
  if (!key) return null;
  const wantedKey = PRODUCT_COST_ALIASES[key] || key;
  const matches = costRows_().filter(function (r) { return r.product.toLowerCase() === wantedKey; });
  if (!matches.length) return null;
  const laneMatch = matches.filter(function (r) { return r.lane.toLowerCase() === String(lane || '').toLowerCase(); });
  const row = laneMatch.length ? laneMatch[0] : matches[0];
  const n = typeof row.cost === 'number' ? row.cost : parseFloat(row.cost);
  return isNaN(n) ? null : n;
}

/** One-time setup: adds the Product/Lane/Cost ex GST/Cost basis table to
 *  the Pricing tab (columns F:I, leaving the existing trailer-rate table
 *  in A:D untouched) and seeds it with the known NCJ/Cowara figures.
 *  Safe to re-run — only adds rows that aren't already there. */
function setupPricingCosts() {
  const sh = costTableSheet_();
  if (!sh) throw new Error('No "' + PRICING_SHEET_NAME + '" tab found.');

  const headers = ['Product', 'Lane', 'Cost ex GST', 'Cost basis'];
  const headerRange = sh.getRange(1, COST_TABLE_COL, 1, 4);
  if (headerRange.getValues()[0].join('|') !== headers.join('|')) {
    headerRange.setValues([headers]).setFontWeight('bold');
  }

  const NCJ = 'NCJ Adamson, delivered, MR-Augusta zone (Cowaramup +5.00/t, Gracetown +10.00/t)';
  const seed = [
    ['White sand', 'Truck 10t (NCJ)', 38.50, NCJ],
    ['Yellow sand', 'Truck 10t (NCJ)', 38.50, NCJ],
    ['Screened topsoil', 'Truck 10t (NCJ)', 42.00, NCJ],
    ['A-grade gravel', 'Truck 10t (NCJ)', 41.80, NCJ],
    ['Main Roads gravel', 'Truck 10t (NCJ)', 47.30, NCJ],
    ['Limestone road base', 'Truck 10t (NCJ)', 49.50, NCJ],
    ['Crushed 50mm red stone', 'Truck 10t (NCJ)', 46.20, NCJ],
    ['Quartz grit', 'Truck 10t (NCJ)', 55.00, NCJ],
    ['Dam fill', 'Truck 10t (NCJ)', 33.60, NCJ],
    ['Rip rap / laterite', 'Truck 10t (NCJ)', 88.00, NCJ],
    ['Granite', 'Truck 10t (NCJ)', 148.50, NCJ],
    ['Bushland mulch', 'Trailer single', 120.00, 'Cowara yard + per-trip delivery'],
    ['Trailer hire', 'Trailer single', 0, 'Trailer hire - nil']
  ];

  const existing = costRows_();
  const already = function (product, lane) {
    return existing.some(function (r) {
      return r.product.toLowerCase() === product.toLowerCase() && r.lane.toLowerCase() === lane.toLowerCase();
    });
  };
  const toAdd = seed.filter(function (r) { return !already(r[0], r[1]); });
  if (toAdd.length) {
    sh.getRange(costTableLastRow_(sh) + 1, COST_TABLE_COL, toAdd.length, 4).setValues(toAdd);
  }
  Logger.log('Cost table ready: added ' + toAdd.length + ' row(s), ' + (seed.length - toAdd.length) + ' already present.');
  return 'OK';
}

/** Called when the quoter is bound to a job ("Save to job"): writes
 *  Quoted $, Qty, Unit and Lane, looks up Cost ex GST for the product +
 *  lane and writes that too (scaled by qty). Sell/Margin then follow on
 *  from syncJobFinance_. Returns whether a cost was actually found, so
 *  the card can show a "no cost set" flag when it wasn't. */
function saveQuoteToJob(row, payload) {
  const sh = sheet_();
  const map = headerMap_(sh);
  const set = function (field, value) {
    if (map[field] !== undefined && value !== undefined && value !== null && value !== '') {
      sh.getRange(row, map[field] + 1).setValue(value);
    }
  };
  set('Quoted $', payload.total);
  set('Qty', payload.qty);
  set('Unit', payload.unit);
  if (payload.lane) set('Lane', payload.lane);

  const cost = lookupCost_(payload.product, payload.lane);
  const costFound = cost !== null;
  if (costFound) set('Cost ex GST', Math.round(cost * payload.qty * 100) / 100);

  inferLaneIfBlank_(sh, map, row);
  syncJobFinance_(row);
  return { costFound: costFound };
}

/* ============ PIPELINE: setup ============ */

function columnLetter_(col) {
  let s = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

/** One-time setup: dropdowns for every Pipeline field on Enquiries.
 *  Safe to re-run any time. */
function setupPipelineValidation() {
  const sh = sheet_();
  const map = headerMap_(sh);
  const maxRows = Math.max(sh.getMaxRows(), 500);

  const applyList = function (field, list) {
    if (map[field] === undefined) return;
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build();
    sh.getRange(2, map[field] + 1, maxRows - 1, 1).setDataValidation(rule);
  };

  applyList('Status', STATUSES);
  applyList('Source', SOURCES);
  applyList('Customer type', CUSTOMER_TYPES);
  applyList('Lane', LANES);
  applyList('Unit', UNITS);
  applyList('Lost reason', LOST_REASONS);
  applyList('Payment method', PAYMENT_METHODS);

  Logger.log('Pipeline dropdowns set up on Enquiries.');
  return 'OK';
}

/** One-time (or re-run any time) backfill: writes the Month/Margin
 *  formulas and current Sell ex GST onto every existing row, for jobs
 *  that were on the sheet before this feature existed. */
function backfillPipelineFormulas() {
  const sh = sheet_();
  const last = sh.getLastRow();
  for (let r = 2; r <= last; r++) syncJobFinance_(r);
  Logger.log('Backfilled ' + (last - 1) + ' row(s).');
  return 'OK';
}

/**
 * One-time setup: builds the Pipeline tab — Sep-26 through Aug-27 across
 * the columns, one row per metric, all COUNTIFS/SUMIFS/AVERAGEIFS-style
 * formulas against Enquiries (plus one SUMPRODUCT for the avg-days
 * metric, which needs a conditional average of a date difference that
 * those functions can't express). Column positions for each Enquiries
 * field are looked up once here and baked into the formulas, so this
 * must be re-run if Enquiries columns are ever deleted or reordered
 * (appending new columns, which is all sheet_() ever does, is fine).
 * Re-running clears and rebuilds the tab from scratch.
 */
function setupPipelineTab() {
  const enq = sheet_();
  const emap = headerMap_(enq);
  const need = ['Month', 'Status', 'Lane', 'Unit', 'Qty', 'Margin ex GST', 'Lost reason', 'Customer type', 'Received', 'Delivered date'];
  need.forEach(function (f) {
    if (emap[f] === undefined) throw new Error('Enquiries is missing "' + f + '" — open the app once first so it creates the Pipeline columns.');
  });
  const col = {};
  need.forEach(function (f) { col[f] = columnLetter_(emap[f] + 1); });
  const ref = function (f) { return 'Enquiries!$' + col[f] + '$2:$' + col[f] + '$5000'; };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName('Pipeline');
  if (!sh) sh = ss.insertSheet('Pipeline'); else sh.clear();

  const months = [];
  const start = new Date(2026, 8, 1); // Sep 2026
  for (let i = 0; i < 12; i++) {
    const d = new Date(start.getFullYear(), start.getMonth() + i, 1);
    months.push(Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM-yy'));
  }

  const labels = [
    'Metric', 'Enquiries (all)', 'Truck loads delivered', 'Trailer deliveries',
    'Tonnes delivered', 'Margin ex GST — delivered', 'Avg margin per truck load',
    'Declined/Lost — Away on swing', 'Declined/Lost — No carrier + Other',
    'Lost on price / no reply', 'Margin lost to swing (est.)',
    'Repeat/trade share of deliveries', 'Avg days enquiry to delivery'
  ];
  sh.getRange(1, 1, labels.length, 1).setValues(labels.map(function (l) { return [l]; }));
  sh.getRange(1, 2, 1, months.length).setValues([months]);
  sh.getRange(1, 1, 1, months.length + 1).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  for (let i = 0; i < months.length; i++) {
    const c = columnLetter_(2 + i);
    const m = c + '$1';

    sh.getRange(c + '2').setFormula('=COUNTIFS(' + ref('Month') + ',' + m + ')');

    sh.getRange(c + '3').setFormula(
      '=COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Delivered",' + ref('Lane') + ',"Truck*")' +
      '+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Paid",' + ref('Lane') + ',"Truck*")'
    );

    sh.getRange(c + '4').setFormula(
      '=COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Delivered",' + ref('Lane') + ',"Trailer*")' +
      '+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Paid",' + ref('Lane') + ',"Trailer*")'
    );

    sh.getRange(c + '5').setFormula(
      '=SUMIFS(' + ref('Qty') + ',' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Delivered",' + ref('Unit') + ',"t")' +
      '+SUMIFS(' + ref('Qty') + ',' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Paid",' + ref('Unit') + ',"t")'
    );

    sh.getRange(c + '6').setFormula(
      '=SUMIFS(' + ref('Margin ex GST') + ',' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Delivered")' +
      '+SUMIFS(' + ref('Margin ex GST') + ',' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Paid")'
    );

    sh.getRange(c + '7').setFormula(
      '=IFERROR((SUMIFS(' + ref('Margin ex GST') + ',' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Delivered",' + ref('Lane') + ',"Truck*")' +
      '+SUMIFS(' + ref('Margin ex GST') + ',' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Paid",' + ref('Lane') + ',"Truck*"))/' + c + '3,"")'
    );

    sh.getRange(c + '8').setFormula('=COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Lost reason') + ',"Away on swing")');

    sh.getRange(c + '9').setFormula(
      '=COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Lost reason') + ',"No carrier")' +
      '+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Lost reason') + ',"Other")'
    );

    sh.getRange(c + '10').setFormula(
      '=COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Lost reason') + ',"Price")' +
      '+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Lost reason') + ',"No reply")'
    );

    sh.getRange(c + '11').setFormula('=' + c + '8*' + c + '7');

    sh.getRange(c + '12').setFormula(
      '=IFERROR((COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Customer type') + ',"Trade",' + ref('Status') + ',"Delivered")' +
      '+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Customer type') + ',"Trade",' + ref('Status') + ',"Paid")' +
      '+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Customer type') + ',"Repeat",' + ref('Status') + ',"Delivered")' +
      '+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Customer type') + ',"Repeat",' + ref('Status') + ',"Paid"))/' +
      '(COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Delivered")+COUNTIFS(' + ref('Month') + ',' + m + ',' + ref('Status') + ',"Paid")),"")'
    );
    sh.getRange(c + '12').setNumberFormat('0%');

    sh.getRange(c + '13').setFormula(
      '=IFERROR(SUMPRODUCT((' + ref('Month') + '=' + m + ')*(' + ref('Status') + '="Delivered")*(' + ref('Delivered date') + '<>"")*(' + ref('Delivered date') + '-' + ref('Received') + '))' +
      '/SUMPRODUCT((' + ref('Month') + '=' + m + ')*(' + ref('Status') + '="Delivered")*(' + ref('Delivered date') + '<>"")),"")'
    );
  }

  sh.getRange('A15').setValue('Target truck loads/month');
  sh.getRange('B15').setValue(10);
  sh.getRange('A16').setValue('Months hitting target');
  sh.getRange('B16').setFormula('=COUNTIF(B3:M3,">="&B15)');
  sh.getRange('A17').setValue('Total loads lost to swing');
  sh.getRange('B17').setFormula('=SUM(B8:M8)');
  sh.getRange('A18').setValue('Total margin to date');
  sh.getRange('B18').setFormula('=SUM(B6:M6)');

  sh.autoResizeColumns(1, 1);
  Logger.log('Pipeline tab created/refreshed.');
  return 'OK';
}

/* ============ MONEY ============ */

/** Australian financial year start (1 July) for a given date. */
function fyStart_(d) {
  const y = d.getMonth() >= 6 ? d.getFullYear() : d.getFullYear() - 1;
  return new Date(y, 6, 1);
}

function asDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Money view: what actually landed, what's still owed, and what it cost.
 *
 * Everything here parses through the same parseAmount_() the invoice
 * generator uses and resolves GST through the same gstAppliesOn() /
 * resolveSupplyDate_() pair, so a figure on this tab and the figure on
 * the matching invoice can't drift apart.
 *
 * monthKey is 'YYYY-MM'; omit it for the current month. The FY-to-date
 * block always runs 1 July → today regardless of the month selected —
 * "year to date" means today, not the end of whichever month you're
 * looking at.
 */
function getMoneySummary(monthKey) {
  const sh = sheet_();
  const map = headerMap_(sh);
  const last = sh.getLastRow();

  const now = new Date();
  if (!monthKey) monthKey = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  const parts = String(monthKey).split('-');
  const mStart = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  const mEnd = new Date(mStart.getFullYear(), mStart.getMonth() + 1, 1);
  const fy = fyStart_(now);

  const blank = function () {
    return { paid: 0, gst: 0, supplierCost: 0, jobs: 0, byType: {}, byMethod: {} };
  };
  const month = blank(), fytd = blank();
  const unpaid = [];
  const months = {};

  if (last >= 2) {
    const rows = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      // Shim matching readJobRow_'s shape so the shared invoice helpers work
      // off the in-memory row rather than a round trip per job.
      const get = function (field) {
        const idx = map[field];
        if (idx === undefined) return '';
        const v = r[idx];
        return v === null || v === undefined ? '' : v;
      };
      const ctx = { get: get };

      const amount = parseAmount_(get('Quoted $'));
      const paidDate = asDate_(get('Paid date'));
      const invoiceNo = String(get('Invoice No') || '').trim();
      const name = String(get('Name') || '').trim();

      if (invoiceNo && !paidDate) {
        unpaid.push({
          row: i + 2,
          name: name || '(no name)',
          invoiceNo: invoiceNo,
          amount: amount,
          amountText: money_(amount),
          type: String(get('Type') || '').trim()
        });
      }

      if (!paidDate) continue;

      months[Utilities.formatDate(paidDate, Session.getScriptTimeZone(), 'yyyy-MM')] = true;

      const gstOnThis = gstAppliesOn(resolveSupplyDate_(ctx).date)
        ? Math.round((amount / 11) * 100) / 100 : 0;
      const cost = parseAmount_(get('Supplier cost'));
      const type = String(get('Type') || '').trim() || 'Unspecified';
      const method = String(get('Payment method') || '').trim() || 'Unrecorded';

      const add = function (b) {
        b.paid += amount;
        b.gst += gstOnThis;
        b.supplierCost += cost;
        b.jobs += 1;
        b.byType[type] = (b.byType[type] || 0) + amount;
        b.byMethod[method] = (b.byMethod[method] || 0) + amount;
      };
      if (paidDate >= mStart && paidDate < mEnd) add(month);
      if (paidDate >= fy && paidDate <= now) add(fytd);
    }
  }

  const finish = function (b) {
    const exGst = b.paid - b.gst;
    return {
      paid: b.paid, paidText: money_(b.paid),
      gst: b.gst, gstText: money_(b.gst),
      supplierCost: b.supplierCost, supplierCostText: money_(b.supplierCost),
      margin: exGst - b.supplierCost, marginText: money_(exGst - b.supplierCost),
      exGst: exGst, exGstText: money_(exGst),
      jobs: b.jobs,
      byType: b.byType, byMethod: b.byMethod
    };
  };

  unpaid.sort(function (a, b) { return b.amount - a.amount; });
  const unpaidTotal = unpaid.reduce(function (s, u) { return s + u.amount; }, 0);

  const monthList = Object.keys(months);
  if (monthList.indexOf(monthKey) === -1) monthList.push(monthKey);
  monthList.sort().reverse();

  return {
    monthKey: monthKey,
    monthLabel: Utilities.formatDate(mStart, Session.getScriptTimeZone(), 'MMMM yyyy'),
    fyLabel: 'FY to date (from ' + Utilities.formatDate(fy, Session.getScriptTimeZone(), 'd MMM yyyy') + ')',
    months: monthList,
    month: finish(month),
    fytd: finish(fytd),
    unpaid: unpaid,
    unpaidTotal: unpaidTotal,
    unpaidTotalText: money_(unpaidTotal)
  };
}

/** Current month's headline numbers off the Pipeline tab, for the
 *  one-line summary on the Jobs screen. Null if the Pipeline tab hasn't
 *  been set up yet. */
function getPipelineSummary() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName('Pipeline');
  if (!sh) return null;

  const monthLabel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM-yy');
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const colIdx = headers.indexOf(monthLabel);
  if (colIdx === -1) return null;

  const col = colIdx + 1;
  return {
    label: monthLabel.split('-')[0],
    truckLoads: sh.getRange(3, col).getValue() || 0,
    margin: sh.getRange(6, col).getValue() || 0,
    swing: sh.getRange(8, col).getValue() || 0
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
.tab{flex:1 1 0;min-width:0;padding:13px 2px;text-align:center;font-weight:600;font-size:.82rem;color:var(--sof);border-bottom:3px solid transparent;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
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
.warn{color:var(--org)}
.picker{position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:60;display:none;align-items:flex-end;justify-content:center}
.picker.on{display:flex}
.picker-sheet{background:var(--drk);border-top:3px solid var(--org);border-radius:12px 12px 0 0;padding:16px;width:100%;max-width:480px;max-height:70vh;overflow-y:auto}
.picker-title{font-weight:700;margin-bottom:10px;font-size:1rem}
.picker-opts{display:flex;flex-direction:column;gap:8px}
.qbind{display:none;background:var(--mid);padding:10px;border-radius:4px;margin-bottom:12px;align-items:center;justify-content:space-between;gap:8px}
.qbind.on{display:flex}
.pline{padding:0 0 10px;color:var(--sof);font-size:.88rem}
.pline b{color:#fff}
.mny-grid{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:10px}
.mny-cell{background:var(--drk);border:1px solid var(--div);border-radius:5px;padding:12px}
.mny-cell.hero{border-left:4px solid var(--org)}
.mny-lbl{font-size:.7rem;text-transform:uppercase;letter-spacing:.06em;color:#8d8d8d;margin-bottom:4px}
.mny-big{font-size:1.5rem;font-weight:700;color:#fff;line-height:1.1}
.mny-mid{font-size:1.05rem;font-weight:700;color:#fff}
.mny-sub{font-size:.75rem;color:var(--sof);margin-top:3px}
.mny-h{font-size:.78rem;text-transform:uppercase;letter-spacing:.07em;color:var(--org);font-weight:700;margin:16px 0 7px}
.mny-row{display:flex;justify-content:space-between;gap:10px;padding:7px 0;border-bottom:1px solid var(--div);font-size:.88rem;color:var(--sof)}
.mny-row b{color:#fff;font-weight:600;white-space:nowrap}
.mny-neg{color:#e06b5a}
</style></head><body>

<header><h1>MRTH <span>Jobs</span></h1></header>
<div class="tabs">
  <div class="tab on" data-t="jobs">Jobs</div>
  <div class="tab" data-t="new">+ Enquiry</div>
  <div class="tab" data-t="quote">Quote</div>
  <div class="tab" data-t="money">Money</div>
  <div class="tab" data-t="settings">Settings</div>
</div>

<div class="wrap">

  <div id="v-jobs">
    <div class="pline" id="pline"></div>
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
    <label>Source</label>
    <select id="n-source"><option value="">—</option><option>FB ad</option><option>FB group post</option><option>Marketplace</option><option>Google</option><option>Website</option><option>Repeat customer</option><option>Referral</option><option>Word of mouth</option><option>Other</option></select>
    <label>Customer type</label>
    <select id="n-custtype"><option value="">—</option><option>Trade</option><option>Repeat</option><option>One-off</option></select>
    <div class="acts"><button class="btn org" id="saveNew">Save enquiry</button></div>
  </div>

  <div id="v-quote" style="display:none">
    <div class="qbind" id="q-bind">
      <span id="q-bind-txt"></span>
      <button class="btn gh" style="flex:none;padding:6px 10px" onclick="clearQuoteBinding()">Clear</button>
    </div>
    <div class="filters">
      <div class="chip on" id="qm-load">Per load</div>
      <div class="chip" id="qm-truck">Truck load (10t)</div>
    </div>

    <div id="q-load-mode">
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
        <button class="btn" id="q-save" style="display:none" onclick="saveLoadQuoteToJob()">Save to job</button>
      </div>
    </div>

    <div id="q-truck-mode" style="display:none">
      <label>Material</label><select id="qt-mat"></select>
      <label>Zone</label><select id="qt-zone"></select>
      <div class="out" id="qt-out"></div>
      <div class="acts" style="margin-top:12px">
        <button class="btn org" id="qt-copy">Copy quote message</button>
        <button class="btn" id="qt-save" style="display:none" onclick="saveTruckQuoteToJob()">Save to job</button>
      </div>
    </div>
  </div>

  <div class="picker" id="picker">
    <div class="picker-sheet">
      <div class="picker-title" id="pk-title"></div>
      <div class="picker-opts" id="pk-opts"></div>
      <button class="btn" id="pk-skip" onclick="pickerSkip()" style="display:none;margin-top:8px">Skip (use Other)</button>
    </div>
  </div>

  <div id="v-money" style="display:none">
    <select id="mn-month" style="margin-bottom:12px"></select>
    <div id="mn-body"><div class="spin">Loading…</div></div>
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
var STATUSES = ['New','Quoted','Paid','Booked','Delivered','Closed','Declined','Lost'];
var SOURCES = ['FB ad','FB group post','Marketplace','Google','Website','Repeat customer','Referral','Word of mouth','Other'];
var CUSTOMER_TYPES = ['Trade','Repeat','One-off'];
var LANES = ['Trailer single','Trailer double','Truck 10t (NCJ)','Truck 10t (other carrier)','Truck 20t+'];
var UNITS = ['t','m3'];
var LOST_REASONS = ['Price','No reply','Away on swing','No carrier','Other'];
var PAYMENT_METHODS = ['Stripe','PayID','Cash','Other'];

function toast(t){var e=document.getElementById('toast');e.textContent=t;e.style.display='block';
  setTimeout(function(){e.style.display='none';},2200);}

function getJob(row){ return JOBS.filter(function(x){return x.row===row;})[0]; }

/* ---------- PICKER (one-tap chooser for dropdown-constrained fields) ---------- */
var PICKER_OPTS = [], PICKER_CB = null, PICKER_SKIP = null;
function openPicker(title, options, cb, skipDefault){
  PICKER_OPTS = options; PICKER_CB = cb; PICKER_SKIP = skipDefault || null;
  document.getElementById('pk-title').textContent = title;
  document.getElementById('pk-opts').innerHTML = options.map(function(o,i){
    return '<button class="btn gh" onclick="pickerChoose('+i+')">'+esc(o)+'</button>';
  }).join('');
  document.getElementById('pk-skip').style.display = skipDefault ? '' : 'none';
  document.getElementById('picker').classList.add('on');
}
function pickerChoose(i){
  document.getElementById('picker').classList.remove('on');
  var cb=PICKER_CB; PICKER_CB=null;
  if(cb) cb(PICKER_OPTS[i]);
}
function pickerSkip(){
  document.getElementById('picker').classList.remove('on');
  var cb=PICKER_CB, def=PICKER_SKIP; PICKER_CB=null;
  if(cb && def) cb(def);
}
function pickerCancel(){
  document.getElementById('picker').classList.remove('on');
  PICKER_CB=null;
}
document.getElementById('picker').addEventListener('click',function(e){
  if(e.target.id==='picker') pickerCancel();
});

/* tabs */
document.querySelectorAll('.tab').forEach(function(tb){
  tb.onclick=function(){
    document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on');});
    tb.classList.add('on');
    ['jobs','new','quote','money','settings'].forEach(function(v){
      document.getElementById('v-'+v).style.display = (v===tb.dataset.t)?'block':'none';
    });
    if(tb.dataset.t==='settings') loadSettings();
    if(tb.dataset.t==='money') loadMoney();
  };
});

/* ---------- JOBS ---------- */
function drawFilters(){
  var f=['Active'].concat(STATUSES).concat(['All']);
  document.getElementById('filters').innerHTML = f.map(function(s){
    return '<div class="chip'+(s===FILTER?' on':'')+'" data-f="'+s+'">'+s+'</div>';
  }).join('');
  document.querySelectorAll('#filters .chip').forEach(function(c){
    c.onclick=function(){FILTER=c.dataset.f;drawFilters();drawList();};
  });
}

function visible(){
  if(FILTER==='All') return JOBS;
  if(FILTER==='Active') return JOBS.filter(function(j){
    return j.status!=='Delivered' && j.status!=='Closed' && j.status!=='Declined' && j.status!=='Lost';});
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
      + (j.source?'<div class="kv">Source: <b>'+esc(j.source)+'</b></div>':'')
      + (j.custType?'<div class="kv">Customer: <b>'+esc(j.custType)+'</b></div>':'')
      + (j.lane?'<div class="kv">Lane: <b>'+esc(j.lane)+'</b></div>':'')
      + (j.pqty?'<div class="kv">Qty: <b>'+esc(j.pqty)+' '+esc(j.unit)+'</b></div>':'')
      + (j.cost!==''?'<div class="kv">Cost ex GST: <b>$'+esc(j.cost)+'</b></div>':(j.pqty?'<div class="kv warn">⚠ No cost set — add one on the Pricing tab</div>':''))
      + (j.sell!==''?'<div class="kv">Sell ex GST: <b>$'+esc(j.sell)+'</b></div>':'')
      + (j.margin!==''?'<div class="kv">Margin ex GST: <b>$'+esc(j.margin)+'</b></div>':'')
      + (j.lostReason?'<div class="kv">Lost reason: <b>'+esc(j.lostReason)+'</b></div>':'')
      + (j.closedDate?'<div class="kv">Closed: <b>'+esc(j.closedDate)+'</b></div>':'')
      + (j.supplierCost?'<div class="kv">Supplier cost: <b>$'+esc(j.supplierCost)+'</b></div>':'')
      + (j.payMethod?'<div class="kv">Paid by: <b>'+esc(j.payMethod)+'</b></div>':'')
      + '<div class="acts">'
      + (tel?'<a class="btn" href="'+tel+'">Call</a>':'')
      + (sms?'<a class="btn" href="'+sms+'">Text</a>':'')
      + '<button class="btn gh" onclick="setQuoted('+j.row+')">Set $</button>'
      + '<button class="btn gh" onclick="setCarrier('+j.row+')">Carrier</button>'
      + '<button class="btn gh" onclick="setAddress('+j.row+')">Address</button>'
      + '<button class="btn gh" onclick="setJobDate('+j.row+')">Date</button>'
      + '<button class="btn gh" onclick="setNote('+j.row+')">Note</button>'
      + '<button class="btn gh" onclick="setSourceField('+j.row+')">Source</button>'
      + '<button class="btn gh" onclick="setCustType('+j.row+')">Cust.</button>'
      + '<button class="btn gh" onclick="setQtyUnit('+j.row+')">Qty</button>'
      + '<button class="btn gh" onclick="setLaneField('+j.row+')">Lane</button>'
      + '<button class="btn gh" onclick="setSupplierCost('+j.row+')">Cost</button>'
      + '<button class="btn gh" onclick="setPayMethod('+j.row+')">Paid by</button>'
      + '<button class="btn gh" onclick="quoteForJob('+j.row+')">Quote</button>'
      + '<button class="btn org" onclick="genInvoice('+j.row+')">Invoice</button>'
      + '<button class="btn gh" onclick="declineJob('+j.row+')">Declined</button>'
      + '<button class="btn gh" onclick="loseJob('+j.row+')">Lost</button>'
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
function setStatus(row,s){
  if(s==='Declined'||s==='Lost'){ openPicker(s+' — reason?', LOST_REASONS, function(r){ finishClose(row,s,r); }, 'Other'); return; }
  save(row,'Status',s,s);
}
function declineJob(row){ openPicker('Declined — reason?', LOST_REASONS, function(r){ finishClose(row,'Declined',r); }, 'Other'); }
function loseJob(row){ openPicker('Lost — reason?', LOST_REASONS, function(r){ finishClose(row,'Lost',r); }, 'Other'); }
function finishClose(row,status,reason){
  google.script.run.withSuccessHandler(function(){
    var j=getJob(row); if(j){ j.status=status; j.lostReason=reason; }
    toast(status+' — '+reason); load(); loadPipelineLine();
  }).withFailureHandler(function(e){toast('Failed: '+e.message);}).closeJob(row,status,reason);
}
function setSourceField(row){ openPicker('Source', SOURCES, function(v){ save(row,'Source',v,'Source saved'); }); }
function setCustType(row){ openPicker('Customer type', CUSTOMER_TYPES, function(v){ save(row,'Customer type',v,'Customer type saved'); }); }
function setLaneField(row){ openPicker('Lane', LANES, function(v){ save(row,'Lane',v,'Lane saved'); }); }
function setSupplierCost(row){
  var j=getJob(row);
  var v=prompt('Supplier cost ex GST (what you paid NCJ/Cowara)', j?j.supplierCost:'');
  if(v!==null) save(row,'Supplier cost',v,'Supplier cost saved');
}
function setPayMethod(row){
  openPicker('Payment method', PAYMENT_METHODS, function(v){ save(row,'Payment method',v,'Payment method saved'); });
}
function setQtyUnit(row){
  var j=getJob(row);
  var q=prompt('Qty', j?j.pqty:''); if(q===null) return;
  openPicker('Unit', UNITS, function(u){
    save(row,'Unit',u,'Unit saved');
    save(row,'Qty',q,'Qty saved');
    setTimeout(load,400);
  });
}
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

/* ---------- MONEY ---------- */
var MONEY_MONTH = null;
function loadMoney(){
  document.getElementById('mn-body').innerHTML='<div class="spin">Loading…</div>';
  google.script.run.withSuccessHandler(drawMoney)
    .withFailureHandler(function(e){
      document.getElementById('mn-body').innerHTML='<div class="empty">Could not load: '+e.message+'</div>';
    }).getMoneySummary(MONEY_MONTH);
}

function monthName(key){
  var p=String(key).split('-');
  var names=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return names[Number(p[1])-1]+' '+p[0];
}

function splitRows(obj){
  var keys=Object.keys(obj||{});
  if(!keys.length) return '<div class="mny-row">Nothing yet</div>';
  keys.sort(function(a,b){return obj[b]-obj[a];});
  return keys.map(function(k){
    return '<div class="mny-row"><span>'+esc(k)+'</span><b>$'+Math.round(obj[k]).toLocaleString()+'</b></div>';
  }).join('');
}

function drawMoney(d){
  MONEY_MONTH=d.monthKey;
  var sel=document.getElementById('mn-month');
  sel.innerHTML=d.months.map(function(m){
    return '<option value="'+m+'"'+(m===d.monthKey?' selected':'')+'>'+monthName(m)+'</option>';
  }).join('');

  var m=d.month, f=d.fytd;
  var marginCls=m.margin<0?' mny-neg':'';
  var h='<div class="mny-grid">'
    + '<div class="mny-cell hero"><div class="mny-lbl">Paid — '+esc(d.monthLabel)+'</div>'
    + '<div class="mny-big">'+m.paidText+'</div><div class="mny-sub">'+m.jobs+' job'+(m.jobs===1?'':'s')+'</div></div>'
    + '<div class="mny-cell hero"><div class="mny-lbl">Margin</div>'
    + '<div class="mny-big'+marginCls+'">'+m.marginText+'</div><div class="mny-sub">paid ex GST − supplier cost</div></div>'
    + '</div>'
    + '<div class="mny-grid">'
    + '<div class="mny-cell"><div class="mny-lbl">GST collected</div><div class="mny-mid">'+m.gstText+'</div></div>'
    + '<div class="mny-cell"><div class="mny-lbl">Supplier cost</div><div class="mny-mid">'+m.supplierCostText+'</div></div>'
    + '</div>';

  h+='<div class="mny-h">'+esc(d.fyLabel)+'</div>'
    + '<div class="mny-row"><span>Paid</span><b>'+f.paidText+'</b></div>'
    + '<div class="mny-row"><span>GST collected</span><b>'+f.gstText+'</b></div>'
    + '<div class="mny-row"><span>Supplier cost</span><b>'+f.supplierCostText+'</b></div>'
    + '<div class="mny-row"><span>Margin</span><b'+(f.margin<0?' class="mny-neg"':'')+'>'+f.marginText+'</b></div>';

  h+='<div class="mny-h">Invoiced — unpaid · '+d.unpaidTotalText+'</div>';
  if(!d.unpaid.length){
    h+='<div class="mny-row">Nothing outstanding</div>';
  } else {
    h+=d.unpaid.map(function(u){
      return '<div class="mny-row"><span>'+esc(u.name)+'<br><span style="color:#8d8d8d;font-size:.78rem">'
        +esc(u.invoiceNo)+(u.type?' · '+esc(u.type):'')+'</span></span><b>'+u.amountText+'</b></div>';
    }).join('');
  }

  h+='<div class="mny-h">'+esc(d.monthLabel)+' by type</div>'+splitRows(m.byType);
  h+='<div class="mny-h">'+esc(d.monthLabel)+' by payment method</div>'+splitRows(m.byMethod);

  document.getElementById('mn-body').innerHTML=h;
}

document.getElementById('mn-month').addEventListener('change',function(){
  MONEY_MONTH=this.value; loadMoney();
});

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

function loadPipelineLine(){
  google.script.run.withSuccessHandler(function(p){
    var el=document.getElementById('pline');
    if(!p){ el.textContent=''; return; }
    el.innerHTML='<b>'+esc(p.label)+':</b> '+p.truckLoads+' truck loads · $'+Math.round(p.margin)+' margin · '+p.swing+' lost to swing';
  }).withFailureHandler(function(){}).getPipelineSummary();
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
    notes:document.getElementById('n-notes').value,
    source:document.getElementById('n-source').value,
    custType:document.getElementById('n-custtype').value
  };
  if(!o.name && !o.phone){toast('Add a name or number');return;}
  this.textContent='Saving…'; this.disabled=true;
  var btn=this;
  google.script.run.withSuccessHandler(function(){
    ['n-name','n-phone','n-material','n-qty','n-suburb','n-notes'].forEach(function(id){
      document.getElementById(id).value='';});
    document.getElementById('n-source').value='';
    document.getElementById('n-custtype').value='';
    btn.textContent='Save enquiry'; btn.disabled=false;
    toast('Enquiry added'); load(); loadPipelineLine();
    document.querySelector('.tab[data-t="jobs"]').click();
  }).withFailureHandler(function(e){
    btn.textContent='Save enquiry'; btn.disabled=false; toast('Failed: '+e.message);
  }).addJob(o);
};

/* ---------- QUOTE ---------- */
var QUOTE_ROW = null;
function quoteForJob(row){
  QUOTE_ROW = row;
  var j = getJob(row);
  document.getElementById('q-bind-txt').textContent = 'Quoting for: ' + (j && j.name ? j.name : 'this job');
  document.getElementById('q-bind').classList.add('on');
  document.getElementById('q-save').style.display = '';
  document.getElementById('qt-save').style.display = '';
  document.querySelector('.tab[data-t="quote"]').click();
}
function clearQuoteBinding(){
  QUOTE_ROW = null;
  document.getElementById('q-bind').classList.remove('on');
  document.getElementById('q-save').style.display = 'none';
  document.getElementById('qt-save').style.display = 'none';
}

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

function saveLoadQuoteToJob(){
  if(!QUOTE_ROW){ toast('Open this from a job\\'s Quote button first'); return; }
  var sel=document.getElementById('q-mat'); var o=sel.options[sel.selectedIndex];
  if(!o){ toast('Pick a material first'); return; }
  var price=+o.dataset.price, unit=o.dataset.unit, dens=+o.dataset.dens;
  var L=+document.getElementById('q-l').value, W=+document.getElementById('q-w').value,
      D=+document.getElementById('q-d').value, amt=+document.getElementById('q-amt').value,
      au=document.getElementById('q-unit').value, del=+document.getElementById('q-del').value;
  var m3=0,t=0;
  if(amt>0){ if(au==='m3'){m3=amt; t=dens?amt*dens:0;} else {t=amt; m3=dens?amt/dens:amt;} }
  else if(L>0&&W>0&&D>0){ m3=L*W*(D/1000); t=dens?m3*dens:0; }
  if(m3<=0){ toast('Enter a quantity first'); return; }
  var bill=(unit==='t')?t:m3;
  var mat=bill*price;
  var tot=Math.round(mat+(del>0?del:0));
  // No lane guess from tonnage — the per-load calculator is trailer
  // stock; the server defaults to a trailer lane unless this job is
  // already Type "Trailer hire". Truck lanes only ever come from the
  // truck-load tab's own six NCJ products.
  google.script.run.withSuccessHandler(function(r){
    toast(r.costFound?'Saved to job':'Saved — no cost set for this product yet');
    load(); loadPipelineLine();
  }).withFailureHandler(function(e){toast('Failed: '+e.message);})
   .saveQuoteToJob(QUOTE_ROW,{product:o.value,qty:Math.round(bill*100)/100,unit:unit,total:tot});
}

/* ---------- TRUCK LOAD (10t flat price) ---------- */
document.getElementById('qm-load').onclick=function(){ setQuoteMode('load'); };
document.getElementById('qm-truck').onclick=function(){ setQuoteMode('truck'); };
function setQuoteMode(m){
  document.getElementById('qm-load').classList.toggle('on', m==='load');
  document.getElementById('qm-truck').classList.toggle('on', m==='truck');
  document.getElementById('q-load-mode').style.display = (m==='load') ? '' : 'none';
  document.getElementById('q-truck-mode').style.display = (m==='truck') ? '' : 'none';
}

var TRUCK = {zones:[], materials:[]};
function buildTruckMats(){
  google.script.run.withSuccessHandler(function(t){
    TRUCK = t;
    var zoneSel=document.getElementById('qt-zone'); zoneSel.innerHTML='';
    t.zones.concat(['Other (quote manually)']).forEach(function(z){
      var o=document.createElement('option'); o.value=z; o.textContent=z; zoneSel.appendChild(o);
    });
    var matSel=document.getElementById('qt-mat'); matSel.innerHTML='';
    t.materials.forEach(function(m){
      var o=document.createElement('option'); o.value=m[0]; o.textContent=m[0]; matSel.appendChild(o);
    });
    truckQuote();
  }).getTruckLoadPrices();
}

function truckQuote(){
  var matName=document.getElementById('qt-mat').value;
  var zoneName=document.getElementById('qt-zone').value;
  var out=document.getElementById('qt-out');
  var zoneIdx=TRUCK.zones.indexOf(zoneName);
  var row=TRUCK.materials.filter(function(m){return m[0]===matName;})[0];
  if(!row || zoneIdx===-1){
    out.innerHTML='<div class="ln">Not a standard truck-load zone — quote this one manually.</div>';
    out.classList.add('on');
    out.dataset.msg='';
    return;
  }
  var price=row[zoneIdx+1];
  out.innerHTML='<div class="big">$'+price+'</div>'
    + '<div class="ln">10 t '+matName+' — '+zoneName+'</div>'
    + '<div class="ln">All-in: material + delivery, no extras.</div>';
  out.classList.add('on');
  out.dataset.msg='A full truck load — 10 t of '+matName+' delivered, all up $'+price+'. '
    + 'One drop, no fuel levy, no extras.\\n\\nTo lock it in, payment first please: https://buy.stripe.com/6oUbIUdbo9APbiobRW7AI0B\\n'
    + 'Just enter $'+price+' when it asks for the amount 👍\\n\\nOnce that\\'s through I\\'ll confirm your delivery window.';
}
document.getElementById('qt-mat').addEventListener('change', truckQuote);
document.getElementById('qt-zone').addEventListener('change', truckQuote);
document.getElementById('qt-copy').onclick=function(){
  var m=document.getElementById('qt-out').dataset.msg;
  if(!m){toast('Pick a standard zone first');return;}
  var ta=document.createElement('textarea'); ta.value=m; document.body.appendChild(ta);
  ta.select(); try{document.execCommand('copy');toast('Quote copied');}catch(e){toast('Copy failed');}
  document.body.removeChild(ta);
};

function saveTruckQuoteToJob(){
  if(!QUOTE_ROW){ toast('Open this from a job\\'s Quote button first'); return; }
  var matName=document.getElementById('qt-mat').value;
  var zoneName=document.getElementById('qt-zone').value;
  var zoneIdx=TRUCK.zones.indexOf(zoneName);
  var row=TRUCK.materials.filter(function(m){return m[0]===matName;})[0];
  if(!row||zoneIdx===-1){ toast('Pick a standard zone first'); return; }
  var price=row[zoneIdx+1];
  google.script.run.withSuccessHandler(function(r){
    toast(r.costFound?'Saved to job':'Saved — no cost set for this product yet');
    load(); loadPipelineLine();
  }).withFailureHandler(function(e){toast('Failed: '+e.message);})
   .saveQuoteToJob(QUOTE_ROW,{product:matName,qty:10,unit:'t',lane:'Truck 10t (NCJ)',total:price});
}

load(); buildMats(); buildTruckMats(); loadPipelineLine();
</script>
</body></html>
`;