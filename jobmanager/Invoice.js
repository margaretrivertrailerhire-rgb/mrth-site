/**
 * MRTH Invoice Builder — ported into the Job Manager web app.
 * Generates invoices / tax invoices as PDFs from a row on the Enquiries
 * sheet (the same sheet the Jobs tab reads and writes). No spreadsheet
 * menu, no SpreadsheetApp.getUi() — every entry point here is called from
 * the phone app via google.script.run and returns plain data.
 *
 * GST
 *   Leave the GST start date unset until the ATO confirms registration.
 *   Jobs whose supply date falls before that date print as a plain
 *   invoice with no GST; on or after, they print as a tax invoice with
 *   the GST line and ABN. See gstAppliesOn().
 */

const INVOICE_CONFIG = {
  BUSINESS_NAME: 'MR Trailer Hire & Landscape Delivery', // exact ASIC-registered name (8 Sept 2026)
  TRADING_SHORT: 'MR Trailer Hire',
  ABN: '60 322 641 541',
  PHONE: '0429 016 758',
  EMAIL: 'margaretrivertrailerhire@gmail.com',
  WEBSITE: 'margaretrivertrailerhire.com.au',

  DRIVE_FOLDER: 'MRTH Invoices',
  PRICING_SHEET: 'Pricing',

  PAYMENT_TERMS: 'Payment due on delivery unless otherwise agreed.',
  STRIPE_LINK: 'https://buy.stripe.com/6oUbIUdbo9APbiobRW7AI0B', // generic choose-your-amount link, also used by the in-app quoter
  PAYID: 'margaretrivertrailerhire@gmail.com', // email-form PayID — printed instead of BSB/account, fewer digits to mistype and the payer's bank confirms the account name

  // 'draft' creates a Gmail draft, 'send' sends immediately. Leave on
  // draft until Neil says otherwise — every invoice gets a human look
  // before it leaves.
  EMAIL_MODE: 'draft',

  REVIEW_LINK: 'https://g.page/r/CaYV_-uiUviZEBI/review',

  // Hosted rather than inlined: mail clients routinely strip data-URI
  // images, so an emailed logo has to come from a public URL. The PDF
  // keeps its embedded copy (see Logo.js) because the HTML-to-PDF
  // converter can't be relied on to fetch anything at render time.
  LOGO_URL: 'https://www.margaretrivertrailerhire.com.au/assets/mrth-logo-light-horizontal.png',

  DARK: '#1A1A1A',
  ORANGE: '#E8651A',
  GREY: '#6B6B6B',
  RULE: '#E0DEDB',
};

const INVOICE_PROPS = PropertiesService.getScriptProperties();

// ---------------------------------------------------------------------------
// GST state
// ---------------------------------------------------------------------------

/** Returns a Date, or null if not registered yet. */
function gstStartDate() {
  const raw = INVOICE_PROPS.getProperty('GST_START');
  if (!raw) return null;
  const d = new Date(raw + 'T00:00:00');
  return isNaN(d.getTime()) ? null : d;
}

/**
 * GST applies to a supply made on or after the registration start date.
 * Before that date you are not registered, so you cannot charge GST and
 * must not issue a document titled "Tax Invoice".
 */
function gstAppliesOn(supplyDate) {
  const start = gstStartDate();
  if (!start) return false;
  return supplyDate.getTime() >= start.getTime();
}

function formatInvoiceNumber(n) {
  return 'MRTH-' + String(n).padStart(4, '0');
}

function nextInvoiceNumber_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const next = Number(INVOICE_PROPS.getProperty('INVOICE_SEQ') || 0) + 1;
    INVOICE_PROPS.setProperty('INVOICE_SEQ', String(next));
    return next;
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// API called from the phone — settings screen
// ---------------------------------------------------------------------------

function getInvoiceSettings() {
  const start = INVOICE_PROPS.getProperty('GST_START') || null;
  const nextSeq = Number(INVOICE_PROPS.getProperty('INVOICE_SEQ') || 0) + 1;
  return {
    abn: INVOICE_CONFIG.ABN,
    gstStart: start,
    nextSeq: nextSeq,
    nextInvoiceNumber: formatInvoiceNumber(nextSeq),
    emailMode: INVOICE_CONFIG.EMAIL_MODE,
    driveFolder: INVOICE_CONFIG.DRIVE_FOLDER,
  };
}

/** Pass '' to clear back to unregistered. */
function setGstStartDate(val) {
  val = String(val || '').trim();
  if (!val) {
    INVOICE_PROPS.deleteProperty('GST_START');
    return getInvoiceSettings();
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(val) || isNaN(new Date(val + 'T00:00:00').getTime())) {
    throw new Error('That is not a valid date (use YYYY-MM-DD).');
  }
  INVOICE_PROPS.setProperty('GST_START', val);
  return getInvoiceSettings();
}

function setInvoiceCounter(n) {
  n = parseInt(n, 10);
  if (isNaN(n) || n < 1) throw new Error('Enter a whole number of 1 or more.');
  INVOICE_PROPS.setProperty('INVOICE_SEQ', String(n - 1));
  return getInvoiceSettings();
}

// ---------------------------------------------------------------------------
// Reading a job row (Enquiries sheet — sheet_()/headerMap_() from Code.js)
// ---------------------------------------------------------------------------

function readJobRow_(row) {
  const sh = sheet_();
  const map = headerMap_(sh);
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const get = function (field) {
    const idx = map[field];
    if (idx === undefined) return '';
    const v = values[idx];
    return v === null || v === undefined ? '' : v;
  };
  return { sheet: sh, row: row, map: map, get: get };
}

function toDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (v) {
    const d = new Date(v);
    if (!isNaN(d.getTime())) return d;
  }
  return new Date();
}

/** Quoted $ is free-typed on a phone ("850", "$850", "850 + delivery"...).
 *  Pull the first dollar figure out of it — the caller must always show
 *  the parsed number back before generating, so a figure like
 *  "850 + delivery" can't silently become a final $850 invoice. */
function parseAmount_(v) {
  if (typeof v === 'number') return v;
  const m = String(v || '').match(/[\d,]+(\.\d+)?/);
  if (!m) return 0;
  return parseFloat(m[0].replace(/,/g, '')) || 0;
}

function money_(n) {
  return '$' + n.toLocaleString('en-AU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function fmtDate_(d) {
  return Utilities.formatDate(d, 'Australia/Perth', 'd MMMM yyyy');
}

function escapeHtml_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Which date is the supply for GST purposes. Paid date wins over Delivered
 * date: the job lifecycle takes payment before the carrier is booked, so
 * plenty of invoices are issued before delivery happens. On cash-basis
 * GST the liability attaches when the money lands, not when the job is
 * completed — that's the date this business should be reporting against.
 */
function resolveSupplyDate_(ctx) {
  const paid = ctx.get('Paid date');
  if (paid) return { date: toDate_(paid), source: 'Paid date' };
  const delivered = ctx.get('Delivered date');
  if (delivered) return { date: toDate_(delivered), source: 'Delivered date' };
  return { date: new Date(), source: 'today — no Paid or Delivered date on this job yet' };
}

// ---------------------------------------------------------------------------
// Pricing tab — Stripe link lookup for trailer-hire jobs
// ---------------------------------------------------------------------------

function pricingRows_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(INVOICE_CONFIG.PRICING_SHEET);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 4).getValues()
    .map(function (r) {
      return { trailer: String(r[0] || '').trim(), hireType: String(r[1] || '').trim(),
        price: r[2], link: String(r[3] || '').trim() };
    })
    .filter(function (r) { return r.trailer && r.link; });
}

/** Enquiries has no dedicated Trailer/rate column, so this scans the job's
 *  free-text fields for a trailer name plus a rate word. Only returns a
 *  match when exactly one Pricing row fits — anything less certain falls
 *  back to the generic link, and the caller always shows which happened
 *  before generating so a wrong guess is easy to catch. */
function findPricingMatch_(ctx) {
  const hay = [ctx.get('Material'), ctx.get('Notes'), ctx.get('Load Type')]
    .map(function (v) { return String(v || ''); }).join(' ').toLowerCase();
  if (!hay.trim()) return null;
  const matches = pricingRows_().filter(function (r) {
    return hay.indexOf(r.trailer.toLowerCase()) !== -1 && hay.indexOf(r.hireType.toLowerCase()) !== -1;
  });
  return matches.length === 1 ? matches[0] : null;
}

// ---------------------------------------------------------------------------
// Invoice assembly
// ---------------------------------------------------------------------------

/** How many tonnes a trailer load can carry before it should really be
 *  described by volume instead. Matches the figure quoted on the website. */
const TRAILER_TONNE_LIMIT = 2.4;

/**
 * The job's quantity as a number plus a unit, for the invoice's Qty/Unit
 * columns. Prefers the structured Qty/Unit fields the quoter writes;
 * falls back to reading them out of the free-text Quantity field
 * ("4 m3", "2.5 t") for jobs entered before those existed or typed by
 * hand. Returns blank qty when neither yields a number, in which case
 * the caller keeps the free text in the description instead.
 */
function resolveQuantity_(ctx) {
  const norm = function (u) {
    u = String(u || '').trim().toLowerCase();
    if (u === 't' || u === 'tonne' || u === 'tonnes') return 't';
    if (u === 'm3' || u === 'm³' || u === 'cube' || u === 'cubes') return 'm3';
    return '';
  };
  var qty = parseFloat(ctx.get('Qty'));
  var unit = norm(ctx.get('Unit'));

  if (isNaN(qty) || qty <= 0) {
    const m = String(ctx.get('Quantity') || '').match(/([\d.]+)\s*(tonnes|tonne|m3|m³|cubes|cube|t)\b/i);
    if (m) { qty = parseFloat(m[1]); unit = norm(m[2]); }
  }
  if (isNaN(qty) || qty <= 0) return { qty: 0, unit: '', tonnes: 0 };
  return { qty: qty, unit: unit, tonnes: unit === 't' ? qty : 0 };
}

/**
 * Builds the invoice model from a job row. Sheet prices are the final
 * amount the customer pays — when GST applies, that amount is
 * GST-inclusive and GST is one eleventh of it, never ten percent on top.
 */
function buildInvoiceModel_(ctx, invoiceNumberFormatted) {
  const supply = resolveSupplyDate_(ctx);
  const gst = gstAppliesOn(supply.date);

  const rawQuoted = String(ctx.get('Quoted $') || '').trim();
  const total = parseAmount_(rawQuoted);
  const gstAmount = gst ? Math.round((total / 11) * 100) / 100 : 0;

  const material = String(ctx.get('Material') || '').trim();
  const type = String(ctx.get('Type') || '').trim();
  const qtyText = String(ctx.get('Quantity') || '').trim();
  const loadType = String(ctx.get('Load Type') || '').trim();
  const lane = String(ctx.get('Lane') || '').trim();
  const address = String(ctx.get('Delivery address') || '').trim() || String(ctx.get('Suburb') || '').trim();

  const measure = resolveQuantity_(ctx);

  var description = material || type || 'Materials delivery';
  if (loadType) description += ' — ' + loadType;
  // Only fold the free-text quantity back into the description when it
  // couldn't be split into its own Qty/Unit columns, so it isn't printed twice.
  if (!measure.qty && qtyText) description += ' — ' + qtyText;
  // The address isn't repeated here — it already prints in the "Billed to"
  // block, and both read from the same field, so a "Delivered to" line
  // would print the same string twice on a one-page invoice.

  const pricingMatch = findPricingMatch_(ctx);

  return {
    invoiceNumber: invoiceNumberFormatted,
    issueDate: new Date(),
    supplyDate: supply.date,
    supplyDateSource: supply.source,
    gst: gst,
    title: gst ? 'Tax Invoice' : 'Invoice',
    customerName: String(ctx.get('Name') || '').trim(),
    customerEmail: String(ctx.get('Email') || '').trim(),
    customerPhone: String(ctx.get('Phone') || '').trim(),
    customerAddress: address,
    lane: lane,
    qty: measure.qty,
    unit: measure.unit,
    tonnes: measure.tonnes,
    lines: [{ description: description, qty: measure.qty, unit: measure.unit, amount: total }],
    rawQuoted: rawQuoted,
    total: total,
    gstAmount: gstAmount,
    subtotal: total - gstAmount,
    stripeLink: pricingMatch ? pricingMatch.link : INVOICE_CONFIG.STRIPE_LINK,
    pricingMatch: pricingMatch ? (pricingMatch.trailer + ' / ' + pricingMatch.hireType) : null,
  };
}

/**
 * A tax invoice of $1,000 or more must identify the buyer — a name or an
 * ABN satisfies that, an address is not required. The customer Name field
 * is what's checked; the missing address column on Enquiries is not a
 * compliance blocker.
 */
function validateModel_(m) {
  const problems = [];
  if (!m.total || m.total <= 0) {
    problems.push('No usable amount found in "Quoted $" — check that field on this job.');
  }
  if (!m.customerName) {
    problems.push('No customer name on this job.');
  }
  if (m.gst && m.total >= 1000 && !m.customerName) {
    problems.push('Tax invoices of $1,000 or more must identify the buyer — this job has no name.');
  }
  if (/^trailer/i.test(m.lane || '') && m.tonnes > TRAILER_TONNE_LIMIT) {
    problems.push('Trailer load billed as ' + m.tonnes + ' t — over the ' + TRAILER_TONNE_LIMIT +
      ' t a trailer carries. Above that a trailer load should be described by volume (m³), not weight.');
  }
  return problems;
}

// ---------------------------------------------------------------------------
// Template
// ---------------------------------------------------------------------------

function renderInvoiceHtml_(m) {
  const fmtQty = function (n) {
    if (!n) return '';
    return String(Math.round(n * 100) / 100);
  };
  const unitLabel = function (u) {
    return u === 'm3' ? 'm&sup3;' : escapeHtml_(u || '');
  };

  const linesHtml = m.lines.map(function (l) {
    const parts = String(l.description).split('\n');
    const head = escapeHtml_(parts[0]);
    const rest = parts.slice(1).map(function (p) {
      return '<div class="sub">' + escapeHtml_(p) + '</div>';
    }).join('');
    return '<tr>' +
      '<td>' + head + rest + '</td>' +
      '<td class="num">' + fmtQty(l.qty) + '</td>' +
      '<td class="num">' + unitLabel(l.unit) + '</td>' +
      '<td class="amt">' + money_(l.amount) + '</td>' +
      '</tr>';
  }).join('');

  const totalsHtml = m.gst
    ? '<tr><td>Subtotal</td><td class="amt">' + money_(m.subtotal) + '</td></tr>' +
      '<tr><td>GST</td><td class="amt">' + money_(m.gstAmount) + '</td></tr>' +
      '<tr class="grand"><td>Total</td><td class="amt">' + money_(m.total) + '</td></tr>'
    : '<tr class="grand"><td>Total</td><td class="amt">' + money_(m.total) + '</td></tr>';

  const gstNote = m.gst
    ? 'Total includes GST of ' + money_(m.gstAmount) + '.'
    : INVOICE_CONFIG.TRADING_SHORT + ' is not registered for GST. No GST has been charged on this invoice.';

  var payLines = '<div class="pay-line">' + escapeHtml_(INVOICE_CONFIG.PAYMENT_TERMS) + '</div>';
  if (m.stripeLink) {
    payLines += '<div class="pay-line">Card or Apple Pay: ' + escapeHtml_(m.stripeLink) + '</div>';
  }
  if (INVOICE_CONFIG.PAYID) {
    payLines += '<div class="pay-line">PayID: ' + escapeHtml_(INVOICE_CONFIG.PAYID) +
      ' &nbsp;—&nbsp; reference ' + escapeHtml_(m.invoiceNumber) + '</div>';
  }

  const buyerBlock = [
    m.customerName ? '<strong>' + escapeHtml_(m.customerName) + '</strong>' : '',
    m.customerAddress ? escapeHtml_(m.customerAddress) : '',
    m.customerPhone ? escapeHtml_(String(m.customerPhone)) : '',
    m.customerEmail ? escapeHtml_(m.customerEmail) : '',
  ].filter(Boolean).join('<br>');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    '@page { size: A4; margin: 16mm 15mm; }' +
    'body { font-family: Helvetica, Arial, sans-serif; color: #000; font-size: 10pt; line-height: 1.45; }' +
    'table { border-collapse: collapse; }' +

    'table.hdr { width: 100%; margin-bottom: 10px; }' +
    'table.hdr td { vertical-align: top; }' +
    'td.logo { width: 68mm; }' +
    'td.logo img { width: 64mm; display: block; }' +
    'td.ident { text-align: right; font-size: 8.5pt; line-height: 1.5; }' +
    // The logo already carries the wordmark, so the name is set small here
    // rather than as a second competing headline — it stays as text because
    // an invoice whose only supplier identity is inside an image has none
    // at all if that image fails to render.
    'td.ident .nm { font-size: 9pt; font-weight: bold; }' +
    '.rule { border-bottom: 2.5pt solid ' + INVOICE_CONFIG.ORANGE + '; margin-bottom: 16px; }' +

    'h1.doc { font-size: 16pt; margin: 0 0 10px; font-weight: bold; letter-spacing: 0.3px; }' +
    'table.meta { width: 100%; margin-bottom: 20px; font-size: 9.5pt; }' +
    'table.meta td { vertical-align: top; width: 50%; padding-right: 10px; }' +
    'table.meta .lbl { font-size: 7.5pt; text-transform: uppercase; letter-spacing: 0.6px; ' +
      'color: ' + INVOICE_CONFIG.GREY + '; padding-bottom: 3px; }' +
    'table.meta .row { padding-bottom: 6px; }' +

    'table.items { width: 100%; margin-bottom: 2px; }' +
    'table.items th { background: #F2F2F2; text-align: left; font-size: 7.5pt; text-transform: uppercase; ' +
      'letter-spacing: 0.6px; font-weight: bold; padding: 7px 8px; border-bottom: 0.75pt solid #CCC; }' +
    'table.items td { padding: 9px 8px; border-bottom: 0.75pt solid #E2E2E2; vertical-align: top; }' +
    'table.items th.num, table.items td.num { text-align: right; white-space: nowrap; width: 15mm; }' +
    'table.items th.amt, table.items td.amt { text-align: right; white-space: nowrap; width: 26mm; }' +
    '.sub { color: ' + INVOICE_CONFIG.GREY + '; font-size: 8.5pt; margin-top: 2px; }' +

    'table.totals { margin-left: auto; margin-top: 8px; min-width: 62mm; }' +
    'table.totals td { padding: 4px 8px; font-size: 10pt; }' +
    'table.totals td.amt { text-align: right; white-space: nowrap; }' +
    'table.totals tr.grand td { border-top: 1pt solid #000; padding-top: 7px; font-weight: bold; font-size: 12pt; }' +
    '.gst-note { font-size: 8.5pt; color: ' + INVOICE_CONFIG.GREY + '; text-align: right; margin-top: 6px; }' +

    '.pay { margin-top: 26px; border: 0.75pt solid #CCC; padding: 12px 14px; }' +
    '.pay .ttl { font-size: 7.5pt; text-transform: uppercase; letter-spacing: 0.6px; font-weight: bold; ' +
      'margin-bottom: 6px; }' +
    '.pay-line { font-size: 9.5pt; padding: 1.5px 0; }' +
    '.review { margin-top: 10px; font-size: 8.5pt; color: ' + INVOICE_CONFIG.GREY + '; }' +
    '</style></head><body>' +

    '<table class="hdr"><tr>' +
      '<td class="logo"><img src="' + INVOICE_LOGO_DATA_URI + '"></td>' +
      '<td class="ident">' +
        '<div class="nm">' + escapeHtml_(INVOICE_CONFIG.BUSINESS_NAME) + '</div>' +
        '<div>ABN ' + escapeHtml_(INVOICE_CONFIG.ABN) + '</div>' +
        '<div>' + escapeHtml_(INVOICE_CONFIG.PHONE) + '</div>' +
        '<div>' + escapeHtml_(INVOICE_CONFIG.EMAIL) + '</div>' +
        '<div>' + escapeHtml_(INVOICE_CONFIG.WEBSITE) + '</div>' +
      '</td>' +
    '</tr></table>' +
    '<div class="rule"></div>' +

    '<h1 class="doc">' + escapeHtml_(m.title) + '</h1>' +

    '<table class="meta"><tr>' +
      '<td>' +
        '<div class="row"><div class="lbl">Invoice number</div>' + escapeHtml_(m.invoiceNumber) + '</div>' +
        '<div class="row"><div class="lbl">Invoice date</div>' + fmtDate_(m.issueDate) + '</div>' +
        '<div class="row"><div class="lbl">Date of supply</div>' + fmtDate_(m.supplyDate) + '</div>' +
      '</td>' +
      '<td>' +
        '<div class="lbl">Billed to</div>' + (buyerBlock || '—') +
      '</td>' +
    '</tr></table>' +

    '<table class="items">' +
      '<tr><th>Description</th><th class="num">Qty</th><th class="num">Unit</th><th class="amt">Amount</th></tr>' +
      linesHtml +
    '</table>' +

    '<table class="totals">' + totalsHtml + '</table>' +
    '<div class="gst-note">' + escapeHtml_(gstNote) + '</div>' +

    '<div class="pay">' +
      '<div class="ttl">Payment</div>' +
      payLines +
    '</div>' +

    '<div class="review">Thanks for supporting a local business. If the delivery went well, ' +
      'a Google review genuinely helps: ' + escapeHtml_(INVOICE_CONFIG.REVIEW_LINK) + '</div>' +

    '</body></html>';
}

// ---------------------------------------------------------------------------
// API called from the phone — generate / email
// ---------------------------------------------------------------------------

/** Read-only: builds the model against the *current* next invoice number
 *  without consuming it, so cancelling never burns a number. */
function previewInvoice(row) {
  const ctx = readJobRow_(row);
  const seq = Number(INVOICE_PROPS.getProperty('INVOICE_SEQ') || 0) + 1;
  const model = buildInvoiceModel_(ctx, formatInvoiceNumber(seq));
  const problems = validateModel_(model);
  const existing = String(ctx.get('Invoice No') || '').trim();

  return {
    row: row,
    title: model.title,
    invoiceNumber: model.invoiceNumber,
    gst: model.gst,
    customerName: model.customerName,
    supplyDateFormatted: fmtDate_(model.supplyDate),
    supplyDateSource: model.supplyDateSource,
    rawQuoted: model.rawQuoted,
    totalFormatted: money_(model.total),
    gstAmountFormatted: money_(model.gstAmount),
    pricingMatch: model.pricingMatch,
    alreadyInvoiced: existing || null,
    warnings: problems,
  };
}

/** Actually issues the invoice: consumes the next number under lock,
 *  renders the PDF, saves it to Drive, writes the invoice number/link
 *  back onto the job row. */
function confirmGenerateInvoice(row) {
  const ctx = readJobRow_(row);
  const model = buildInvoiceModel_(ctx, formatInvoiceNumber(nextInvoiceNumber_()));

  const filename = model.invoiceNumber + ' — ' + (model.customerName || 'Customer').replace(/[\\/:*?"<>|]/g, '');
  const pdf = Utilities
    .newBlob(renderInvoiceHtml_(model), MimeType.HTML, filename + '.html')
    .getAs(MimeType.PDF)
    .setName(filename + '.pdf');

  const file = getInvoiceFolder_().createFile(pdf);
  ctx.sheet.getRange(ctx.row, ctx.map['Invoice No'] + 1).setValue(model.invoiceNumber);
  ctx.sheet.getRange(ctx.row, ctx.map['Invoice Link'] + 1).setValue(file.getUrl());

  return {
    invoiceNumber: model.invoiceNumber,
    title: model.title,
    gst: model.gst,
    totalFormatted: money_(model.total),
    gstAmountFormatted: money_(model.gstAmount),
    driveUrl: file.getUrl(),
    customerEmail: model.customerEmail,
    customerName: model.customerName,
  };
}

/** Emails (or drafts) the invoice that was just generated, re-fetching the
 *  saved PDF from Drive rather than re-rendering — the row already has an
 *  invoice number by this point, so the model is rebuilt around that
 *  fixed number rather than drawing a new one. */
function createInvoiceEmailDraft(row) {
  const ctx = readJobRow_(row);
  const invoiceNumber = String(ctx.get('Invoice No') || '').trim();
  const invoiceLink = String(ctx.get('Invoice Link') || '').trim();
  if (!invoiceNumber || !invoiceLink) throw new Error('This job has not had an invoice generated yet.');

  const model = buildInvoiceModel_(ctx, invoiceNumber);
  if (!model.customerEmail) throw new Error('No email address on this job.');

  const fileId = extractDriveFileId_(invoiceLink);
  const pdf = DriveApp.getFileById(fileId).getBlob();

  const first = (model.customerName || '').split(' ')[0] || 'there';
  const subject = model.title + ' ' + model.invoiceNumber + ' — ' + INVOICE_CONFIG.TRADING_SHORT;
  const body =
    'Hi ' + first + ',\n\n' +
    'Your ' + model.title.toLowerCase() + ' for ' + money_(model.total) + ' is attached.\n\n' +
    INVOICE_CONFIG.PAYMENT_TERMS + '\n' +
    (model.stripeLink ? 'Pay by card or Apple Pay: ' + model.stripeLink + '\n' : '') +
    (INVOICE_CONFIG.PAYID ? 'Or PayID: ' + INVOICE_CONFIG.PAYID + ' — reference ' + model.invoiceNumber + '\n' : '') +
    '\nThanks again for the job. If you were happy with the delivery, a quick Google review makes a real difference to a small local business:\n' +
    INVOICE_CONFIG.REVIEW_LINK + '\n\n' +
    'Cheers,\n' + INVOICE_CONFIG.TRADING_SHORT + '\n' + INVOICE_CONFIG.WEBSITE;

  // Plain `body` above stays as the text fallback for clients that don't
  // render HTML; htmlBody is what most people will actually see.
  const htmlBody =
    '<div style="font-family:Helvetica,Arial,sans-serif;font-size:14px;line-height:1.55;color:#1A1A1A">' +
      (INVOICE_CONFIG.LOGO_URL
        ? '<img src="' + INVOICE_CONFIG.LOGO_URL + '" width="200" alt="' +
          escapeHtml_(INVOICE_CONFIG.BUSINESS_NAME) +
          '" style="width:200px;max-width:100%;height:auto;display:block;margin:0 0 20px">'
        : '') +
      '<p>Hi ' + escapeHtml_(first) + ',</p>' +
      '<p>Your ' + escapeHtml_(model.title.toLowerCase()) + ' for ' + money_(model.total) + ' is attached.</p>' +
      '<p>' + escapeHtml_(INVOICE_CONFIG.PAYMENT_TERMS) + '<br>' +
      (model.stripeLink
        ? 'Pay by card or Apple Pay: <a href="' + model.stripeLink + '">' + escapeHtml_(model.stripeLink) + '</a><br>'
        : '') +
      (INVOICE_CONFIG.PAYID
        ? 'Or PayID: ' + escapeHtml_(INVOICE_CONFIG.PAYID) + ' — reference ' + escapeHtml_(model.invoiceNumber)
        : '') +
      '</p>' +
      '<p>Thanks again for the job. If you were happy with the delivery, a quick Google review ' +
      'makes a real difference to a small local business:<br>' +
      '<a href="' + INVOICE_CONFIG.REVIEW_LINK + '">' + escapeHtml_(INVOICE_CONFIG.REVIEW_LINK) + '</a></p>' +
      '<p>Cheers,<br>' + escapeHtml_(INVOICE_CONFIG.TRADING_SHORT) + '<br>' +
      '<a href="https://' + INVOICE_CONFIG.WEBSITE + '">' + escapeHtml_(INVOICE_CONFIG.WEBSITE) + '</a></p>' +
    '</div>';

  const opts = { attachments: [pdf], name: INVOICE_CONFIG.TRADING_SHORT, htmlBody: htmlBody };
  if (INVOICE_CONFIG.EMAIL_MODE === 'send') {
    GmailApp.sendEmail(model.customerEmail, subject, body, opts);
  } else {
    GmailApp.createDraft(model.customerEmail, subject, body, opts);
  }
  return true;
}

/**
 * Makes an already-generated invoice PDF shareable and hands back the link.
 *
 * Drive files are created private, so the Invoice Link stored on the row
 * only opens for the account that made it — send that to a customer and
 * they get a "request access" screen instead of their invoice. This flips
 * the file to view-by-link at the moment of sharing rather than exposing
 * every invoice at generation time, so only the ones actually sent out
 * become readable by anyone holding the URL.
 */
function getInvoiceShareLink(row) {
  const ctx = readJobRow_(row);
  const invoiceNumber = String(ctx.get('Invoice No') || '').trim();
  const link = String(ctx.get('Invoice Link') || '').trim();
  if (!invoiceNumber || !link) throw new Error('No invoice has been generated for this job yet.');

  const fileId = extractDriveFileId_(link);
  const file = DriveApp.getFileById(fileId);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // file.getUrl() is Drive's /view page — on a phone that deep-links into
  // the Drive app and asks for a sign-in, which is no use to a customer
  // (or to anyone with a screen lock on Drive). The uc?export=download
  // endpoint hands back the PDF bytes instead, so it opens in whatever
  // PDF viewer the recipient already has, no Google account involved.
  const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;

  const title = gstAppliesOn(resolveSupplyDate_(ctx).date) ? 'Tax Invoice' : 'Invoice';
  return {
    url: downloadUrl,
    viewUrl: file.getUrl(),
    title: title,
    invoiceNumber: invoiceNumber,
    shareText: title + ' ' + invoiceNumber + ' from ' + INVOICE_CONFIG.TRADING_SHORT + '\n' + downloadUrl
  };
}

function extractDriveFileId_(url) {
  const m = String(url).match(/[-\w]{25,}/);
  if (!m) throw new Error('Could not read a Drive file ID from the saved invoice link.');
  return m[0];
}

function getInvoiceFolder_() {
  const it = DriveApp.getFoldersByName(INVOICE_CONFIG.DRIVE_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(INVOICE_CONFIG.DRIVE_FOLDER);
}
