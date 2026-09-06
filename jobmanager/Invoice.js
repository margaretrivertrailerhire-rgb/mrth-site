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
  BUSINESS_NAME: 'Margaret River Trailer Hire & Landscape Delivery',
  TRADING_SHORT: 'Margaret River Trailer Hire',
  ABN: '60 322 641 541',
  EMAIL: 'margaretrivertrailerhire@gmail.com',
  WEBSITE: 'margaretrivertrailerhire.com.au',

  DRIVE_FOLDER: 'MRTH Invoices',
  PRICING_SHEET: 'Pricing',

  PAYMENT_TERMS: 'Payment due on delivery unless otherwise agreed.',
  STRIPE_LINK: 'https://buy.stripe.com/6oUbIUdbo9APbiobRW7AI0B', // generic choose-your-amount link, also used by the in-app quoter
  BANK_NAME: '',
  BANK_BSB: '',
  BANK_ACCT: '',

  // 'draft' creates a Gmail draft, 'send' sends immediately. Leave on
  // draft until Neil says otherwise — every invoice gets a human look
  // before it leaves.
  EMAIL_MODE: 'draft',

  REVIEW_LINK: 'https://g.page/r/CaYV_-uiUviZEBI/review',

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
  const qty = String(ctx.get('Quantity') || '').trim();
  const loadType = String(ctx.get('Load Type') || '').trim();
  const address = String(ctx.get('Delivery address') || '').trim() || String(ctx.get('Suburb') || '').trim();

  var description = material || type || 'Materials delivery';
  const bits = [qty, loadType].filter(Boolean).join(' — ');
  if (bits) description += ' — ' + bits;
  if (address) description += '\nDelivered to ' + address;

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
    lines: [{ description: description, amount: total }],
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
  return problems;
}

// ---------------------------------------------------------------------------
// Template
// ---------------------------------------------------------------------------

function renderInvoiceHtml_(m) {
  const linesHtml = m.lines.map(function (l) {
    return '<tr>' +
      '<td class="desc">' + escapeHtml_(l.description).replace(/\n/g, '<br><span class="sub">') + '</td>' +
      '<td class="amt">' + money_(l.amount) + '</td>' +
      '</tr>';
  }).join('');

  const totalsHtml = m.gst
    ? '<tr><td>Subtotal</td><td class="amt">' + money_(m.subtotal) + '</td></tr>' +
      '<tr><td>GST</td><td class="amt">' + money_(m.gstAmount) + '</td></tr>' +
      '<tr class="grand"><td>Total</td><td class="amt">' + money_(m.total) + '</td></tr>'
    : '<tr class="grand"><td>Total</td><td class="amt">' + money_(m.total) + '</td></tr>';

  const gstNote = m.gst
    ? '<p class="note">Total includes GST of ' + money_(m.gstAmount) + '.</p>'
    : '<p class="note">' + INVOICE_CONFIG.TRADING_SHORT + ' is not registered for GST. No GST has been charged on this invoice.</p>';

  var payment = '<p class="note">' + escapeHtml_(INVOICE_CONFIG.PAYMENT_TERMS) + '</p>';
  if (m.stripeLink) {
    payment += '<p class="note">Pay by card or Apple Pay: <a href="' + m.stripeLink + '">' + m.stripeLink + '</a></p>';
  }
  if (INVOICE_CONFIG.BANK_BSB && INVOICE_CONFIG.BANK_ACCT) {
    payment += '<p class="note">Bank transfer: ' + escapeHtml_(INVOICE_CONFIG.BANK_NAME) +
      ' &nbsp; BSB ' + escapeHtml_(INVOICE_CONFIG.BANK_BSB) +
      ' &nbsp; Acct ' + escapeHtml_(INVOICE_CONFIG.BANK_ACCT) +
      '<br>Reference: ' + m.invoiceNumber + '</p>';
  }

  const buyerBlock = [
    m.customerName ? '<strong>' + escapeHtml_(m.customerName) + '</strong>' : '',
    m.customerAddress ? escapeHtml_(m.customerAddress) : '',
    m.customerPhone ? escapeHtml_(String(m.customerPhone)) : '',
    m.customerEmail ? escapeHtml_(m.customerEmail) : '',
  ].filter(Boolean).join('<br>');

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    '@page { size: A4; margin: 18mm 16mm; }' +
    'body { font-family: Helvetica, Arial, sans-serif; color: ' + INVOICE_CONFIG.DARK + '; font-size: 11pt; line-height: 1.5; }' +
    '.head { border-bottom: 3px solid ' + INVOICE_CONFIG.ORANGE + '; padding-bottom: 14px; margin-bottom: 26px; }' +
    '.biz { font-size: 15pt; font-weight: bold; letter-spacing: -0.2px; }' +
    '.biz-meta { color: ' + INVOICE_CONFIG.GREY + '; font-size: 9.5pt; margin-top: 5px; }' +
    '.title { float: right; text-align: right; }' +
    '.title h1 { font-size: 17pt; margin: 0; font-weight: bold; }' +
    '.title .num { color: ' + INVOICE_CONFIG.ORANGE + '; font-weight: bold; font-size: 11pt; margin-top: 3px; }' +
    '.parties { width: 100%; margin-bottom: 28px; }' +
    '.parties td { vertical-align: top; width: 50%; font-size: 10pt; }' +
    '.lbl { color: ' + INVOICE_CONFIG.GREY + '; font-size: 8.5pt; margin-bottom: 5px; }' +
    'table.items { width: 100%; border-collapse: collapse; margin-bottom: 4px; }' +
    'table.items th { text-align: left; font-size: 8.5pt; color: ' + INVOICE_CONFIG.GREY + '; font-weight: normal; ' +
      'border-bottom: 1px solid ' + INVOICE_CONFIG.DARK + '; padding-bottom: 6px; }' +
    'table.items th.amt, table.items td.amt { text-align: right; white-space: nowrap; }' +
    'table.items td { padding: 11px 0; border-bottom: 1px solid ' + INVOICE_CONFIG.RULE + '; vertical-align: top; }' +
    '.sub { color: ' + INVOICE_CONFIG.GREY + '; font-size: 9.5pt; }' +
    'table.totals { margin-left: auto; margin-top: 10px; border-collapse: collapse; min-width: 230px; }' +
    'table.totals td { padding: 5px 0; font-size: 10.5pt; }' +
    'table.totals .amt { text-align: right; }' +
    'table.totals tr.grand td { border-top: 2px solid ' + INVOICE_CONFIG.DARK + '; padding-top: 9px; font-weight: bold; font-size: 12pt; }' +
    '.note { font-size: 9.5pt; color: ' + INVOICE_CONFIG.GREY + '; margin: 6px 0; }' +
    '.pay { margin-top: 30px; border-top: 1px solid ' + INVOICE_CONFIG.RULE + '; padding-top: 14px; }' +
    '.foot { margin-top: 34px; font-size: 9pt; color: ' + INVOICE_CONFIG.GREY + '; }' +
    '</style></head><body>' +

    '<div class="head">' +
      '<div class="title"><h1>' + m.title + '</h1><div class="num">' + m.invoiceNumber + '</div></div>' +
      '<div class="biz">' + escapeHtml_(INVOICE_CONFIG.BUSINESS_NAME) + '</div>' +
      '<div class="biz-meta">ABN ' + INVOICE_CONFIG.ABN + ' &nbsp;·&nbsp; ' + INVOICE_CONFIG.EMAIL + ' &nbsp;·&nbsp; ' + INVOICE_CONFIG.WEBSITE + '</div>' +
      '<div style="clear:both"></div>' +
    '</div>' +

    '<table class="parties"><tr>' +
      '<td><div class="lbl">Billed to</div>' + (buyerBlock || '—') + '</td>' +
      '<td><div class="lbl">Invoice date</div>' + fmtDate_(m.issueDate) +
        '<div class="lbl" style="margin-top:12px">Date of supply</div>' + fmtDate_(m.supplyDate) + '</td>' +
    '</tr></table>' +

    '<table class="items">' +
      '<tr><th>Description</th><th class="amt">Amount</th></tr>' +
      linesHtml +
    '</table>' +

    '<table class="totals">' + totalsHtml + '</table>' +
    gstNote +

    '<div class="pay">' + payment + '</div>' +

    '<div class="foot">Thanks for supporting a local business. If the delivery went well, a Google review genuinely helps: ' +
      INVOICE_CONFIG.REVIEW_LINK + '</div>' +

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
    '\nThanks again for the job. If you were happy with the delivery, a quick Google review makes a real difference to a small local business:\n' +
    INVOICE_CONFIG.REVIEW_LINK + '\n\n' +
    'Cheers,\n' + INVOICE_CONFIG.TRADING_SHORT + '\n' + INVOICE_CONFIG.WEBSITE;

  const opts = { attachments: [pdf], name: INVOICE_CONFIG.TRADING_SHORT };
  if (INVOICE_CONFIG.EMAIL_MODE === 'send') {
    GmailApp.sendEmail(model.customerEmail, subject, body, opts);
  } else {
    GmailApp.createDraft(model.customerEmail, subject, body, opts);
  }
  return true;
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
