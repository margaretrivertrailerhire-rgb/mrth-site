/**
 * MRTH WEBSITE BACKEND — unified entry point
 * Handles BOTH web app deployments from one place:
 *   - Quote form (index.html)  → "Enquiries" tab
 *   - Booking form (book.html) → "Enquiries" tab (logic in Booking Receiver.js)
 * Both forms land in the same Enquiries sheet — the one the MRTH Job
 * Manager phone app reads — so every job, however it comes in, shows up
 * in one list with one status lifecycle. The old "Bookings" tab never
 * held real data and is left as-is, just no longer written to.
 *
 * There is exactly one doPost/doGet for the whole project — Apps Script
 * shares one global scope across every file, so a second doPost here
 * would be a redeclaration error, not a second entry point. doPost
 * below looks at the shape of the incoming payload and routes to the
 * right handler.
 *
 * Rows are written by header NAME, not column position (see
 * appendEnquiryRow_ below) — the Job Manager app and this project each
 * independently guarantee their required columns exist, so either one
 * can add a column without the other silently writing into the wrong
 * cell.
 *
 * ── SETUP (once) ────────────────────────────────────────────────────────
 * 1. Open your bookings Google Sheet.
 * 2. Extensions → Apps Script. Make sure Booking.js, Booking Receiver.js
 *    and Code.js are all present (this project's 3 files).
 * 3. Change ALERT_EMAIL below to your address.
 * 4. Run testPost() and testPostBooking() once from the editor (▶ button)
 *    and approve the permissions prompt. Check the sheet + your inbox —
 *    you should see two test rows and a test email.
 * 6. Deploy → New deployment → type: Web app
 *      - Execute as:      Me
 *      - Who has access:  Anyone
 *    Copy the Web app URL and paste it into APPS_SCRIPT_URL in the
 *    homepage HTML (and/or BOOKING_URL in book.html).
 *
 * ── AFTER ANY CODE EDIT (important!) ────────────────────────────────────
 * Edits do NOT go live on an existing deployment URL automatically.
 * Deploy → Manage deployments → ✏️ edit the deployment → Version: "New
 * version" → Deploy. Same URL, new code. This project has TWO separate
 * deployments (one for each form) — both need redeploying for a fix to
 * take effect on both sites.
 * ────────────────────────────────────────────────────────────────────────
 */

const ALERT_EMAIL = "margaretrivertrailerhire@gmail.com";        // ← change me
const ENQUIRIES_SHEET = "Enquiries";

/** Columns this project guarantees exist on Enquiries. Order here only
 *  matters for a brand-new sheet — appendEnquiryRow_ always writes by
 *  header name, never by position, so an existing sheet's real column
 *  order (however the Job Manager app has grown it) is never disturbed. */
const REQUIRED_ENQUIRY_COLS = [
  "Received", "Status", "Type", "Name", "Phone", "Email",
  "Material", "Quantity", "Suburb", "Timeframe", "Access notes",
  "Quoted $", "Notes", "Carrier", "Load Type", "Delivery address",
  "Paid date", "Delivered date", "Invoice No", "Invoice Link",
  "Start Date", "End Date",
  "Source", "Customer type", "Lane", "Qty", "Unit", "Cost ex GST",
  "Sell ex GST", "Margin ex GST", "Lost reason", "Closed date", "Month"
];

/** Health check — visit the web app URL in a browser to confirm it's live. */
function doGet() {
  return ContentService.createTextOutput("MRTH backend OK");
}

/** Single entry point for both web app deployments. Routes by payload shape:
 *  the booking form (book.html) always sends "Customer Name"; the quote
 *  form (index.html) never does. */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (Object.prototype.hasOwnProperty.call(d, "Customer Name")) {
      return handleBooking_(d);
    }
    return handleEnquiry_(d);
  } catch (err) {
    // Log the failure somewhere visible instead of dying silently
    try {
      getSheet_("Errors", ["When", "Error", "Raw payload"]).appendRow([
        new Date(), String(err), e && e.postData ? e.postData.contents : "(no payload)"
      ]);
    } catch (ignore) {}
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: String(err) })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/** Quote-form submissions (materials delivery / trailer hire enquiries). */
function handleEnquiry_(d) {
  appendEnquiryRow_({
    "Received": new Date(),
    "Status": "New",
    "Type": d.enquiryType || "Enquiry",
    "Name": d.name || "",
    "Phone": d.phone ? "'" + d.phone : "",   // leading ' keeps leading zeros
    "Email": d.email || "",
    "Material": d.material || "",
    "Quantity": d.quantity || "",
    "Suburb": d.suburb || "",
    "Timeframe": d.timeframe || "",
    "Access notes": d.access || "",
    "Load Type": d.loadType || "",
    "Source": "Website",
  });
  sendAlert_(d);

  return ContentService.createTextOutput(
    JSON.stringify({ ok: true })
  ).setMimeType(ContentService.MimeType.JSON);
}

/** Gets a tab, creating it with the given headers if it doesn't exist yet.
 *  Used as-is for the simple, fixed-shape "Errors" log; Enquiries writes
 *  go through appendEnquiryRow_ instead, which is safe against column
 *  reordering. */
function getSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

/** The Enquiries tab, guaranteed to have every column this project needs
 *  — matching the pattern the Job Manager app uses on its side, so
 *  whichever script runs first self-heals the sheet for the other. */
function enquiriesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ENQUIRIES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ENQUIRIES_SHEET);
    sh.appendRow(REQUIRED_ENQUIRY_COLS);
    sh.getRange(1, 1, 1, REQUIRED_ENQUIRY_COLS.length).setFontWeight("bold");
    sh.setFrozenRows(1);
    return sh;
  }
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  REQUIRED_ENQUIRY_COLS.forEach(function (c) {
    if (headers.indexOf(c) === -1) {
      sh.getRange(1, headers.length + 1).setValue(c).setFontWeight("bold");
      headers.push(c);
    }
  });
  return sh;
}

function enquiryHeaderMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const map = {};
  headers.forEach(function (h, i) { map[h] = i; });
  return map;
}

/** Appends one row to Enquiries, placing each field by header name
 *  wherever that column actually lives — never by position. Fields not
 *  passed are left blank; fields with no matching column are ignored
 *  rather than shifting everything else along. */
function appendEnquiryRow_(fields) {
  const sh = enquiriesSheet_();
  const map = enquiryHeaderMap_(sh);
  const width = sh.getLastColumn();
  const row = new Array(width).fill("");
  Object.keys(fields).forEach(function (k) {
    if (map[k] !== undefined) row[map[k]] = fields[k];
  });
  sh.appendRow(row);
}

/** Email alert so you know within a minute, even on site. */
function sendAlert_(d) {
  const type = d.enquiryType || "Enquiry";
  const subject = "MRTH " + type + ": " +
    (d.material || "?") + " → " + (d.suburb || "?");

  const body =
    "New enquiry from the website\n" +
    "----------------------------------------\n" +
    "Type:       " + type + "\n" +
    "Name:       " + (d.name || "") + "\n" +
    "Phone:      " + (d.phone || "") + "\n" +
    "Email:      " + (d.email || "-") + "\n" +
    "Material:   " + (d.material || "") + "\n" +
    "Quantity:   " + (d.quantity || "") + "\n" +
    "Load type:  " + (d.loadType || "") + "\n" +
    "Suburb:     " + (d.suburb || "") + "\n" +
    "Timeframe:  " + (d.timeframe || "") + "\n" +
    "Access:     " + (d.access || "-") + "\n" +
    "----------------------------------------\n" +
    "Reply by SMS to lock it in quickly.";

  MailApp.sendEmail({
    to: ALERT_EMAIL,
    subject: subject,
    body: body,
    replyTo: d.email || ALERT_EMAIL,
    name: "MR Trailer Hire & Landscape Delivery"
  });
}

/** Run this once from the editor to test the quote-form pipeline
 *  (permissions, sheet write, email) without touching the website. */
function testPost() {
  const fake = {
    postData: {
      contents: JSON.stringify({
        enquiryType: "Materials delivery",
        name: "Test Customer",
        phone: "0400000000",
        email: "test@example.com",
        material: "Pine bark mulch",
        quantity: "4 m3",
        loadType: "Trailer load",
        suburb: "Witchcliffe",
        timeframe: "This week",
        access: "Tip on the left of the driveway",
        submitted: new Date().toISOString()
      })
    }
  };
  const out = doPost(fake);
  Logger.log(out.getContent());
}

/** Run this once from the editor to test the booking-form pipeline
 *  (trailer-hire shape) — checks it lands in Enquiries with the trailer
 *  name normalised and Start/End Date filled in. */
function testPostBooking() {
  const fake = {
    postData: {
      contents: JSON.stringify({
        "Customer Name": "Test Booking Customer",
        "Phone": "0400000000",
        "Email": "test@example.com",
        "Notes": "Test booking — safe to delete this row.",
        "Hire Type": "Trailer Hire",
        "Trailer": "10×5 Medium",
        "Start Date": "2026-09-10",
        "End Date": "2026-09-11"
      })
    }
  };
  const out = doPost(fake);
  Logger.log(out.getContent());
}
