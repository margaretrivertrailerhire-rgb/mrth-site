/**
 * MR Trailer Hire & Landscape Delivery — booking form handler
 * ------------------------------------------------------
 * handleBooking_() is called from doPost() in Booking.js (the shared
 * entry point for both web app deployments) whenever the incoming
 * payload looks like a book.html submission. Writes into the Enquiries
 * sheet — the same one the quote form and the MRTH Job Manager phone app
 * use — via appendEnquiryRow_ (Booking.js), so a trailer-hire booking
 * shows up in the job list and can be invoiced from the phone the same
 * way a materials job can.
 *
 * The old "Bookings" tab (Customer Name | Phone | Email | Trailer |
 * Start Date | End Date | Display End Date | Notes | Status | Hire Type |
 * Amount | Payment Status | Lost reason) never held real data and is no
 * longer written to. setupBookingsValidation() below is left in place in
 * case that tab is still wanted for something, but nothing calls it
 * automatically any more.
 *
 * book.html's own "Hire Type" field just distinguishes which half of the
 * form was used ("Trailer Hire" vs "Materials Order") — it is NOT the
 * same as the Pricing tab's Hire Type (24hr / Weekend / Tradie), which
 * this form never asks for. That still has to be worked out and noted
 * manually.
 */

var BOOKINGS_SHEET = "Bookings";

/** book.html's trailer dropdown uses marketing names; the Pricing tab
 *  (and the invoice's Payment Link lookup) uses its own short names. The
 *  day rates match 1:1, so this is just a label difference — normalising
 *  it here means a trailer-hire job from the website can still match a
 *  Pricing row once the rate type is noted. */
var TRAILER_NAME_MAP = {
  "8×5 Compact": "8x5 Tipper",
  "10×5 Medium": "10x5 Tipper",
  "10×6 Large": "10x6 Tipper"
};
function normaliseTrailerName_(raw) {
  raw = String(raw || "").trim();
  return TRAILER_NAME_MAP[raw] || raw;
}

/** Booking-form submissions (trailer hire / materials order via book.html).
 *  Called from the shared doPost() in Booking.js — data is already parsed. */
function handleBooking_(data) {
  var isTrailer = data["Hire Type"] === "Trailer Hire";
  var start = data["Start Date"] || "";
  var end = data["End Date"] || "";

  appendEnquiryRow_({
    "Received": new Date(),
    "Status": "New",
    "Type": isTrailer ? "Trailer hire" : "Materials delivery",
    "Name": data["Customer Name"] || "",
    "Phone": data["Phone"] ? "'" + data["Phone"] : "",
    "Email": data["Email"] || "",
    "Material": isTrailer ? normaliseTrailerName_(data["Trailer"]) : "",
    "Timeframe": isTrailer ? (start + (end ? " to " + end : "")) : start,
    "Notes": data["Notes"] || "",
    "Start Date": start,
    "End Date": isTrailer ? end : "",
    "Source": "Website"
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Kept for the old Bookings tab, only if you still want its dropdowns —
 * nothing writes to that tab any more, so this has no effect on the live
 * booking flow. Safe to ignore or delete later.
 */
function setupBookingsValidation() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOOKINGS_SHEET);
  if (!sheet) return;
  var maxRows = Math.max(sheet.getMaxRows(), 500);

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["New", "Pending", "Confirmed", "Delivered", "Lost"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 9, maxRows - 1, 1).setDataValidation(statusRule); // column I

  var hireTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Trailer Hire", "Materials Order", "Tradie", "24hr"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 10, maxRows - 1, 1).setDataValidation(hireTypeRule); // column J

  if (sheet.getRange(1, 13).getValue() !== "Lost reason") {
    sheet.getRange(1, 13).setValue("Lost reason"); // column M header
  }

  var lostReasonRange = sheet.getRange(2, 13, maxRows - 1, 1); // column M
  var flagMissingReason = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($I2="Lost", $M2="")')
    .setBackground("#f4cccc")
    .setRanges([lostReasonRange])
    .build();
  var rules = sheet.getConditionalFormatRules().filter(function (r) {
    return r.getRanges().every(function (rg) { return rg.getColumn() !== 13; });
  });
  rules.push(flagMissingReason);
  sheet.setConditionalFormatRules(rules);

  Logger.log("Bookings validation + Lost reason column set up.");
}
