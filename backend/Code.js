/**
 * Margaret River Trailer Hire — new booking email alert
 * -------------------------------------------------------
 * Emails you whenever a new row is added to the "Bookings" tab,
 * so a booking from the website doesn't sit unseen while you're
 * on shift.
 *
 * SETUP (one-time):
 * 1. Open the booking Google Sheet.
 * 2. Extensions → Apps Script.
 * 3. Delete anything in Code.gs and paste this whole file in.
 * 4. Change YOUR_EMAIL below to your real email address.
 * 5. Save (disk icon).
 * 6. Run the function "setup" once (select it from the dropdown
 *    at the top, then click Run). It'll ask for permissions —
 *    accept them (click through the "unsafe" warning, it's your
 *    own script).
 * 7. Done. From now on, any new booking row triggers an email
 *    within a few minutes.
 *
 * To stop the alerts later: Apps Script → clock icon on the left
 * (Triggers) → delete the trigger for "checkForNewBookings".
 */

var YOUR_EMAIL = "margaretrivertrailerhire@gmail.com";
var ALERT_SHEET_NAME = "Bookings";    // the tab with booking rows
var CHECK_EVERY_MINUTES = 5;          // how often to check for new rows

// One-time setup: creates the recurring trigger.
function setup() {
  // remove any existing triggers for this function first (avoid duplicates)
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "checkForNewBookings") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("checkForNewBookings")
    .timeBased()
    .everyMinutes(CHECK_EVERY_MINUTES)
    .create();

  // remember how many rows currently exist so we don't email about old bookings
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ALERT_SHEET_NAME);
  var props = PropertiesService.getScriptProperties();
  props.setProperty("lastRow", String(sheet.getLastRow()));

  Logger.log("Setup complete. Currently " + sheet.getLastRow() + " rows. Alerts will fire for new rows from here on.");
}

// Runs automatically every few minutes.
function checkForNewBookings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ALERT_SHEET_NAME);
  var props = PropertiesService.getScriptProperties();
  var lastRow = Number(props.getProperty("lastRow") || sheet.getLastRow());
  var currentLastRow = sheet.getLastRow();

  if (currentLastRow > lastRow) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var newRows = sheet.getRange(lastRow + 1, 1, currentLastRow - lastRow, sheet.getLastColumn()).getValues();

    newRows.forEach(function (row) {
      var lines = [];
      for (var i = 0; i < headers.length; i++) {
        if (row[i] !== "" && row[i] !== null) {
          lines.push(headers[i] + ": " + row[i]);
        }
      }
      var body = "New booking received on Margaret River Trailer Hire:\n\n" + lines.join("\n");
      MailApp.sendEmail(YOUR_EMAIL, "New trailer booking", body);
    });
  }

  props.setProperty("lastRow", String(currentLastRow));
}