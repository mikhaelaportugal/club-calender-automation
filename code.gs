// ---------- CONFIG ----------
const CALENDAR_ID = "6e9e33c6c2c8eb22b5d053153f50043e53ec6f89934813f782045c28dd0a330e@group.calendar.google.com"; // your Gmail calendar
const SHEET_NAME = "Form Responses 1";  // sheet tab name
const LOG_SHEET = "ScriptLogs";         // will be created automatically if missing
const FIXED_LOCATION = "CS Lab";        // fixed location (CS LAB) for all events
// ---------- END CONFIG ----------

function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Could not find sheet named: " + SHEET_NAME);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const lastRow = sheet.getLastRow();
    const rowValues = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    const get = (header) => {
      const idx = headers.indexOf(header);
      return idx >= 0 ? rowValues[idx] : "";
    };

    // Match your form headers
    const title = get("Event Title") || "No title";
    const dateVal = get("Date");
    const startVal = get("Start time");
    const endVal = get("End time");
    const other = get("Other") || "";
    const email = get("SFSU Email") || "";

    const eventStart = combineDateAndTime(dateVal, startVal);
    const eventEnd = combineDateAndTime(dateVal, endVal);

    const cal = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!cal) throw new Error("Could not open calendar with ID: " + CALENDAR_ID);

    const options = {
      description: `Other info: ${other}\nEmail: ${email}`,
      guests: email || undefined,
      location: FIXED_LOCATION
    };

    const ev = cal.createEvent(title, eventStart, eventEnd, options);

    logToSheet("SUCCESS", `Created event ${ev.getId()} titled '${title}' for ${eventStart} - ${eventEnd}`);

  } catch (err) {
    logToSheet("ERROR", err.message + (err.stack ? "\n" + err.stack : ""));
    throw err;
  }
}

// Utilities
function combineDateAndTime(dateVal, timeVal) {
  if (!dateVal || !timeVal) throw new Error("Missing date or time");

  // Handle Sheets storing time as Dec 30 1899
  if (dateVal instanceof Date && timeVal instanceof Date) {
    return new Date(
      dateVal.getFullYear(),
      dateVal.getMonth(),
      dateVal.getDate(),
      timeVal.getHours(),
      timeVal.getMinutes(),
      timeVal.getSeconds()
    );
  }

  // Fallback to parsing strings
  const combined = new Date(dateVal + " " + timeVal);
  if (isNaN(combined)) throw new Error("Invalid date/time: '" + dateVal + " " + timeVal + "'");
  return combined;
}

function logToSheet(level, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) logSheet = ss.insertSheet(LOG_SHEET);
  logSheet.appendRow([new Date(), level, message]);
}

// TEST HELPER
function testOnFormSubmit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  const fakeEvent = {
    range: sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()),
    values: sheet.getRange(lastRow, 1, 1, sheet.getLastColumn()).getValues()
  };
  onFormSubmit(fakeEvent);
}