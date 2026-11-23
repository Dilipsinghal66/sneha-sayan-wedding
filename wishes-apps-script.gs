/**
 * Wishes Wall backend for the Sneha & Sayan wedding invitation.
 * Paste this whole file into Extensions > Apps Script of a new Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 * Copy the web-app URL into WISHES_API in index.html.
 *
 * The sheet keeps every wish: Time | Name | Wish | Show.
 * To remove a wish, delete its row or type  hide  in column D.
 */
var MAX_NAME = 40, MAX_WISH = 300;

function sheet_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(["Time", "Name", "Wish", "Show"]);
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var sh = sheet_(), rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  var wishes = rows
    .filter(function (r) { return r[1] && r[2] && String(r[3]).trim().toLowerCase() !== "hide"; })
    .map(function (r) {
      return { t: new Date(r[0]).getTime() || 0, name: String(r[1]).slice(0, MAX_NAME), wish: String(r[2]).slice(0, MAX_WISH) };
    });
  return json_({ ok: true, wishes: wishes });
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents || "{}");
    var name = String(d.name || "").replace(/\s+/g, " ").trim().slice(0, MAX_NAME);
    var wish = String(d.wish || "").trim().slice(0, MAX_WISH);
    if (!name || !wish) return json_({ ok: false, error: "empty" });
    // a leading = + - @ would make the sheet treat the text as a formula
    if (/^[=+\-@]/.test(name)) name = "'" + name;
    if (/^[=+\-@]/.test(wish)) wish = "'" + wish;
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { sheet_().appendRow([new Date(), name, wish, "show"]); }
    finally { lock.releaseLock(); }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: "bad request" });
  }
}
