/** @OnlyCurrentDoc */

// Contatore della solidarietà – Murales "Non voltarti dall'altra parte", Chieti
// Ogni scansione del QR aggiunge una riga (data e ora) al foglio "Scansioni".
//   ?mode=hit  → aggiunge una riga e restituisce il totale
//   ?mode=get  → restituisce il totale senza aggiungere nulla (anteprima)

var SHEET_NAME = "Scansioni";

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Data e ora", "Giorno"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e) {
  var mode = (e && e.parameter && e.parameter.mode) || "get";
  var sh = getSheet_();
  var total;

  if (mode === "hit") {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var now = new Date();
      sh.appendRow([now, Utilities.formatDate(now, "Europe/Rome", "yyyy-MM-dd")]);
      total = sh.getLastRow() - 1;
    } finally {
      lock.releaseLock();
    }
  } else {
    total = Math.max(0, sh.getLastRow() - 1);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ value: total }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Da eseguire una volta dall'editor per creare il foglio e autorizzare lo script.
function setup() {
  getSheet_();
}
