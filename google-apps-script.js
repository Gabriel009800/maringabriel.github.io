/*
  Google Apps Script pentru Chestionarul COPE
  Gratuit: salvează rezultatele într-un Google Sheet și permite sincronizarea lor în pagina ta de rezultate.

  Pași:
  1. Creează un Google Sheet gol.
  2. Extensions → Apps Script.
  3. Șterge codul existent și lipește TOT acest cod.
  4. Schimbă SUBMIT_KEY și ADMIN_KEY.
  5. Deploy → New deployment → Web app.
  6. Execute as: Me.
  7. Who has access: Anyone.
  8. Copiază URL-ul de deploy în cope-config.js.

  IMPORTANT:
  - SUBMIT_KEY se pune și în cope-config.js.
  - ADMIN_KEY rămâne doar aici și în mintea ta. Nu o publica pe GitHub.
*/

const SHEET_NAME = "COPE_REZULTATE";
const SUBMIT_KEY = "schimba-aceasta-cheie-publica-de-trimitere";
const ADMIN_KEY = "schimba-aceasta-cheie-privata-admin";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "receivedAt",
      "id",
      "clientCode",
      "questionnaire",
      "testDate",
      "age",
      "contextNote",
      "payload"
    ]);
  }

  return sheet;
}

function jsonp_(callback, obj) {
  const safeCallback = String(callback || "callback").replace(/[^a-zA-Z0-9_$\.]/g, "");
  return ContentService
    .createTextOutput(safeCallback + "(" + JSON.stringify(obj) + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function doPost(e) {
  try {
    const params = e.parameter || {};
    if (params.submitKey !== SUBMIT_KEY) {
      return ContentService.createTextOutput("forbidden");
    }

    const payloadText = params.payload || "";
    const payload = JSON.parse(payloadText);

    if (!payload || payload.questionnaire !== "COPE" || !payload.responses || !payload.scaleScores) {
      return ContentService.createTextOutput("invalid payload");
    }

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      payload.id || "",
      payload.clientCode || "",
      payload.questionnaire || "",
      payload.testDate || "",
      payload.age || "",
      payload.contextNote || "",
      payloadText
    ]);

    return ContentService.createTextOutput("ok");
  } catch (err) {
    return ContentService.createTextOutput("error: " + err.message);
  }
}

function doGet(e) {
  const params = e.parameter || {};
  const callback = params.callback || "callback";

  try {
    if (params.action !== "list") {
      return jsonp_(callback, { ok: false, error: "Acțiune necunoscută." });
    }

    if (params.adminKey !== ADMIN_KEY) {
      return jsonp_(callback, { ok: false, error: "Cheie admin incorectă." });
    }

    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return jsonp_(callback, { ok: true, results: [] });
    }

    const values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const results = [];

    for (const row of values) {
      const payloadText = row[7];
      if (!payloadText) continue;
      try {
        const parsed = JSON.parse(payloadText);
        results.push(parsed);
      } catch (err) {
        // ignoră rândurile corupte
      }
    }

    results.sort(function(a, b) {
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    });

    return jsonp_(callback, { ok: true, results: results });
  } catch (err) {
    return jsonp_(callback, { ok: false, error: err.message });
  }
}
