const SPREADSHEET_ID = "12n37RjWkKFvWx0ZxrzKMwYtIbKJyALj9EKFIVvdd72g";

// WICHTIG: Ändere diesen Schlüssel auf etwas Eigenes, Geheimes.
// Nur wer diesen Schlüssel kennt, kann die Gästeliste (Namen, Kontakte) einsehen.
const ADMIN_KEY = "viki-sandra-2026-geheim";

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "list") {
    if (e.parameter.key !== ADMIN_KEY) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "Falscher Schlüssel." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    try {
      const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
      const values = sheet.getDataRange().getValues();

      const guests = values.map(function (row) {
        return {
          name: row[0] || "",
          contact: row[1] || "",
          timestamp: row[2] ? new Date(row[2]).toISOString() : "",
          rulesAccepted: row[3] || ""
        };
      }).filter(function (g) {
        return g.name && g.name !== "Name"; // Kopfzeile überspringen, falls vorhanden
      });

      return ContentService
        .createTextOutput(JSON.stringify({ guests: guests }))
        .setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: error.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput("Viki & Sandra Gästeliste funktioniert!");
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
  } catch (lockError) {
    return ContentService.createTextOutput("FEHLER: Server war zu beschäftigt, bitte nochmal versuchen.");
  }

  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];

    let name = "";
    let contact = "";
    let rulesAccepted = "";

    // Normal HTML form POST (auch von fetch() mit FormData)
    if (e && e.parameter) {
      name = e.parameter.name || "";
      contact = e.parameter.contact || "";
      rulesAccepted = e.parameter.rulesAccepted || "";
    }

    // JSON POST (für direkte/API-Tests)
    if (!name && e && e.postData && e.postData.contents) {
      try {
        const data = JSON.parse(e.postData.contents);
        name = data.name || "";
        contact = data.contact || "";
        rulesAccepted = data.rulesAccepted || "";
      } catch (jsonError) {
        // Ignorieren, falls es ein normales Formular-POST war.
      }
    }

    name = name.trim();

    if (!name) {
      return ContentService.createTextOutput("FEHLER: Name fehlt.");
    }

    sheet.appendRow([
      name,
      contact,
      new Date(),
      rulesAccepted
    ]);

    return ContentService.createTextOutput("OK");

  } catch (error) {
    return ContentService
      .createTextOutput("FEHLER: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function testAnmeldung() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  sheet.appendRow([
    "TEST GAST",
    "test@example.com",
    new Date(),
    "Ja"
  ]);
}
