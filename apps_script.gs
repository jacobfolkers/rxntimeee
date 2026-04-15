/**
 * Balloon Session logger
 *
 * Deploy this file as a Google Apps Script web app, then paste the deployed
 * web app URL into APPS_SCRIPT_URL inside bart.html.
 *
 * Required Google setup:
 * 1. Create a Google Sheet.
 * 2. Paste your spreadsheet ID below.
 * 3. Deploy this Apps Script as a web app with access for anyone who has the link.
 *
 * The client sends only human participant data. Simulated player data is kept
 * separate in the browser and should never be posted here.
 */

const SPREADSHEET_ID = "1RGzi8fbA3QsLoSUCsdQuIB3T_By2eSS-b2mScgaXfCM";
const SHEET_NAME = "BalloonSessionLog";

function doGet() {
  return ContentService
    .createTextOutput("Balloon Session logger is live. This endpoint is meant to receive POST requests from the experiment.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    const records = Array.isArray(payload.records) ? payload.records : [payload];
    const sheet = getOrCreateSheet_();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "timestamp",
        "recordType",
        "participantId",
        "displayName",
        "participantType",
        "round",
        "condition",
        "conditionLabel",
        "trial",
        "pumps",
        "popped",
        "points",
        "teamDelta",
        "teamScoreAfter",
        "isPractice",
        "questionnaireTotal",
        "questionnaireAverage",
        "sessionId",
        "adjustedPumps",
        "popRate",
        "averageCashOut",
        "totalScore",
        "teamMode"
      ]);
    }

    records.forEach((record) => {
      sheet.appendRow([
        new Date(),
        record.recordType || "",
        record.participantId || "",
        record.displayName || "",
        record.participantType || "human",
        record.round || "",
        record.condition || "",
        record.conditionLabel || "",
        record.trial || "",
        record.pumps || "",
        record.popped === true ? true : record.popped === false ? false : "",
        record.points || "",
        record.teamDelta || "",
        record.teamScoreAfter || "",
        record.isPractice === true ? true : record.isPractice === false ? false : "",
        record.questionnaireTotal || "",
        record.questionnaireAverage || "",
        record.sessionId || "",
        record.adjustedPumps || "",
        record.popRate || record.popRatePercent || "",
        record.averageCashOut || "",
        record.totalScore || "",
        record.teamMode === true ? true : record.teamMode === false ? false : ""
      ]);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, rowCount: records.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const existingSheet = spreadsheet.getSheetByName(SHEET_NAME);
  return existingSheet || spreadsheet.insertSheet(SHEET_NAME);
}
