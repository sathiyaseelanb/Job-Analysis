/**
 * Job Analysis Simulation — Score Logger
 * Saves each student's performance record to a Google Sheet.
 *
 * ── SETUP (one time, ~3 minutes) ──────────────────────────────────
 * 1. Create a Google Sheet. Note its name; leave the first tab as "Sheet1".
 * 2. In that Sheet: Extensions ▸ Apps Script.
 * 3. Delete any sample code, paste THIS entire file, and Save.
 * 4. Click Deploy ▸ New deployment.
 *      - Type:            Web app
 *      - Description:     Job Analysis Simulation logger
 *      - Execute as:      Me
 *      - Who has access:  Anyone
 *    Click Deploy, authorise when prompted (allow the permissions).
 * 5. Copy the "Web app URL" it gives you (ends in /exec).
 * 6. In the simulation, open Tweaks ▸ Data Logging and paste that URL
 *    into "Google Apps Script Web App URL".
 *
 * From then on, every time a student clicks "Print / Save PDF" on the
 * certificate (and every "Transmit Results"), a row is appended here.
 * ──────────────────────────────────────────────────────────────────
 */

// Column order written to the sheet. Edit labels here if you like.
var HEADERS = [
  'Server Time', 'Full Name', 'Register Number', 'Programme',
  'Role', 'Role Title', 'Score', 'Max', 'Band',
  'Dimension Scores', 'Reflection 1', 'Reflection 2', 'Reflection 3',
  'Saved Via', 'Phase Timestamps', 'Client Time'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // avoid two students writing the same row
  try {
    var data = {};
    if (e && e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);          // form-encoded (default)
    } else if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);       // raw JSON fallback
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    // Write a header row once.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.fullName || '',
      data.registerNumber || '',
      data.programName || '',
      data.roleSelected || '',
      data.roleTitle || '',
      data.calibrationScore != null ? data.calibrationScore : '',
      data.calibrationMax || 24,
      data.calibrationBand || '',
      data.dimensionScores || '',
      data.reflectionPrompt1 || '',
      data.reflectionPrompt2 || '',
      data.reflectionPrompt3 || '',
      data.savedVia || '',
      data.phaseTimestamps || '',
      data.timestamp || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Lets you open the /exec URL in a browser to confirm the app is live.
function doGet() {
  return ContentService.createTextOutput(
    'Job Analysis Simulation logger is running. Submissions are accepted via POST.'
  );
}
