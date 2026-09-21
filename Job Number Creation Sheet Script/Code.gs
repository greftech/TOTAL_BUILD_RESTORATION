// ============================================================
// TBR JOB NUMBER TRIGGER — SELF-CONTAINED TEST BUNDLE
// ============================================================
// Paste this into a fresh Apps Script project bound to your test
// copy of the TBR Job Numbers sheet. Update TRACKING_SHEET_ID
// below to point at the COPY, not production.

// --- CONFIGURATION ---
var TRACKING_SHEET_ID = '1D6kppfGobZ42vRSN6xVTw7mS8YdmthXh1ICuudCmeFE';   // <-- update this
var TRACKING_SHEET_TAB = 'TBR Job Numbers';            // tab name in the sheet
var TBR_PREFIX = 'TBR-';
var ROOT_DESTINATION_ID = 'PASTE_TEST_ROOT_FOLDER_ID_HERE'; // only needed if you want to test backfill


// ============================================================
// EDIT TRIGGER — fires on YES flip
// ============================================================
/*
function onCreateJobFlagEdit(e) {
  if (!e || !e.range) return;

  if (e.source.getId() !== TRACKING_SHEET_ID) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== TRACKING_SHEET_TAB) return;

  var row = e.range.getRow();
  if (row < 2) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var headerIndex = buildHeaderIndex(headers);

  var createJobCol = headerIndex["Create Job #?"];
  if (createJobCol === undefined) return;
  if (e.range.getColumn() !== createJobCol + 1) return;

  var newVal = String(e.value || "").trim().toUpperCase();
  var oldVal = String(e.oldValue || "").trim().toUpperCase();
  if (newVal !== "YES") return;
  if (oldVal === "YES") return;

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    setStatus(sheet, row, headerIndex, "Error: could not acquire lock, try again");
    return;
  }

  try {
    processCreateJobFlag(sheet, row, headerIndex);
  } finally {
    lock.releaseLock();
  }
}
*/
// ============================================================
// DEBUG VERSION — replace your existing onCreateJobFlagEdit with this
// ============================================================
// After you save, flip the test row's flag to NO then back to YES.
// Then: Executions tab → click the latest run → read the log.
// Every guard now announces itself. Share the log output with me.

function onCreateJobFlagEdit(e) {
  Logger.log("=== onCreateJobFlagEdit fired ===");

  if (!e || !e.range) {
    Logger.log("EXIT 1: no event or no range");
    return;
  }

  var srcId = e.source ? e.source.getId() : "(no source)";
  Logger.log("Source spreadsheet ID: " + srcId);
  Logger.log("Expected TRACKING_SHEET_ID: " + TRACKING_SHEET_ID);
  if (srcId !== TRACKING_SHEET_ID) {
    Logger.log("EXIT 2: spreadsheet ID mismatch");
    return;
  }

  var sheet = e.range.getSheet();
  Logger.log("Sheet tab name: [" + sheet.getName() + "]");
  Logger.log("Expected TRACKING_SHEET_TAB: [" + TRACKING_SHEET_TAB + "]");
  if (sheet.getName() !== TRACKING_SHEET_TAB) {
    Logger.log("EXIT 3: tab name mismatch");
    return;
  }

  var row = e.range.getRow();
  Logger.log("Edit row: " + row);
  if (row < 2) {
    Logger.log("EXIT 4: header row");
    return;
  }

  var rows = e.range.getNumRows(), cols = e.range.getNumColumns();
  Logger.log("Range size: " + rows + " x " + cols);
  if (rows !== 1 || cols !== 1) {
    Logger.log("EXIT 5: multi-cell edit");
    return;
  }

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log("Headers: " + JSON.stringify(headers));
  var headerIndex = buildHeaderIndex(headers);

  var createJobCol = headerIndex["Create Job #?"];
  Logger.log("Create Job #? column index (0-based): " + createJobCol);
  if (createJobCol === undefined) {
    Logger.log("EXIT 6: 'Create Job #?' column header not found");
    return;
  }

  var editCol = e.range.getColumn();
  Logger.log("Edit column (1-based): " + editCol + " | Expected: " + (createJobCol + 1));
  if (editCol !== createJobCol + 1) {
    Logger.log("EXIT 7: edit was not in Create Job #? column");
    return;
  }

  var rawNew = e.value;
  var rawOld = e.oldValue;
  Logger.log("Raw e.value: " + JSON.stringify(rawNew) + " | Raw e.oldValue: " + JSON.stringify(rawOld));

  var newVal = String(rawNew || "").trim().toUpperCase();
  var oldVal = String(rawOld || "").trim().toUpperCase();
  Logger.log("Normalized transition: [" + oldVal + "] -> [" + newVal + "]");

  if (newVal !== "YES") {
    Logger.log("EXIT 8: normalized new value is not YES");
    return;
  }
  if (oldVal === "YES") {
    Logger.log("EXIT 9: old value was already YES");
    return;
  }

  Logger.log("All guards passed. Acquiring lock...");

  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    Logger.log("EXIT 10: could not acquire lock");
    setStatus(sheet, row, headerIndex, "Error: could not acquire lock, try again");
    return;
  }

  try {
    Logger.log("Lock acquired, calling processCreateJobFlag...");
    processCreateJobFlag(sheet, row, headerIndex);
    Logger.log("processCreateJobFlag returned");
  } finally {
    lock.releaseLock();
    Logger.log("Lock released");
  }
}

function processCreateJobFlag(sheet, row, headerIndex) {
  var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var projectName = String(rowData[headerIndex["Project Name"]] || "").trim();
  if (!projectName) {
    setStatus(sheet, row, headerIndex, "Error: Project Name missing");
    return;
  }

  var existingNumber = String(rowData[headerIndex["Project Number"]] || "").trim();
  if (existingNumber) {
    setStatus(sheet, row, headerIndex, "Already assigned: " + existingNumber);
    return;
  }

  var folderUrl = String(rowData[headerIndex["Project Folder"]] || "").trim();
  if (!folderUrl) {
    setStatus(sheet, row, headerIndex, "Error: Project Folder URL missing");
    return;
  }

  var folderId = extractDriveFolderId(folderUrl);
  if (!folderId) {
    setStatus(sheet, row, headerIndex, "Error: could not parse folder ID from URL");
    return;
  }

  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (err) {
    setStatus(sheet, row, headerIndex, "Error: folder not accessible (" + err + ")");
    return;
  }

  var projectNumber;
  try {
    projectNumber = generateProjectNumber(projectName, sheet);
  } catch (err) {
    setStatus(sheet, row, headerIndex, "Error: number generation failed (" + err + ")");
    return;
  }

  try {
    folder.setName(projectName + " (" + projectNumber + ")");
  } catch (err) {
    setStatus(sheet, row, headerIndex, "Error: folder rename failed (" + err + ")");
    return;
  }

  sheet.getRange(row, headerIndex["Project Number"] + 1).setValue(projectNumber);

  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd HH:mm");
  setStatus(sheet, row, headerIndex, projectNumber + " generated, folder renamed ✓ " + stamp);
}


// ============================================================
// PROJECT NUMBER GENERATION (mirrored from production script)
// ============================================================

function generateProjectNumber(projectName, sheet) {
  var lastName = String(projectName || "").split(",")[0] || "";
  var letters = lastName.toUpperCase().replace(/[^A-Z]/g, "");

  var existing = getExistingProjectNumbers(sheet);

  if (letters.length >= 3) {
    var attempts = [
      letters.substring(0, 3),
      letters.substring(0, 2) + letters.charAt(3)
    ];
    if (letters.length >= 5) {
      attempts.push(letters.substring(0, 2) + letters.charAt(4));
    }
    for (var i = 0; i < attempts.length; i++) {
      var candidate = TBR_PREFIX + attempts[i];
      if (!existing[candidate]) return candidate;
    }
  }

  var prefix = letters.substring(0, 2);
  var n = 1;
  while (true) {
    var candidate = TBR_PREFIX + prefix + n;
    if (!existing[candidate]) return candidate;
    n++;
  }
}

function getExistingProjectNumbers(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var col = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === "Project Number") { col = i + 1; break; }
  }
  if (col === -1) throw new Error("Project Number column not found");

  var lastRow = sheet.getLastRow();
  var existing = {};
  if (lastRow < 2) return existing;

  var values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (var r = 0; r < values.length; r++) {
    var v = String(values[r][0]).trim();
    if (v) existing[v] = true;
  }
  return existing;
}


// ============================================================
// BACKFILL — populate Project Folder URL on existing rows
// ============================================================

function backfillProjectFolderUrls() {
  var sheet = SpreadsheetApp.openById(TRACKING_SHEET_ID).getSheetByName(TRACKING_SHEET_TAB);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Sheet has no data rows; nothing to backfill.");
    return;
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerIndex = buildHeaderIndex(headers);

  var required = ["Project Folder", "Project Name", "Project Category", "Project Number"];
  for (var i = 0; i < required.length; i++) {
    if (headerIndex[required[i]] === undefined) {
      throw new Error("Missing required column: " + required[i]);
    }
  }

  var createdDateCol = headerIndex["Created Date"];
  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var rootFolder = DriveApp.getFolderById(ROOT_DESTINATION_ID);

  var matched = 0, missing = 0, skipped = 0;

  for (var r = 0; r < data.length; r++) {
    var rowNum = r + 2;
    var existingUrl = String(data[r][headerIndex["Project Folder"]] || "").trim();
    if (existingUrl) { skipped++; continue; }

    var projectName = String(data[r][headerIndex["Project Name"]] || "").trim();
    var category = String(data[r][headerIndex["Project Category"]] || "").trim();
    var projectNumber = String(data[r][headerIndex["Project Number"]] || "").trim();

    if (!projectName || !category) {
      Logger.log("Row " + rowNum + ": skipped (missing project name or category)");
      missing++;
      continue;
    }

    var year = "";
    if (createdDateCol !== undefined) {
      var createdRaw = data[r][createdDateCol];
      if (createdRaw) {
        var d = (Object.prototype.toString.call(createdRaw) === "[object Date]")
          ? createdRaw : new Date(createdRaw);
        if (!isNaN(d.getTime())) year = d.getFullYear().toString();
      }
    }
    if (!year) year = new Date().getFullYear().toString();

    var candidateNames = [];
    if (projectNumber) candidateNames.push(projectName + " (" + projectNumber + ")");
    candidateNames.push(projectName);

    var folder = findProjectFolder(rootFolder, category, year, candidateNames);
    if (folder) {
      sheet.getRange(rowNum, headerIndex["Project Folder"] + 1).setValue(folder.getUrl());
      Logger.log("Row " + rowNum + ": matched [" + folder.getName() + "]");
      matched++;
    } else {
      Logger.log("Row " + rowNum + ": NO MATCH for [" + projectName + "] in " + category + "/" + year);
      missing++;
    }
  }

  Logger.log("Backfill done. matched=" + matched + " missing=" + missing + " skipped=" + skipped);
}

function findProjectFolder(rootFolder, category, year, candidateNames) {
  var categoryFolders = rootFolder.getFoldersByName(category);
  if (!categoryFolders.hasNext()) return null;
  var categoryFolder = categoryFolders.next();

  var yearFolders = categoryFolder.getFoldersByName(year);
  if (!yearFolders.hasNext()) return null;
  var yearFolder = yearFolders.next();

  for (var i = 0; i < candidateNames.length; i++) {
    var match = yearFolder.getFoldersByName(candidateNames[i]);
    if (match.hasNext()) return match.next();
  }
  return null;
}


// ============================================================
// HELPERS
// ============================================================

function buildHeaderIndex(headers) {
  var idx = {};
  for (var i = 0; i < headers.length; i++) {
    var name = String(headers[i]).trim();
    if (name) idx[name] = i;
  }
  return idx;
}

function extractDriveFolderId(url) {
  if (!url) return null;
  var s = String(url);
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

function setStatus(sheet, row, headerIndex, message) {
  var col = headerIndex["Job # Status"];
  if (col === undefined) {
    Logger.log("Job # Status column not found; status: " + message);
    return;
  }
  sheet.getRange(row, col + 1).setValue(message);
}

/**
 * TBR Job-Mover
 * -------------
 * When Job Status (column H) on the "TBR Job Numbers" sheet is set to
 * "Complete" or "JNS (Job Not Sold)", ask for confirmation and move the whole
 * job row (columns A–V) to the matching sheet, freezing calculated values.
 *
 * Setup: paste this into Extensions > Apps Script, Save, reload the sheet,
 * then use the "⚙️ TBR Tools > Install job-mover" menu once. See INSTALL.md.
 */

const CFG = {
  SOURCE_SHEET: 'TBR Job Numbers',
  STATUS_COL: 8,            // column H
  FIRST_DATA_ROW: 2,       // row 1 is the header
  LAST_DATA_COL: 22,       // column V — last real data column
  PROJECT_NUM_COL: 2,      // column B — used for the empty-row guard
  ROUTES: {
    'Complete': { sheet: 'TBR - Completed', label: 'TBR – Completed' },
    'JNS (Job Not Sold)': { sheet: 'TBR - JNS', label: 'TBR – JNS' }
  }
};

/**
 * Installable onEdit handler. Runs on every edit; exits fast unless the edit is
 * a Job Status change (column H, row >= 2) on the source sheet whose new value
 * maps to a route (Complete / JNS). Anything else is ignored.
 */
function onEditInstallable(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.SOURCE_SHEET) return;
  if (e.range.getColumn() !== CFG.STATUS_COL) return;
  const row = e.range.getRow();
  if (row < CFG.FIRST_DATA_ROW) return;

  const newValue = e.value;                        // selected string, or undefined if cleared
  if (!newValue || !CFG.ROUTES[newValue]) return;  // not Complete/JNS → do nothing

  moveJob_(sheet, row, newValue, e.oldValue);
}

/* ------------------------------------------------------------------ */
/* UI wrapper: confirm, guard, revert                                  */
/* ------------------------------------------------------------------ */

/**
 * Ask before moving. Guards empty rows, reverts the dropdown on "No", and
 * reports errors without ever deleting a row.
 * @param {Sheet}  sheet       the source sheet
 * @param {number} row         1-based row that was edited
 * @param {string} statusValue "Complete" or "JNS (Job Not Sold)"
 * @param {*}      oldValue    the previous Job Status value (for revert)
 */
function moveJob_(sheet, row, statusValue, oldValue) {
  const route = CFG.ROUTES[statusValue];
  const ui = SpreadsheetApp.getUi();
  const statusCell = sheet.getRange(row, CFG.STATUS_COL);



  const answer = ui.alert(
    'Move this job?',
    'Move this job to ' + route.label + '?\n\nThis removes it from "' +
      CFG.SOURCE_SHEET + '".',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) {
    revertStatus_(statusCell, oldValue);
    return;
  }

  try {
    performMove_(sheet, row, statusValue);
    SpreadsheetApp.getActiveSpreadsheet().toast('Job moved to ' + route.label + ' ✅');
  } catch (err) {
    ui.alert('Move failed — nothing was changed.\n\n' + err.message);
  }
}

/** Restore the Job Status cell to its previous value (clear it if there was none). */
function revertStatus_(statusCell, oldValue) {
  if (oldValue === undefined || oldValue === null) {
    statusCell.clearContent();
  } else {
    statusCell.setValue(oldValue);
  }
}

/* ------------------------------------------------------------------ */
/* Core move (UI-free, called by the wrapper and the self-test)        */
/* ------------------------------------------------------------------ */

/**
 * Move columns A–V of `row` on srcSheet to the destination mapped by statusValue.
 * Copies formatting, writes FROZEN computed values (no live formulas), then
 * deletes the source row LAST. Throws before any delete if something fails, so a
 * failure can never lose a row (at worst it leaves an un-moved original).
 * @return {string} the destination sheet name
 */
function performMove_(srcSheet, row, statusValue) {
  const route = CFG.ROUTES[statusValue];
  if (!route) throw new Error('No route for status: ' + statusValue);

  const ss = srcSheet.getParent();
  const dest = ss.getSheetByName(route.sheet);
  if (!dest) throw new Error('Destination sheet not found: ' + route.sheet);

  const nCols = CFG.LAST_DATA_COL;
  const srcRange = srcSheet.getRange(row, 1, 1, nCols);
  const destRow = dest.getLastRow() + 1;
  const destRange = dest.getRange(destRow, 1, 1, nCols);

  // 1) Copy formatting only, so the archived row looks identical (currency, %, dates).
  srcRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  // 2) Write frozen values — getValues() resolves formulas to their results.
  destRange.setValues(srcRange.getValues());
  SpreadsheetApp.flush(); // commit the destination write before deleting the source

  // 3) Delete the source row LAST; rows below shift up.
  srcSheet.deleteRow(row);

  return route.sheet;
}
