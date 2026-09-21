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

/**
 * Map trimmed header text to its 0-based column index.
 * If a tab has the same header twice, the FIRST one wins, so a stray duplicate
 * column on the right can never hijack lookups from the real column.
 */
function buildHeaderIndex(headers) {
  var idx = {};
  for (var i = 0; i < headers.length; i++) {
    var name = String(headers[i]).trim();
    if (name && !Object.prototype.hasOwnProperty.call(idx, name)) idx[name] = i;
  }
  return idx;
}

/** 1-based column number to its A1 letter. 1 -> A, 22 -> V, 27 -> AA. */
function columnLetter_(n) {
  var s = '';
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - r - 1) / 26;
  }
  return s;
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
 * When the Job Status cell on the "TBR Job Numbers" sheet is set to
 * "Complete" or "JNS (Job Not Sold)", ask for confirmation and move the whole
 * job row to the matching sheet, freezing calculated values.
 *
 * THE RULE: an archive tab's header row must be identical to the main tab's,
 * same names in the same order. The whole row is then copied straight across,
 * position for position.
 *
 * The code checks that rule before every move and REFUSES the move if it is
 * broken, naming the first columns that disagree. Nothing is copied and nothing
 * is deleted; the status cell is put back the way it was. Whoever added a
 * column to the main tab adds it to the archive tabs in the same position, then
 * sets the status again. It never rewrites an archive tab's headers on its own.
 *
 * This is deliberate. Copying by position without checking is what silently
 * misfiles data: insert one column on the main tab, forget the archive, and
 * every value after that point lands one column off with no error at all.
 *
 * Only the status column is located by header text rather than position, so
 * that adding a column never stops the trigger from firing.
 *
 * Requires an installable "On edit" trigger pointed at onEditInstallable.
 */

const CFG = {
  SOURCE_SHEET: 'TBR Job Numbers',
  STATUS_HEADER: 'Job Status',      // header text of the status column (trimmed, exact match)
  NAME_HEADER: 'Project Name',      // a row with this blank is not a job and is never moved
  NUMBER_HEADER: 'Project Number',  // used only to re-check row identity after the dialog
  FIRST_DATA_ROW: 2,                // row 1 is the header
  ROUTES: {
    'Complete': { sheet: 'TBR - Completed', label: 'TBR – Completed' },
    'JNS (Job Not Sold)': { sheet: 'TBR - JNS', label: 'TBR – JNS' }
  }
};

/**
 * Row striping. The stripes themselves are Sheets' native alternating colours,
 * set up BY HAND once per tab, because Sheets re-flows them on every insert and
 * delete while painted-on fills cannot. This script never creates that banding.
 * All it does is keep the banded RANGE covering exactly the rows that hold data,
 * so a row added at the bottom falls inside it and a deleted one does not.
 *
 * Striping starts at column B. Column A carries a green/red YES-NO fill, and a
 * painted fill always covers banding, so column A cannot show a stripe without
 * losing its colour.
 */
const STRIPE = {
  FIRST_COL: 2,   // column B
  FIRST_ROW: 2    // row 1 is the header and is never striped
};

/**
 * Totals row. Created BY HAND, found by its label in column A so that a stray
 * note below it cannot break the insert position. On a tab that has one, a moved
 * job is inserted ABOVE it; on a tab that does not, the job is appended at the
 * end exactly as before.
 */
const TOTALS = {
  LABEL: 'TOTAL',
  SUM_HEADERS: ['Contract Value', 'Estimated Cost', 'Actual Cost']
};

/**
 * Installable onEdit handler. Runs on every edit; exits fast unless the edit is
 * a single-cell change on the source sheet, in the column whose header is
 * CFG.STATUS_HEADER, to a value that maps to a route. Anything else is ignored.
 */
function onEditInstallable(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CFG.SOURCE_SHEET) return;
  const row = e.range.getRow();
  if (row < CFG.FIRST_DATA_ROW) return;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;

  // One cell read: is the edited column the status column? Header lookup, not a fixed letter.
  const col = e.range.getColumn();
  const header = String(sheet.getRange(1, col).getValue()).trim();
  if (header !== CFG.STATUS_HEADER) return;

  const newValue = String(e.value || '').trim();   // e.value is undefined when the cell is cleared
  if (!newValue || !Object.prototype.hasOwnProperty.call(CFG.ROUTES, newValue)) return;

  moveJob_(sheet, row, col, newValue, e.oldValue);
}

/**
 * Installable "On change" trigger. Keeps the striped range in step when somebody
 * inserts or deletes rows by hand. Install it once from the Triggers page,
 * alongside the existing On edit trigger.
 */
function onChangeInstallable(e) {
  if (!e || (e.changeType !== 'INSERT_ROW' && e.changeType !== 'REMOVE_ROW')) return;
  refreshAllBandingExtents_(SpreadsheetApp.getActive());
}

/** Refresh the striped range on the main tab and on every archive tab. */
function refreshAllBandingExtents_(ss) {
  const names = [CFG.SOURCE_SHEET];
  for (const key in CFG.ROUTES) names.push(CFG.ROUTES[key].sheet);
  for (let i = 0; i < names.length; i++) {
    const sh = ss.getSheetByName(names[i]);
    if (sh) refreshBandingExtent_(sh);
  }
}

/* ------------------------------------------------------------------ */
/* UI wrapper: confirm, guard, revert                                  */
/* ------------------------------------------------------------------ */

/**
 * Ask before moving. Guards empty rows, reverts the dropdown on "No", re-checks
 * the row after the dialog closes, and reports errors without ever deleting a row.
 * @param {Sheet}  sheet       the source sheet
 * @param {number} row         1-based row that was edited
 * @param {number} statusCol   1-based column of the status cell that was edited
 * @param {string} statusValue "Complete" or "JNS (Job Not Sold)"
 * @param {*}      oldValue    the previous status value (for revert)
 */
function moveJob_(sheet, row, statusCol, statusValue, oldValue) {
  const route = CFG.ROUTES[statusValue];
  const ui = SpreadsheetApp.getUi();
  const statusCell = sheet.getRange(row, statusCol);

  // Empty-row guard: no Project Name means this is not a job row.
  const before = readRowIdentity_(sheet, row);
  if (!before.name) {
    revertStatus_(statusCell, oldValue);
    ui.alert('This row has no ' + CFG.NAME_HEADER + ', so it was not moved.');
    return;
  }

  const dest = sheet.getParent().getSheetByName(route.sheet);
  if (!dest) {
    revertStatus_(statusCell, oldValue);
    ui.alert('The tab "' + route.sheet + '" was not found, so nothing was moved.');
    return;
  }

  // Header guard, checked BEFORE asking, so nobody confirms a move that then fails.
  // performMove_ checks again for itself; this one exists only for the better message.
  const check = compareHeaders_(sheet, dest);
  if (check.problems.length) {
    revertStatus_(statusCell, oldValue);
    ui.alert('Columns do not match. Job not moved.',
             headerMismatchMessage_(route.sheet, check.problems),
             ui.ButtonSet.OK);
    return;
  }

  const answer = ui.alert(
    'Move this job?',
    'Move "' + before.name + '" to ' + route.label + '?\n\nThis removes it from "' +
      CFG.SOURCE_SHEET + '".',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) {
    revertStatus_(statusCell, oldValue);
    return;
  }

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(15000)) {
    revertStatus_(statusCell, oldValue);
    ui.alert('The sheet is busy; nothing was moved. Set the status again to retry.');
    return;
  }
  try {
    // The dialog may have sat open while someone else moved or deleted a row above,
    // which would shift this row number onto a different job. Re-check before touching anything.
    const after = readRowIdentity_(sheet, row);
    if (after.name !== before.name || after.number !== before.number || after.status !== statusValue) {
      ui.alert('The sheet changed while the dialog was open; nothing was moved.\n\n' +
               'Find the job, set its status to something else, then back to "' +
               statusValue + '" to retry.');
      return;
    }

    performMove_(sheet, row, statusValue);
    SpreadsheetApp.getActiveSpreadsheet().toast('Job moved to ' + route.label + ' ✅');
  } catch (err) {
    ui.alert('Move failed. The job is still on "' + CFG.SOURCE_SHEET + '".\n\n' + err.message);
  } finally {
    lock.releaseLock();
  }
}

/** Restore the status cell to its previous value (clear it if there was none). */
function revertStatus_(statusCell, oldValue) {
  if (oldValue === undefined || oldValue === null) {
    statusCell.clearContent();
  } else {
    statusCell.setValue(oldValue);
  }
}

/** Read the fields that identify a job row, by header name. A missing header reads as ''. */
function readRowIdentity_(sheet, row) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = buildHeaderIndex(headers);
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const pick = function (name) {
    return idx[name] === undefined ? '' : String(values[idx[name]] || '').trim();
  };
  return {
    name: pick(CFG.NAME_HEADER),
    number: pick(CFG.NUMBER_HEADER),
    status: pick(CFG.STATUS_HEADER)
  };
}

/* ------------------------------------------------------------------ */
/* Header guard                                                        */
/* ------------------------------------------------------------------ */

/**
 * Compare the source header row to an archive tab's header row, position by
 * position, up to the last NAMED column on the source. Anything past that
 * (trailing blank headers on either tab, extra archive columns) is ignored.
 * Comparison is on trimmed text and is case-sensitive.
 * @return {{width: number, problems: string[]}} width is the number of columns a
 *         move would copy; problems is empty when the tabs line up.
 */
function compareHeaders_(srcSheet, destSheet) {
  const srcLast = srcSheet.getLastColumn();
  const srcHeaders = srcLast ? srcSheet.getRange(1, 1, 1, srcLast).getValues()[0] : [];

  let width = 0;
  for (let i = 0; i < srcHeaders.length; i++) {
    if (String(srcHeaders[i]).trim()) width = i + 1;
  }
  if (!width) throw new Error('"' + srcSheet.getName() + '" has no column headers in row 1.');

  const destLast = destSheet.getLastColumn();
  const destHeaders = destLast ? destSheet.getRange(1, 1, 1, destLast).getValues()[0] : [];

  const problems = [];
  for (let i = 0; i < width; i++) {
    const want = String(srcHeaders[i]).trim();
    const got = String(destHeaders[i] === undefined ? '' : destHeaders[i]).trim();
    if (got === want) continue;
    problems.push('Column ' + columnLetter_(i + 1) + ': this tab says ' +
                  (got ? '"' + got + '"' : '(blank)') + ', expected "' + want + '"');
  }
  return { width: width, problems: problems };
}

/**
 * Wording for a refused move. Only the first few disagreements are listed: one
 * inserted column shifts everything after it, and the first line is the one that
 * tells you where to look.
 */
function headerMismatchMessage_(destName, problems) {
  const SHOW = 4;
  let body = problems.slice(0, SHOW).join('\n');
  if (problems.length > SHOW) {
    body += '\n...and ' + (problems.length - SHOW) + ' more column(s) after that.';
  }
  return 'The columns on "' + destName + '" no longer match "' + CFG.SOURCE_SHEET + '":\n\n' +
         body + '\n\n' +
         'Nothing was moved and nothing was deleted.\n\n' +
         'Fix the header row on "' + destName + '" so it is identical to "' + CFG.SOURCE_SHEET +
         '", same names in the same order, then set Job Status again.';
}

/* ------------------------------------------------------------------ */
/* Row striping and the totals row                                     */
/* ------------------------------------------------------------------ */

/** Last column on a tab that actually has a header in row 1. 0 if there are none. */
function lastHeaderColumn_(sheet) {
  const n = sheet.getLastColumn();
  if (!n) return 0;
  const headers = sheet.getRange(1, 1, 1, n).getValues()[0];
  let last = 0;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim()) last = i + 1;
  }
  return last;
}

/**
 * Stretch or shrink this tab's existing alternating-colours banding so it covers
 * exactly the data rows. Idempotent, and a no-op when the range is already right.
 *
 * Deliberately does NOT create banding. Setting that up is a one-time manual job,
 * so a tab without it is reported in the log and otherwise left alone.
 * @return {boolean} true when the range was actually changed.
 */
function refreshBandingExtent_(sheet) {
  const bandings = sheet.getBandings();
  if (bandings.length !== 1) {
    Logger.log('Striping: "' + sheet.getName() + '" has ' + bandings.length +
               ' banding range(s), expected exactly 1. Left alone. Apply ' +
               'Format > Alternating colours to that tab by hand.');
    return false;
  }

  const lastCol = lastHeaderColumn_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastCol < STRIPE.FIRST_COL || lastRow < STRIPE.FIRST_ROW) return false;

  const want = sheet.getRange(STRIPE.FIRST_ROW, STRIPE.FIRST_COL,
                              lastRow - STRIPE.FIRST_ROW + 1,
                              lastCol - STRIPE.FIRST_COL + 1);
  if (bandings[0].getRange().getA1Notation() === want.getA1Notation()) return false;
  bandings[0].setRange(want);
  return true;
}

/**
 * Row number of this tab's totals row, or 0 when it has none. Scans column A
 * upward from the bottom, since that is where the totals row lives.
 */
function findTotalsRow_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const colA = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    if (String(colA[i][0]).trim().toUpperCase() === TOTALS.LABEL) return i + 2;
  }
  return 0;
}

/**
 * Rewrite the totals row's sums so they span exactly row 2 to the row above it.
 *
 * This is necessary, not cosmetic. A hand-written =SUM(J2:J9) does NOT grow when
 * a row is inserted at row 10, so without this the newest job would be left out
 * of the total. Rewriting the bounds every time makes that drift impossible.
 */
function updateTotals_(sheet, totalsRow) {
  const lastCol = sheet.getLastColumn();
  if (!lastCol) return;
  const idx = buildHeaderIndex(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
  const lastDataRow = totalsRow - 1;

  for (let i = 0; i < TOTALS.SUM_HEADERS.length; i++) {
    const name = TOTALS.SUM_HEADERS[i];
    if (idx[name] === undefined) {
      Logger.log('Totals: no "' + name + '" column on "' + sheet.getName() + '"; skipped.');
      continue;
    }
    const cell = sheet.getRange(totalsRow, idx[name] + 1);
    if (lastDataRow < 2) {
      cell.setValue(0);
    } else {
      const letter = columnLetter_(idx[name] + 1);
      cell.setFormula('=SUM(' + letter + '2:' + letter + lastDataRow + ')');
    }
  }
}

/* ------------------------------------------------------------------ */
/* Core move (UI-free, called by the wrapper)                          */
/* ------------------------------------------------------------------ */

/**
 * Move `row` on srcSheet to the archive tab mapped by statusValue.
 *
 * Refuses outright unless the two header rows are identical, so a row can never
 * be filed one column off. Then copies the whole row straight across by
 * position: formatting first, then FROZEN values (formulas resolve to their
 * results), and deletes the source row LAST. Every failure path throws before
 * the delete, so a failure can never lose a row; at worst it leaves the
 * original un-moved.
 * @return {{sheet: string, width: number}}
 */
function performMove_(srcSheet, row, statusValue) {
  const route = CFG.ROUTES[statusValue];
  if (!route) throw new Error('No route for status: ' + statusValue);

  const ss = srcSheet.getParent();
  const dest = ss.getSheetByName(route.sheet);
  if (!dest) throw new Error('Destination sheet not found: ' + route.sheet);

  // Authoritative check. moveJob_ already ran this for a friendlier message, but
  // performMove_ must never write a misaligned row even if called some other way.
  const check = compareHeaders_(srcSheet, dest);
  if (check.problems.length) throw new Error(headerMismatchMessage_(route.sheet, check.problems));

  const width = check.width;
  const srcRange = srcSheet.getRange(row, 1, 1, width);

  // Where does the job land? Above the totals row on a tab that has one, so the
  // totals stay at the bottom. Otherwise straight onto the end, as before.
  const totalsRow = findTotalsRow_(dest);
  let destRow;
  if (totalsRow) {
    dest.insertRowBefore(totalsRow);   // totals shifts down to totalsRow + 1
    destRow = totalsRow;
  } else {
    destRow = dest.getLastRow() + 1;
    if (dest.getMaxRows() < destRow) dest.insertRowsAfter(dest.getMaxRows(), 1);
  }
  const destRange = dest.getRange(destRow, 1, 1, width);

  // 1) Formatting, so the archived row looks identical (currency, dates, fills).
  srcRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  // 2) Frozen values: getValues() resolves formulas to their results.
  destRange.setValues(srcRange.getValues());
  SpreadsheetApp.flush(); // commit the destination write before deleting the source

  // 3) Delete the source row LAST; rows below shift up.
  srcSheet.deleteRow(row);

  // 4) Housekeeping, all of it after the delete because none of it can lose a
  //    row. The totals row sits one lower now that a job was inserted above it.
  if (totalsRow) updateTotals_(dest, totalsRow + 1);
  refreshBandingExtent_(dest);
  refreshBandingExtent_(srcSheet);

  return {
    sheet: route.sheet,
    width: width,
    row: destRow,
    totalsRow: totalsRow ? totalsRow + 1 : 0
  };
}
