// --- CONFIGURATION ---
var ROOT_DESTINATION_ID = '1DBzGWFTMqaTHvMAl8V0ipZddEwEqo5Sm';
var TEMPLATE_MAP = {
  "EMERGENCY SERVICE": "1RdyEYmSm8HIex0ZiXiQKZg8Z0fYYacrE",
  "MITIGATION": "14X47uUIlaOUlNQUcdtE3zfnVjJ-UpEDi",
  "REPAIRS_Core Claims": "14ncBTxNaJ2u0egFaY77YRFgMrYWkJ4v4",
  "REPAIRS_Large Loss": "12JxI_oJC-zbXkAw_rqxIUtLI9QIO3hD7"
};

var TRACKING_SHEET_ID = '1OD-IkT6S43YqMwqdPhNNC4kNsPTqcmK7JXKn3hF_aaw';
var TRACKING_SHEET_TAB = 'TBR Job Numbers';
var TBR_PREFIX = 'TBR-';


// ============================================================
// OLD FORM HANDLER (unchanged behavior)
// ============================================================

function onFormSubmit(e) {
  var itemResponses = e.response.getItemResponses();
  var projectName = "";
  var year = "";
  var category = "";

  for (var i = 0; i < itemResponses.length; i++) {
    var title = itemResponses[i].getItem().getTitle().trim();
    var answer = itemResponses[i].getResponse();

    Logger.log("Found Question: [" + title + "] Answer: [" + answer + "]");

    if (title === "Project Name") { projectName = answer; }
    else if (title === "Target Year (only fill out if not current year)") { year = answer; }
    else if (title === "Project Category") { category = answer; }
  }

  if (!year || year.trim() === "") {
    year = new Date().getFullYear().toString();
  }

  if (projectName && category) {
    createProjectStructure(projectName, year, category);
  } else {
    Logger.log("Error: Missing Project Name or Category. Found Project: " + projectName + ", Category: " + category);
  }
}


// ============================================================
// NEW LEAD FORM HANDLER
// ============================================================

function onLeadFormSubmit(e) {
  var answers = parseFormAnswers(e);
  var submittedAt = e.response.getTimestamp();

  var insurance = resolveInsurance(answers);

  var projectName = answers["Project Name"] || "";
  var category = answers["Project Category"] || "";
  var year = answers["Target Year (only fill out if not current year)"] || "";
  if (!year || year.trim() === "") {
    year = new Date().getFullYear().toString();
  }

  var createFolder = (answers["Create Project Folder?"] || "").trim() === "Yes";
  var createProjectNumRaw = (answers["Create Project Number?"] || "").trim();
  var createProjectNum = createProjectNumRaw === "Yes";
  var createJobCell = createProjectNumRaw ? createProjectNumRaw.toUpperCase() : "";

  var sheet = SpreadsheetApp.openById(TRACKING_SHEET_ID).getSheetByName(TRACKING_SHEET_TAB);

  var projectNumber = "";
  if (createProjectNum) {
    try {
      projectNumber = generateProjectNumber(projectName, sheet);
    } catch (err) {
      Logger.log("Project number generation failed: " + err);
    }
  }

  var projectFolderUrl = "";
  if (createFolder && projectName && category) {
    try {
      var folderName = projectNumber ? projectName + " (" + projectNumber + ")" : projectName;
      var newFolder = createProjectStructure(projectName, year, category, folderName);
      projectFolderUrl = newFolder.getUrl();
    } catch (err) {
      Logger.log("Folder creation failed for [" + projectName + "]: " + err);
    }
  }

  var rowValues = {
    "Create Job #?": createJobCell,
    "Project Number": projectNumber,
    "Project Name": projectName,
    "Project Category": category,
    "Project Address": answers["Project Address"] || "",
    "Contact Number": answers["Contact Number"] || "",
    "Email": answers["Email"] || "",
    "Contract Value": answers["Contract Value"] || "",
    "Estimated Cost": answers["Estimated Cost"] || "",
    "Job Completion Date": formatDateForSheet(answers["Job Completion Date"]),
    "Insurance": insurance,
    "Claim number": answers["Claim Number"] || "",
    "Date of Loss": formatDateForSheet(answers["Date of Loss"]),
    "Type of Loss": answers["Type of Loss"] || "",
    "Referral Name": answers["Referral Name"] || "",
    "Referral Number": answers["Referral Number"] || "",
    "Created Date": Utilities.formatDate(submittedAt, Session.getScriptTimeZone(), "MM/dd/yyyy"),
    "Project Folder": projectFolderUrl
  };

  try {
    appendLeadRow(sheet, rowValues);
  } catch (err) {
    Logger.log("Sheet append failed for [" + projectName + "]: " + err);
  }

  Logger.log("Lead submitted: project=[" + projectName + "] projectNumber=[" + projectNumber +
             "] folderUrl=[" + projectFolderUrl + "]");
}

function parseFormAnswers(e) {
  var itemResponses = e.response.getItemResponses();
  var answers = {};
  for (var i = 0; i < itemResponses.length; i++) {
    var title = itemResponses[i].getItem().getTitle().trim();
    answers[title] = itemResponses[i].getResponse();
  }
  return answers;
}

function resolveInsurance(answers) {
  var company = (answers["Insurance Company"] || "").trim();
  if (company === "Other") {
    return (answers["Other Insurance Company"] || "").trim();
  }
  return company;
}

function formatDateForSheet(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd/yyyy");
  }
  var str = String(value).trim();
  if (!str) return "";
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return iso[2] + "/" + iso[3] + "/" + iso[1];
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "MM/dd/yyyy");
  }
  return str;
}

function appendLeadRow(sheet, rowValues) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  var headerIndex = {};
  for (var i = 0; i < headers.length; i++) {
    var name = String(headers[i]).trim();
    if (name) headerIndex[name] = i;
  }

  var row = new Array(lastCol).fill("");
  for (var key in rowValues) {
    if (headerIndex.hasOwnProperty(key)) {
      row[headerIndex[key]] = rowValues[key];
    } else {
      Logger.log("Warning: sheet is missing column [" + key + "]");
    }
  }

  sheet.appendRow(row);

  var newRow = sheet.getLastRow();
  inheritRowValidation(sheet, newRow);
  refreshBandingExtent(sheet);
}

/** Last column on the tab that actually has a header in row 1. 0 if there are none. */
function lastHeaderColumn(sheet) {
  var n = sheet.getLastColumn();
  if (!n) return 0;
  var headers = sheet.getRange(1, 1, 1, n).getValues()[0];
  var last = 0;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim()) last = i + 1;
  }
  return last;
}

/**
 * Give the row we just appended the same dropdowns as the row above it, so the
 * YES/NO picklist on "Create Job #?", and the Project Manager and Job Status
 * lists, are there the moment a lead lands.
 *
 * Why copy instead of covering the column: data validation has no self-
 * maintaining equivalent of alternating colours. Banding is a single range rule
 * that re-flows when rows move; validation is a property of each individual
 * cell. So the choice is either a dropdown on every blank row to the bottom of
 * the sheet, or a new row inheriting them from its predecessor. This is the
 * second. A row inserted by hand, or by the Job-Mover, already inherits from the
 * row above; only an appended row does not, which is why this lives here.
 *
 * Copies the whole row rather than a named list of columns, so a dropdown added
 * to some new column later is picked up with no code change.
 *
 * Wrapped in try/catch because a missing dropdown must never cost a lead.
 */
function inheritRowValidation(sheet, row) {
  try {
    if (row < 3) return;              // row 2 is the first data row; nothing above it to copy
    var lastCol = lastHeaderColumn(sheet);
    if (!lastCol) return;

    var rules = sheet.getRange(row - 1, 1, 1, lastCol).getDataValidations();
    var any = false;
    for (var i = 0; i < rules[0].length; i++) {
      if (rules[0][i]) { any = true; break; }
    }
    if (!any) return;                 // the row above has no dropdowns; nothing to inherit

    sheet.getRange(row, 1, 1, lastCol).setDataValidations(rules);
  } catch (err) {
    Logger.log('Validation inherit skipped: ' + err);
  }
}

/**
 * Keep the tracking tab's alternating-colours banding covering the row we just
 * appended, so a lead added by this form is striped like every other row.
 *
 * Duplicated from the TBR Job Numbers sheet-bound script on purpose. This runs in
 * a different Apps Script project, and an Apps Script change does not reliably
 * fire another project's triggers, so that script's own refresh may never run for
 * a form submission. generateProjectNumber is duplicated across these two files
 * the same way.
 *
 * Never creates banding; setting that up is a one-time manual job per tab.
 * Wrapped in try/catch because a cosmetic problem must never cost a lead.
 */
function refreshBandingExtent(sheet) {
  try {
    var bandings = sheet.getBandings();
    if (bandings.length !== 1) return;

    var lastCol = lastHeaderColumn(sheet);
    var lastRow = sheet.getLastRow();
    if (lastCol < 2 || lastRow < 2) return;   // striping runs from B2; col A keeps its own fill

    var want = sheet.getRange(2, 2, lastRow - 1, lastCol - 1);
    if (bandings[0].getRange().getA1Notation() !== want.getA1Notation()) {
      bandings[0].setRange(want);
    }
  } catch (err) {
    Logger.log('Banding refresh skipped: ' + err);
  }
}


// ============================================================
// PROJECT NUMBER GENERATION
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
// FOLDER CREATION (existing logic; returns the new folder)
// ============================================================

function createProjectStructure(projectName, year, category, folderName) {
  var rootFolder = DriveApp.getFolderById(ROOT_DESTINATION_ID);
  var selectedTemplateId = TEMPLATE_MAP[category];

  if (!selectedTemplateId) {
    Logger.log("ERROR: No Template ID found for category: " + category);
    return null;
  }

  var templateFolder = DriveApp.getFolderById(selectedTemplateId);

  var categoryFolder = getOrCreateSubFolder(rootFolder, category);
  var yearFolder = getOrCreateSubFolder(categoryFolder, year);
  var newProjectFolder = yearFolder.createFolder(folderName || projectName);

  copyFolder(templateFolder, newProjectFolder, projectName);

  return newProjectFolder;
}

function getOrCreateSubFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

function copyFolder(source, target, projectName) {
  var folders = source.getFolders();
  var files = source.getFiles();

  while (files.hasNext()) {
    var file = files.next();
    var fileName = file.getName();
    fileName = projectName + " - " + fileName;
    file.makeCopy(fileName, target);
  }

  while (folders.hasNext()) {
    var subFolder = folders.next();
    var newSubFolder = target.createFolder(subFolder.getName());
    copyFolder(subFolder, newSubFolder, projectName);
  }
}
