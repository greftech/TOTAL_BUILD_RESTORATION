/**
 * TBR Doc Gen — bound script on the TBR Job Numbers spreadsheet
 *
 * Sidebar UI that generates the project documents (Contract, Work Authorization,
 * Mold Waiver, Certificate of Completion) from Doc templates via replaceText
 * merge, lands them in the project's Drive folder, and dual-logs every
 * generation event (admin sheet row + per-project Activity Log Doc line).
 *
 * SETUP (one time):
 *   1. Create an empty Google Sheet in the TBR admin folder; copy its ID.
 *   2. Paste the ID into CONFIG_SHEET_ID below; save the script.
 *   3. From the Apps Script editor's Run dropdown, run `initializeConfigSheet`
 *      once. It will scaffold 4 tabs in the admin sheet with headers and seed
 *      rows. After it runs, you never run it again.
 *   4. In the admin sheet, paste the real Template Doc IDs into the Doc Types
 *      tab, then populate Placeholder Map and Required Fields rows for each
 *      active doc type.
 *   5. Reload the TBR Job Numbers sheet; "TBR Docs → Open Sidebar" appears.
 */

// --- CONFIGURATION ---
var CONFIG_SHEET_ID = '';   // <-- paste admin Config spreadsheet ID here, then save.
var JOB_NUMBERS_TAB = 'TBR Job Numbers';

var DOC_TYPES_TAB        = 'Doc Types';
var PLACEHOLDER_MAP_TAB  = 'Placeholder Map';
var REQUIRED_FIELDS_TAB  = 'Required Fields';
var ACTIVITY_TAB         = 'TBR Document Activity';

// Header name of the column on TBR Job Numbers that holds the project's Drive
// folder URL. Default matches what the Form and Sheet scripts write/read today.
// Overridable via Script Property FOLDER_URL_COLUMN_NAME for installs that
// rename the column.
var FOLDER_URL_COLUMN_DEFAULT = 'Project Folder';

var TIMEZONE = 'America/New_York';
var LOCK_TIMEOUT_MS = 30000;


// =====================================================================
// Menu + sidebar entry
// =====================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TBR Docs')
    .addItem('Open Sidebar', 'showDocGenSidebar')
    .addToUi();
}

function showDocGenSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('TBR Docs')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}


// =====================================================================
// Sidebar callbacks — invoked by google.script.run
// =====================================================================

function getInitialContext() {
  var jobSheet = SpreadsheetApp.getActive().getSheetByName(JOB_NUMBERS_TAB);
  if (!jobSheet) throw new Error('Tab "' + JOB_NUMBERS_TAB + '" not found on this spreadsheet.');

  var projects = listProjects_(jobSheet);
  var activeRow = SpreadsheetApp.getActiveRange() ? SpreadsheetApp.getActiveRange().getRow() : 0;

  var defaultRow;
  if (projects.length === 0) {
    defaultRow = 0;
  } else if (activeRow >= 2 && projects.some(function (p) { return p.rowIndex === activeRow; })) {
    defaultRow = activeRow;
  } else {
    defaultRow = projects[0].rowIndex;
  }

  var validation = defaultRow ? validateProject(defaultRow) : null;
  return { projects: projects, defaultRow: defaultRow, validation: validation };
}

function validateProject(rowIndex) {
  if (!rowIndex) throw new Error('No row selected.');
  var jobSheet = SpreadsheetApp.getActive().getSheetByName(JOB_NUMBERS_TAB);
  if (!jobSheet) throw new Error('Tab "' + JOB_NUMBERS_TAB + '" not found on this spreadsheet.');

  var headers = jobSheet.getRange(1, 1, 1, jobSheet.getLastColumn()).getValues()[0];
  var headerIndex = buildHeaderIndex_(headers);
  var rowData = jobSheet.getRange(rowIndex, 1, 1, jobSheet.getLastColumn()).getValues()[0];

  var jobNumber = String(rowData[headerIndex['Project Number']] || '').trim();
  var clientName = String(rowData[headerIndex['Project Name']] || '').trim();
  var folderColumnName = getFolderColumnName_();
  var folderUrl = String(rowData[headerIndex[folderColumnName]] || '').trim();

  var folderId = extractDriveFolderId_(folderUrl);
  var invalidFolder = false;
  if (!folderUrl) {
    invalidFolder = true;
  } else if (!folderId) {
    invalidFolder = true;
  } else {
    try {
      Drive.Files.get(folderId, { supportsAllDrives: true, fields: 'id,name,mimeType,trashed' });
    } catch (err) {
      invalidFolder = true;
    }
  }

  var configSS = openConfigSpreadsheet_();
  var activeDocTypes = readActiveDocTypes_(configSS);
  var requiredFieldsByDocType = readRequiredFields_(configSS);
  var lastByDocType = readLastGenerationByDocType_(configSS, jobNumber);

  var missingFields = [];
  var missingSet = {};
  activeDocTypes.forEach(function (dt) {
    var required = requiredFieldsByDocType[dt.name] || [];
    required.forEach(function (col) {
      if (headerIndex[col] === undefined) {
        var msg = '(missing column on sheet) ' + col;
        if (!missingSet[msg]) { missingSet[msg] = true; missingFields.push(msg); }
        return;
      }
      var val = rowData[headerIndex[col]];
      if (val === '' || val === null || val === undefined) {
        if (!missingSet[col]) { missingSet[col] = true; missingFields.push(col); }
      }
    });
  });

  var docTypes = activeDocTypes.map(function (dt) {
    var last = lastByDocType[dt.name];
    return {
      name: dt.name,
      hasTemplate: !!dt.templateDocId,
      templateVersion: dt.templateVersion,
      alreadyGenerated: !!last,
      lastGeneratedAt: last ? formatTimestamp_(last.timestamp) : '',
      lastGeneratedBy: last ? last.user : ''
    };
  });

  var canGenerate = !invalidFolder && missingFields.length === 0 && docTypes.some(function (d) { return d.hasTemplate; });

  return {
    rowIndex: rowIndex,
    jobNumber: jobNumber,
    clientName: clientName,
    folderUrl: folderUrl,
    folderId: folderId,
    invalidFolder: invalidFolder,
    missingFields: missingFields,
    docTypes: docTypes,
    canGenerate: canGenerate
  };
}

function generateSelected(rowIndex, selectedDocTypes) {
  if (!rowIndex) throw new Error('No row selected.');
  if (!selectedDocTypes || selectedDocTypes.length === 0) throw new Error('No doc types selected.');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
    throw new Error('Another generation is in progress. Try again in a few seconds.');
  }

  try {
    var jobSheet = SpreadsheetApp.getActive().getSheetByName(JOB_NUMBERS_TAB);
    var headers = jobSheet.getRange(1, 1, 1, jobSheet.getLastColumn()).getValues()[0];
    var headerIndex = buildHeaderIndex_(headers);
    var rowData = jobSheet.getRange(rowIndex, 1, 1, jobSheet.getLastColumn()).getValues()[0];

    var jobNumber = String(rowData[headerIndex['Project Number']] || '').trim();
    var clientName = String(rowData[headerIndex['Project Name']] || '').trim();
    var folderUrl = String(rowData[headerIndex[getFolderColumnName_()]] || '').trim();
    var folderId = extractDriveFolderId_(folderUrl);
    if (!folderId) throw new Error('Project Folder URL on the row is missing or invalid.');

    var configSS = openConfigSpreadsheet_();
    var docTypeMap = {};
    readActiveDocTypes_(configSS).forEach(function (dt) { docTypeMap[dt.name] = dt; });
    var placeholderMap = readPlaceholderMap_(configSS);
    var user = getActiveUserEmail_();

    var docsFolderId = getOrCreateSubfolderId_(folderId, 'Documents');

    var results = [];
    selectedDocTypes.forEach(function (docTypeName) {
      var dt = docTypeMap[docTypeName];
      if (!dt) {
        results.push({ docType: docTypeName, ok: false, error: 'Doc type not active in Config sheet.' });
        return;
      }
      if (!dt.templateDocId) {
        var msg = 'No Template Doc ID set for "' + docTypeName + '" in Config.';
        appendActivityRow_(configSS, [new Date(), jobNumber, clientName, docTypeName, dt.templateVersion, user, '', '', 'error', msg]);
        results.push({ docType: docTypeName, ok: false, error: msg });
        return;
      }
      try {
        var mappings = placeholderMap[docTypeName] || [];
        var out = generateOneDoc_(dt, mappings, headerIndex, rowData, jobNumber, clientName, docsFolderId);
        appendActivityRow_(configSS, [new Date(), jobNumber, clientName, docTypeName, dt.templateVersion, user, out.docUrl, out.pdfUrl, 'success', '']);
        appendActivityLogDocLine_(folderId, jobNumber, clientName, docTypeName, user, out.pdfUrl, out.docUrl);
        results.push({ docType: docTypeName, ok: true, pdfUrl: out.pdfUrl, docUrl: out.docUrl });
      } catch (err) {
        var detail = (err && err.message) ? err.message : String(err);
        appendActivityRow_(configSS, [new Date(), jobNumber, clientName, docTypeName, dt.templateVersion, user, '', '', 'error', detail]);
        results.push({ docType: docTypeName, ok: false, error: detail });
      }
    });

    return { results: results, validation: validateProject(rowIndex) };
  } finally {
    lock.releaseLock();
  }
}


// =====================================================================
// Merge engine
// =====================================================================

function generateOneDoc_(docType, mappings, headerIndex, rowData, jobNumber, clientName, docsFolderId) {
  var now = new Date();
  var baseName = '(' + jobNumber + ') - (' + docType.name + ') - ' + formatFilenameTimestamp_(now);

  var copied = Drive.Files.copy({ name: baseName, parents: [docsFolderId] }, docType.templateDocId, { supportsAllDrives: true });
  var newDocId = copied.id;

  var doc = DocumentApp.openById(newDocId);
  var body = doc.getBody();
  mappings.forEach(function (m) {
    var raw = headerIndex[m.sourceColumn] !== undefined ? rowData[headerIndex[m.sourceColumn]] : '';
    var formatted = formatFieldValue_(raw, m.format);
    body.replaceText(escapeReplacePattern_(m.placeholder), formatted);
  });
  doc.saveAndClose();

  var pdfBlob = DriveApp.getFileById(newDocId).getAs('application/pdf').setName(baseName + '.pdf');
  var pdf = Drive.Files.create({ name: baseName + '.pdf', parents: [docsFolderId] }, pdfBlob, { supportsAllDrives: true });

  return {
    docUrl: 'https://docs.google.com/document/d/' + newDocId + '/edit',
    pdfUrl: 'https://drive.google.com/file/d/' + pdf.id + '/view'
  };
}

function formatFieldValue_(value, format) {
  if (value === '' || value === null || value === undefined) return '';
  var fmt = String(format || 'text').toLowerCase();
  if (fmt === 'date') {
    var d = (Object.prototype.toString.call(value) === '[object Date]') ? value : new Date(value);
    if (isNaN(d.getTime())) return String(value);
    return Utilities.formatDate(d, TIMEZONE, 'MMMM d, yyyy');
  }
  if (fmt === 'currency') {
    var n = Number(value);
    if (isNaN(n)) return String(value);
    return '$' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return String(value);
}

function escapeReplacePattern_(s) {
  // replaceText treats its first arg as a regex; escape regex specials.
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


// =====================================================================
// Dual logging
// =====================================================================

function appendActivityRow_(configSS, row) {
  var sheet = configSS.getSheetByName(ACTIVITY_TAB);
  if (!sheet) throw new Error('Config sheet is missing the "' + ACTIVITY_TAB + '" tab. Run initializeConfigSheet from the Apps Script editor.');
  sheet.appendRow(row);
}

function appendActivityLogDocLine_(folderId, jobNumber, clientName, docTypeName, user, pdfUrl, docUrl) {
  var logName = '(' + jobNumber + ') - (' + clientName + ') Activity Log';
  var docId = findOrCreateDocInFolder_(folderId, logName);

  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  var prefix = formatTimestamp_(new Date()) + ' — ' + docTypeName + ' generated by ' + user + '. ';
  var p = body.appendParagraph(prefix);

  var pdfText = p.appendText('[PDF]');
  var len1 = p.getText().length;
  pdfText.setLinkUrl(len1 - 5, len1 - 1, pdfUrl);

  p.appendText(' ');
  var docText = p.appendText('[Doc]');
  var len2 = p.getText().length;
  docText.setLinkUrl(len2 - 5, len2 - 1, docUrl);

  doc.saveAndClose();
}

function findOrCreateDocInFolder_(folderId, name) {
  var q = "'" + folderId + "' in parents and name = '" + escapeQueryString_(name) + "' and mimeType = 'application/vnd.google-apps.document' and trashed = false";
  var resp = Drive.Files.list({
    q: q,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    fields: 'files(id,name)',
    pageSize: 10
  });
  if (resp.files && resp.files.length > 0) return resp.files[0].id;

  var created = Drive.Files.create({
    name: name,
    parents: [folderId],
    mimeType: 'application/vnd.google-apps.document'
  }, null, { supportsAllDrives: true });
  return created.id;
}


// =====================================================================
// Config sheet readers
// =====================================================================

function openConfigSpreadsheet_() {
  if (!CONFIG_SHEET_ID) throw new Error('CONFIG_SHEET_ID is not set in Code.gs. Paste the admin Config spreadsheet ID and save.');
  try {
    return SpreadsheetApp.openById(CONFIG_SHEET_ID);
  } catch (err) {
    throw new Error('Could not open Config spreadsheet (' + CONFIG_SHEET_ID + '): ' + err.message);
  }
}

function readSheetAsObjects_(configSS, tabName) {
  var sheet = configSS.getSheetByName(tabName);
  if (!sheet) throw new Error('Config tab "' + tabName + '" not found. Run initializeConfigSheet from the Apps Script editor.');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var obj = {};
    var hasAny = false;
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      var v = values[r][c];
      obj[headers[c]] = v;
      if (v !== '' && v !== null && v !== undefined) hasAny = true;
    }
    if (hasAny) out.push(obj);
  }
  return out;
}

function readActiveDocTypes_(configSS) {
  return readSheetAsObjects_(configSS, DOC_TYPES_TAB)
    .filter(function (r) { return String(r['Active'] || '').trim().toLowerCase() === 'yes'; })
    .map(function (r) {
      return {
        name: String(r['Doc Type Name'] || '').trim(),
        templateDocId: String(r['Template Doc ID'] || '').trim(),
        templateVersion: String(r['Template Version'] || '').trim()
      };
    })
    .filter(function (r) { return r.name; });
}

function readPlaceholderMap_(configSS) {
  var rows = readSheetAsObjects_(configSS, PLACEHOLDER_MAP_TAB);
  var byDocType = {};
  rows.forEach(function (r) {
    var dt = String(r['Doc Type'] || '').trim();
    var ph = String(r['Placeholder'] || '').trim();
    var src = String(r['Source Column'] || '').trim();
    var fmt = String(r['Field Format'] || '').trim();
    if (!dt || !ph || !src) return;
    if (!byDocType[dt]) byDocType[dt] = [];
    byDocType[dt].push({ placeholder: ph, sourceColumn: src, format: fmt });
  });
  return byDocType;
}

function readRequiredFields_(configSS) {
  var rows = readSheetAsObjects_(configSS, REQUIRED_FIELDS_TAB);
  var byDocType = {};
  rows.forEach(function (r) {
    var dt = String(r['Doc Type'] || '').trim();
    var src = String(r['Source Column'] || '').trim();
    if (!dt || !src) return;
    if (!byDocType[dt]) byDocType[dt] = [];
    byDocType[dt].push(src);
  });
  return byDocType;
}

function readLastGenerationByDocType_(configSS, jobNumber) {
  if (!jobNumber) return {};
  var sheet = configSS.getSheetByName(ACTIVITY_TAB);
  if (!sheet) return {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};
  var values = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idx = buildHeaderIndex_(headers);
  var byDocType = {};
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][idx['Job Number']]).trim() !== jobNumber) continue;
    if (String(values[r][idx['Status']]).trim().toLowerCase() !== 'success') continue;
    var dt = String(values[r][idx['Doc Type']]).trim();
    var ts = values[r][idx['Timestamp']];
    var tsDate = (Object.prototype.toString.call(ts) === '[object Date]') ? ts : new Date(ts);
    var user = String(values[r][idx['Generated By']] || '').trim();
    var existing = byDocType[dt];
    if (!existing || tsDate.getTime() > existing.timestamp.getTime()) {
      byDocType[dt] = { timestamp: tsDate, user: user };
    }
  }
  return byDocType;
}


// =====================================================================
// One-time bootstrap — run from the Apps Script editor's Run dropdown.
// Intentionally NOT wired to the menu.
// =====================================================================

function initializeConfigSheet() {
  if (!CONFIG_SHEET_ID) throw new Error('Paste the admin Config spreadsheet ID into CONFIG_SHEET_ID at the top of Code.gs first.');
  var ss = SpreadsheetApp.openById(CONFIG_SHEET_ID);

  var summary = [];

  summary.push(ensureTab_(ss, DOC_TYPES_TAB,
    ['Doc Type Name', 'Template Doc ID', 'Template Version', 'Active'],
    [
      ['Contract', '', '1', 'yes'],
      ['Work Authorization', '', '1', 'yes'],
      ['Mold Waiver', '', '1', 'yes'],
      ['Certificate of Completion', '', '1', 'yes']
    ]));

  summary.push(ensureTab_(ss, PLACEHOLDER_MAP_TAB,
    ['Doc Type', 'Placeholder', 'Source Column', 'Field Format'],
    [
      ['Contract', '{{client_name}}', 'Project Name', 'text'],
      ['Contract', '{{property_address}}', 'Project Address', 'text'],
      ['Contract', '{{loss_date}}', 'Date of Loss', 'date'],
      ['Contract', '{{loss_type}}', 'Type of Loss', 'text'],
      ['Contract', '{{insurance_carrier}}', 'Insurance', 'text'],
      ['Contract', '{{claim_number}}', 'Claim number', 'text'],
      ['Contract', '{{job_number}}', 'Project Number', 'text']
    ]));

  summary.push(ensureTab_(ss, REQUIRED_FIELDS_TAB,
    ['Doc Type', 'Source Column'],
    [
      ['Contract', 'Project Name'],
      ['Contract', 'Project Address'],
      ['Contract', 'Date of Loss'],
      ['Contract', 'Type of Loss'],
      ['Contract', 'Insurance'],
      ['Contract', 'Claim number']
    ]));

  summary.push(ensureTab_(ss, ACTIVITY_TAB,
    ['Timestamp', 'Job Number', 'Client Name', 'Doc Type', 'Template Version', 'Generated By', 'Editable Doc URL', 'PDF URL', 'Status', 'Error Detail'],
    []));

  summary.forEach(function (line) { Logger.log(line); });
  Logger.log('Done. Open the Config spreadsheet to fill in Template Doc IDs and finish populating Placeholder Map / Required Fields for the other doc types.');
}

function ensureTab_(ss, tabName, headers, seedRows) {
  var existing = ss.getSheetByName(tabName);
  if (existing) {
    return tabName + ': left alone (already exists).';
  }
  var sheet = ss.insertSheet(tabName);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (seedRows && seedRows.length) {
    sheet.getRange(2, 1, seedRows.length, headers.length).setValues(seedRows);
  }
  sheet.setFrozenRows(1);
  return tabName + ': created with ' + (seedRows ? seedRows.length : 0) + ' seed row(s).';
}


// =====================================================================
// Helpers
// =====================================================================

function listProjects_(jobSheet) {
  var lastRow = jobSheet.getLastRow();
  if (lastRow < 2) return [];
  var headers = jobSheet.getRange(1, 1, 1, jobSheet.getLastColumn()).getValues()[0];
  var idx = buildHeaderIndex_(headers);
  var jobCol = idx['Project Number'];
  var nameCol = idx['Project Name'];
  if (jobCol === undefined || nameCol === undefined) {
    throw new Error('TBR Job Numbers sheet is missing "Project Number" or "Project Name" column.');
  }
  var values = jobSheet.getRange(2, 1, lastRow - 1, jobSheet.getLastColumn()).getValues();
  var out = [];
  for (var r = 0; r < values.length; r++) {
    var jn = String(values[r][jobCol] || '').trim();
    var cn = String(values[r][nameCol] || '').trim();
    if (jn && cn) out.push({ rowIndex: r + 2, jobNumber: jn, clientName: cn });
  }
  return out;
}

function getOrCreateSubfolderId_(parentFolderId, name) {
  var q = "'" + parentFolderId + "' in parents and name = '" + escapeQueryString_(name) + "' and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
  var resp = Drive.Files.list({
    q: q,
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    fields: 'files(id,name)',
    pageSize: 10
  });
  if (resp.files && resp.files.length > 0) return resp.files[0].id;
  var created = Drive.Files.create({
    name: name,
    parents: [parentFolderId],
    mimeType: 'application/vnd.google-apps.folder'
  }, null, { supportsAllDrives: true });
  return created.id;
}

function getFolderColumnName_() {
  var override = PropertiesService.getScriptProperties().getProperty('FOLDER_URL_COLUMN_NAME');
  return (override && override.trim()) || FOLDER_URL_COLUMN_DEFAULT;
}

function buildHeaderIndex_(headers) {
  var idx = {};
  for (var i = 0; i < headers.length; i++) {
    var name = String(headers[i]).trim();
    if (name) idx[name] = i;
  }
  return idx;
}

function extractDriveFolderId_(url) {
  if (!url) return null;
  var s = String(url);
  var m = s.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

function escapeQueryString_(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function getActiveUserEmail_() {
  var e = Session.getActiveUser().getEmail();
  if (e) return e;
  return Session.getEffectiveUser().getEmail() || '';
}

function formatTimestamp_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'MMM d, yyyy h:mm a');
}

function formatFilenameTimestamp_(date) {
  return Utilities.formatDate(date, TIMEZONE, 'yyyy-MM-dd HHmm');
}
