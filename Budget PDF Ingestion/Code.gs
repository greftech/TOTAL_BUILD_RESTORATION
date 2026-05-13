/**
 * TBR Budget — bound script (thin wrapper)
 *
 * Hosts the menu, sidebar, `google.script.run` entry points, and the
 * onEdit trigger. All real work is delegated to the TBRBudgetTools
 * library.
 *
 * Library setup (one time, on this bound script):
 *   1. In the Apps Script editor, click + next to "Libraries" in the
 *      left sidebar.
 *   2. Paste the TBRBudgetTools library Script ID.
 *   3. Pick HEAD as the version (or a numbered version once the
 *      library is stable).
 *   4. Set identifier to `TBRBudgetTools`. Click Add.
 *
 * No GEMINI_API_KEY needed in this project — the library owns it.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('TBR Tools')
    .addItem('Ingest Scope PDF', 'showIngestSidebar')
    .addToUi();
}

function showIngestSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Ingest Scope PDF')
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---------- Sidebar callbacks (delegates to the library) ----------
// `google.script.run` only calls top-level functions in this bound
// project, so the library's functions get exposed via these wrappers.

function checkBudgetState() {
  return TBRBudgetTools.checkBudgetState();
}

function extractFromPdf(base64Pdf, fileName) {
  return TBRBudgetTools.extractFromPdf(base64Pdf, fileName);
}

function writeBudget(lineItems, grandTotal) {
  return TBRBudgetTools.writeBudget(lineItems, grandTotal);
}

// ---------- Format guardian ----------
// Simple trigger — fires on every user edit. Drag-moves and pastes
// inside the data region get their gridlines repaired immediately.
// No installation needed: defining a function literally named `onEdit`
// auto-registers it as a simple trigger when the script is saved.

function onEdit(e) {
  TBRBudgetTools.restoreGridlinesOnEdit(e);
}