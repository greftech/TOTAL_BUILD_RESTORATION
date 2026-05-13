/**
 * TBR Budget Tools — Library
 *
 * Centralized logic for the TBR Budget PDF Ingestion Tool. Used as a
 * library by the bound script on the TBR Budget Template (and copies
 * thereof). Keeping the API key and core logic here means new TBR
 * projects inherit the tool automatically when copied from the template
 * — no per-project setup required.
 *
 * Setup (one-time, in this library project):
 *   Project Settings > Script properties
 *   Add: GEMINI_API_KEY = <your key>
 *
 * Sharing:
 *   Drive-share this script file with anyone whose bound scripts will
 *   call it. View access is sufficient. The API key in Script Properties
 *   is NOT visible to viewers — only the project owner can see it.
 *
 * Updating consumers:
 *   When bound scripts reference this library at HEAD, edits here go
 *   live immediately for all users on next call. For tighter control,
 *   create numbered versions and have consumers reference a specific
 *   version (Deploy > Manage deployments > Library).
 */

const GEMINI_MODEL = 'gemini-2.5-flash';
const GEMINI_API_URL =
  `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent`;

const BUDGET = {
  SHEET_NAME: 'Budget Sheet (Totals)',
  CATEGORY_COL: 'C',
  AMOUNT_COL: 'D',
  TOTAL_BUDGET_COL: 'E',
  TOTAL_BUDGET_RATIO: 0.7,                 // E{r} = D{r} * 0.7
  CATEGORY_START_ROW: 3,
  GRAND_TOTAL_CELL: 'B5',
  JOB_COSTS_CELL: 'B11',
  SUBTOTAL_COLS: ['D','E','F','G','H','I','J','K','L','M'],
  JOB_COSTS_COLS: ['F','G','H','I','J','K','L','M'],
  CLEAR_RANGE: 'C3:M50',
  TOTAL_LABEL: 'Totals',                   // Label written into C{subtotalRow}; whole row C–M bolded
};

/**
 * Reports whether the bound budget sheet already has data we'd be
 * overwriting. Called from the sidebar at open time.
 */
function checkBudgetState() {
  const sheet = getBudgetSheet_();
  const c3 = sheet.getRange(`${BUDGET.CATEGORY_COL}${BUDGET.CATEGORY_START_ROW}`).getValue();
  const d3 = sheet.getRange(`${BUDGET.AMOUNT_COL}${BUDGET.CATEGORY_START_ROW}`).getValue();
  const grandTotal = sheet.getRange(BUDGET.GRAND_TOTAL_CELL).getValue();
  return {
    hasExistingData: c3 !== '' || d3 !== '' || grandTotal !== '',
  };
}

/**
 * Sends the PDF to Gemini and returns extracted line items + grand total.
 * Descriptions are returned with any leading 3-letter category code
 * stripped (e.g. "CAB CABINETRY" → "CABINETRY").
 */
function extractFromPdf(base64Pdf, fileName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not set in the TBRBudgetTools library Script Properties.');
  }

  const prompt = `You are extracting data from a TBR (Total Build Restorations) construction scope-of-work / cost document.

Return a JSON object with two fields:

1. "lineItems" — an array of every individual category line item with a dollar amount. For each:
   - description: the line item text WITHOUT any leading 3-letter category code.
       Example: "CAB CABINETRY" → "CABINETRY"
       Example: "FCV FLOOR COVERING - VINYL" → "FLOOR COVERING - VINYL"
       Example: "FNC FINISH CARPENTRY / TRIMWORK" → "FINISH CARPENTRY / TRIMWORK"
       If a line has no leading 3-letter code, return the text as-is.
   - amount: the dollar value as a number (no $, no commas, no quotes).

   Skip the column header row (often "CAT Total" — CAT is the column label for category code, Total is the column label for the dollar amount).
   Skip footer rows: Subtotal, Material Sales Tax, Sales Tax, Overhead, Profit, and Total.
   Skip page metadata: name, date, page number, etc.

2. "grandTotal" — the final "Total" value at the bottom of the document. This is the all-in number that includes subtotal, taxes, overhead, and profit. Number, no $, no commas.`;

  const payload = {
    contents: [{
      parts: [
        { text: prompt },
        { inline_data: { mime_type: 'application/pdf', data: base64Pdf } },
      ],
    }],
    generationConfig: {
      responseMimeType: 'application/json',
      responseSchema: {
        type: 'OBJECT',
        properties: {
          lineItems: {
            type: 'ARRAY',
            items: {
              type: 'OBJECT',
              properties: {
                description: { type: 'STRING' },
                amount: { type: 'NUMBER' },
              },
              required: ['description', 'amount'],
            },
          },
          grandTotal: { type: 'NUMBER' },
        },
        required: ['lineItems', 'grandTotal'],
      },
    },
  };

  const res = UrlFetchApp.fetch(`${GEMINI_API_URL}?key=${apiKey}`, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    Logger.log(`Gemini error (${fileName}): ${res.getContentText()}`);
    throw new Error(`Gemini API returned ${res.getResponseCode()}. Check the library's execution log.`);
  }

  const body = JSON.parse(res.getContentText());
  const text = body?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('Gemini returned an empty response.');

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    Logger.log(`Could not parse Gemini response: ${text}`);
    throw new Error('Gemini response was not valid JSON. See library execution log.');
  }

  // Safety net: strip any leading 3-letter uppercase code Gemini may have
  // left in despite the prompt. Pattern: 3 uppercase letters + whitespace.
  const lineItems = (parsed.lineItems || []).map((item) => ({
    description: String(item.description || '').replace(/^[A-Z]{3}\s+/, '').trim(),
    amount: Number(item.amount) || 0,
  }));

  return {
    lineItems,
    grandTotal: Number(parsed.grandTotal) || 0,
    fileName,
  };
}

/**
 * Writes the confirmed extraction to the bound budget sheet.
 *
 * Layout (for N line items, starting at row 3):
 *   • Categories          → C3:C{2+N}
 *   • Scope totals        → D3:D{2+N}
 *   • Total budget        → E3:E{2+N}, each row = D{r} * 0.7
 *   • Subtotal row {3+N}  → C = bold "Total" label; D–M = SUM formulas
 *   • B11 (Job Costs)     → =SUM(F{3+N}:M{3+N})
 *   • B5 (Contract Total) → grand total from PDF
 */
function writeBudget(lineItems, grandTotal) {
  if (!Array.isArray(lineItems) || lineItems.length === 0) {
    throw new Error('No line items to write.');
  }

  const sheet = getBudgetSheet_();
  const startRow = BUDGET.CATEGORY_START_ROW;
  const lastDataRow = startRow + lineItems.length - 1;
  const subtotalRow = lastDataRow + 1;

  // Reset content + font weight in the data region, so a re-run leaves
  // no stale rows AND no stale bold formatting from a previous run's
  // subtotal label landing at a different row count.
  const clearRange = sheet.getRange(BUDGET.CLEAR_RANGE);
  clearRange.clearContent();
  clearRange.setFontWeight('normal');

  const categoryValues = lineItems.map((item) => [String(item.description || '').trim()]);
  const amountValues = lineItems.map((item) => [Number(item.amount) || 0]);

  sheet.getRange(`${BUDGET.CATEGORY_COL}${startRow}:${BUDGET.CATEGORY_COL}${lastDataRow}`).setValues(categoryValues);
  sheet.getRange(`${BUDGET.AMOUNT_COL}${startRow}:${BUDGET.AMOUNT_COL}${lastDataRow}`).setValues(amountValues);

  // E column — Total Budget = Scope Total * 0.7, per row
  const totalBudgetFormulas = lineItems.map((_, i) => {
    const r = startRow + i;
    return [`=${BUDGET.AMOUNT_COL}${r}*${BUDGET.TOTAL_BUDGET_RATIO}`];
  });
  sheet
    .getRange(`${BUDGET.TOTAL_BUDGET_COL}${startRow}:${BUDGET.TOTAL_BUDGET_COL}${lastDataRow}`)
    .setFormulas(totalBudgetFormulas);

  // Subtotal row across D–M
  BUDGET.SUBTOTAL_COLS.forEach((col) => {
    const formula = `=SUM(${col}${startRow}:${col}${lastDataRow})`;
    sheet.getRange(`${col}${subtotalRow}`).setFormula(formula);
  });

  // Subtotal row label in C
  sheet.getRange(`${BUDGET.CATEGORY_COL}${subtotalRow}`).setValue(BUDGET.TOTAL_LABEL);

  // Bold the whole subtotal row (label + SUM formulas) so it reads as a summary, not a data row
  const lastSubtotalCol = BUDGET.SUBTOTAL_COLS[BUDGET.SUBTOTAL_COLS.length - 1];
  sheet.getRange(`${BUDGET.CATEGORY_COL}${subtotalRow}:${lastSubtotalCol}${subtotalRow}`).setFontWeight('bold');

  // B11 — Job Costs (subs only)
  const firstSubCol = BUDGET.JOB_COSTS_COLS[0];
  const lastSubCol = BUDGET.JOB_COSTS_COLS[BUDGET.JOB_COSTS_COLS.length - 1];
  sheet.getRange(BUDGET.JOB_COSTS_CELL).setFormula(`=SUM(${firstSubCol}${subtotalRow}:${lastSubCol}${subtotalRow})`);

  // B5 — contract grand total
  sheet.getRange(BUDGET.GRAND_TOTAL_CELL).setValue(Number(grandTotal) || 0);

  // Lay down thin gridline borders across the data region. This is the
  // canonical border state that restoreGridlinesOnEdit maintains afterward.
  applyGridlines_(sheet);

  return {
    rowsWritten: lineItems.length,
    subtotalRow,
    grandTotal: Number(grandTotal) || 0,
  };
}

// ---------- Format guardian — keeps gridlines from disappearing ----------

/**
 * Restores thin gridline borders to the budget data region. Called from
 * the bound script's onEdit trigger so drag-moves and paste actions that
 * strip borders are immediately repaired.
 *
 * Brute-force: reapply the canonical border style across the whole data
 * region on every relevant edit. The region is ~650 cells and completes
 * in well under a second — the user won't see it run.
 */
function restoreGridlinesOnEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== BUDGET.SHEET_NAME) return;

  // Bail if the edit doesn't overlap the managed region. Saves work on
  // header / formula edits and prevents the trigger from churning on
  // unrelated cell changes elsewhere in the sheet.
  // CLEAR_RANGE 'C3:M50' = cols 3-13, rows 3-50
  const r = e.range;
  const overlaps =
    r.getLastColumn() >= 3 && r.getColumn() <= 13 &&
    r.getLastRow() >= 3 && r.getRow() <= 50;
  if (!overlaps) return;

  applyGridlines_(sheet);
}

// ---------- Internal helpers (private — trailing underscore) ----------

function applyGridlines_(sheet) {
  // Black borders, applied explicitly so they survive drag/paste actions
  // that strip default gridline rendering.
  sheet.getRange(BUDGET.CLEAR_RANGE).setBorder(
    true, true, true, true, true, true,
    '#000000',
    SpreadsheetApp.BorderStyle.SOLID
  );
}

function getBudgetSheet_() {
  // SpreadsheetApp.getActive() resolves to the bound caller's spreadsheet,
  // not the library project itself. This is what we want.
  const sheet = SpreadsheetApp.getActive().getSheetByName(BUDGET.SHEET_NAME);
  if (!sheet) {
    throw new Error(`Sheet '${BUDGET.SHEET_NAME}' not found in this workbook.`);
  }
  return sheet;
}