/**
 * ONE-OFF, THROWAWAY — clear the hand-painted stripes
 * ---------------------------------------------------
 * Phase 2 of docs/DEPLOY-2026-10-08.md. Paste at the bottom of the Sheet
 * script's Code.gs, run clearPaintedStripes once, read the Execution log, then
 * DELETE this function and save. It is not part of the script.
 *
 * Native alternating colours render UNDERNEATH a painted cell fill, so the
 * stripes stay invisible until the old painted ones are gone. Run this before
 * applying Format > Alternating colours.
 *
 * Clears the fill from column B across to the last column that has a header,
 * but ONLY where the fill is currently white or #d9d9d9. Every other colour is
 * left alone, so a cell somebody highlighted to flag something survives.
 * Column A is outside the range, so its green/red YES-NO fill is never touched,
 * and row 1 keeps its blue header.
 */
function clearPaintedStripes() {
  var CLEARABLE = ['#ffffff', '#d9d9d9'];      // the two banding colours, nothing else
  var tabs = ['TBR Job Numbers', 'TBR - Completed', 'TBR - JNS'];
  var ss = SpreadsheetApp.getActive();

  for (var t = 0; t < tabs.length; t++) {
    var sh = ss.getSheetByName(tabs[t]);
    if (!sh) { Logger.log('!! TAB NOT FOUND: ' + tabs[t]); continue; }

    // Right edge = last column that has a header, so any unnamed junk columns
    // further right are left alone.
    var lastCol = 0;
    var n = sh.getLastColumn();
    if (n) {
      var hdrs = sh.getRange(1, 1, 1, n).getValues()[0];
      for (var h = 0; h < hdrs.length; h++) if (String(hdrs[h]).trim()) lastCol = h + 1;
    }
    if (lastCol < 2) { Logger.log(tabs[t] + ': no headers past column A, skipped'); continue; }

    var rows = sh.getMaxRows() - 1;             // everything below the header row
    var rng = sh.getRange(2, 2, rows, lastCol - 1);
    var bg = rng.getBackgrounds();

    var cleared = 0, kept = {};
    for (var r = 0; r < bg.length; r++) {
      for (var c = 0; c < bg[r].length; c++) {
        var v = String(bg[r][c]).toLowerCase();
        if (CLEARABLE.indexOf(v) !== -1) {
          bg[r][c] = null;                       // null = no fill, banding shows through
          cleared++;
        } else {
          kept[v] = (kept[v] || 0) + 1;
        }
      }
    }
    rng.setBackgrounds(bg);

    var keptList = [];
    for (var k in kept) keptList.push(k + ' x' + kept[k]);
    Logger.log(tabs[t] + ': cleared ' + cleared + ' cells in B2:' +
               colLtr_(lastCol) + sh.getMaxRows() +
               ' | left alone: ' + (keptList.length ? keptList.join(', ') : 'none'));
  }
  Logger.log('Done. Now apply Format > Alternating colours, then delete this function.');
}

function colLtr_(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}
