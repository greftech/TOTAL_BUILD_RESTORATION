/**
 * HEADER DUMP — run this whenever the sheet layout changes
 * --------------------------------------------------------
 * Paste at the bottom of the Sheet script's Code.gs, pick dumpHeaders, Run,
 * then send me the Execution log. Delete it when you're done.
 *
 * PRIVACY: row 1 only. No job rows, names or customer data are read.
 *
 * Reports each tab's headers with their column letters, flags any position
 * where the three tabs disagree (which is what makes the Job-Mover refuse a
 * move), and names the first free column for adding a new one.
 */
function dumpHeaders() {
  var ss = SpreadsheetApp.getActive();
  var tabs = ['TBR Job Numbers', 'TBR - Completed', 'TBR - JNS'];
  var grid = {};

  for (var t = 0; t < tabs.length; t++) {
    var sh = ss.getSheetByName(tabs[t]);
    if (!sh) { Logger.log('!! TAB NOT FOUND: ' + tabs[t]); continue; }

    var n = sh.getLastColumn();
    var hdrs = n ? sh.getRange(1, 1, 1, n).getValues()[0] : [];
    var named = 0;
    for (var i = 0; i < hdrs.length; i++) if (String(hdrs[i]).trim()) named = i + 1;

    grid[tabs[t]] = hdrs.map(function (h) { return String(h).trim(); });

    Logger.log('===== ' + tabs[t] + ' =====');
    Logger.log('lastCol=' + n + ' (' + cl_(n) + ')  lastNamedCol=' + named + ' (' + cl_(named) +
               ')  lastRow=' + sh.getLastRow() + '  maxRows=' + sh.getMaxRows());
    for (var j = 0; j < hdrs.length; j++) {
      Logger.log('   ' + cl_(j + 1) + ': ' + JSON.stringify(String(hdrs[j])));
    }
    Logger.log('   first free column for a NEW header: ' + cl_(named + 1));
    Logger.log('');
  }

  // Where do the three tabs disagree? This is exactly what blocks a job move.
  var src = grid['TBR Job Numbers'] || [];
  var width = 0;
  for (var s = 0; s < src.length; s++) if (src[s]) width = s + 1;

  ['TBR - Completed', 'TBR - JNS'].forEach(function (name) {
    var dst = grid[name] || [];
    var bad = [];
    for (var i = 0; i < width; i++) {
      var want = src[i] || '';
      var got = dst[i] === undefined ? '' : dst[i];
      if (want !== got) {
        bad.push(cl_(i + 1) + ': this tab ' + (got ? '"' + got + '"' : '(blank)') +
                 ', main tab "' + want + '"');
      }
    }
    Logger.log('MOVES TO ' + name + ': ' +
               (bad.length ? 'BLOCKED, ' + bad.length + ' column(s) disagree' : 'OK, headers match'));
    bad.slice(0, 8).forEach(function (b) { Logger.log('   ' + b); });
    if (bad.length > 8) Logger.log('   ...and ' + (bad.length - 8) + ' more');
    Logger.log('');
  });
}

function cl_(n) {
  var s = '';
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}
