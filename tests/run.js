#!/usr/bin/env node
// TBR Apps Script test suite.
//   node tests/run.js
// Runs the real .gs sources against fake Sheets objects built from the
// production layout captured by a diagnostic on 2026-09-21: 23 named columns
// A-W, Project Manager at D, Job Status at I, Contract Value J, Estimated K,
// Actual L, plus Type of Job and Created Date appended at X and Y.
const fs = require('fs'), vm = require('vm'), path = require('path');
const { makeSheet, sandbox } = require('./fakeSheets');

const ROOT = path.join(__dirname, '..');
const MOVER = path.join(ROOT, 'Job Number Creation Sheet Script', 'Code.gs');
const FORM  = path.join(ROOT, 'Job Number Creation Form', 'Code.gs');

// Live layout as of the header dump on 2026-10-08. Since 2026-09-21 the sheet
// gained "Profit $" at M and "Type of Job" at O, and "Job Completion Date"
// moved from N to the far right. Nothing in the scripts keys off position, but
// these fixtures must track reality or the tests stop meaning anything.
const BASE = ['Create Job #?','Project Number','Project Name','Project Manager','Project Folder',
  'Project Address','Contact Number','Email','Job Status','Contract Value','Estimated Cost',
  'Actual Cost','Profit $','Job Profit %','Type of Job','Insurance','Claim Number','Date of Loss',
  'Type of Loss','Referral Name','Referral Number','Job # Status','Project Status',
  'Estimated Completion','Job Completion Date'];
const W = BASE.length;        // 25, A through Y
const COL = n => BASE.indexOf(n);

let pass = 0, fail = 0;
const ck = (n, c, d) => {
  if (c) { console.log('  PASS  ' + n); pass++; }
  else { console.log('  FAIL  ' + n + (d ? '\n        ' + d : '')); fail++; }
};
const suite = n => console.log('\n--- ' + n + ' ---');

const load = (file, sb, tail) => {
  vm.createContext(sb);
  vm.runInContext(fs.readFileSync(file, 'utf8') + '\n;' + tail, sb);
  return sb;
};

/* ================= Job-Mover ================= */

function jobRow(tag, c, e, a) {
  const r = new Array(W).fill('');
  r[COL('Create Job #?')]='YES';
  r[COL('Project Name')]='Job '+tag;
  r[COL('Project Manager')]='PM '+tag;
  r[COL('Job Status')]='JNS (Job Not Sold)';
  r[COL('Contract Value')]=c; r[COL('Estimated Cost')]=e; r[COL('Actual Cost')]=a;
  return r;
}
const totalsRow = () => { const r = new Array(W).fill(''); r[0]='TOTAL'; return r; };

function moverWorld(destRows, opts) {
  opts = opts || {};
  const src = makeSheet('TBR Job Numbers', [BASE, jobRow('X',100,60,55)],
                        { formatTag:'SRC', banding: opts.srcBand === undefined ? 'B2:Y2' : opts.srcBand });
  const dst = makeSheet('TBR - JNS', destRows,
                        { formatTag:'DEST', banding: opts.destBand === undefined ? 'B2:Y2' : opts.destBand });
  const ss = { getSheetByName: n => ({'TBR Job Numbers':src,'TBR - JNS':dst}[n] || null) };
  src._ss = ss; dst._ss = ss;
  return { src, dst, ss };
}

function move(w) {
  const sb = sandbox(w.ss); sb.__src = w.src;
  load(MOVER, sb,
    'try{ globalThis.__r={ok:performMove_(globalThis.__src,2,"JNS (Job Not Sold)")}; }catch(e){ globalThis.__r={err:e.message}; }' +
    ';globalThis.__api={findTotalsRow_,refreshBandingExtent_,lastHeaderColumn_};');
  return { r: sb.__r, logs: sb.__logs, api: sb.__api };
}

suite('Mover: tab WITH a totals row');
let w = moverWorld([BASE, jobRow('A',1000,600,550), jobRow('B',2000,1200,1100), totalsRow()]);
let o = move(w);
ck('no error', !o.r.err, o.r.err);
ck('job inserted at row 4, above totals', w.dst._v[3] && w.dst._v[3][2]==='Job X');
ck('TOTAL pushed to row 5', w.dst._v[4] && w.dst._v[4][0]==='TOTAL');
ck('existing jobs undisturbed', w.dst._v[1][2]==='Job A' && w.dst._v[2][2]==='Job B');
ck('Contract Value sum spans J2:J4', w.dst._fx[4][COL('Contract Value')]==='=SUM(J2:J4)',
   w.dst._fx[4][COL('Contract Value')]);
ck('Estimated Cost sum spans K2:K4', w.dst._fx[4][COL('Estimated Cost')]==='=SUM(K2:K4)',
   w.dst._fx[4][COL('Estimated Cost')]);
ck('Actual Cost sum spans L2:L4', w.dst._fx[4][COL('Actual Cost')]==='=SUM(L2:L4)',
   w.dst._fx[4][COL('Actual Cost')]);
ck('Job Profit % NOT summed, averaging percentages is meaningless',
   w.dst._fx[4][COL('Job Profit %')]==='');
ck('Profit $ NOT summed either, pending Joe\'s call on adding it',
   w.dst._fx[4][COL('Profit $')]==='');
ck('source row deleted', w.src._v.length===1);
ck('formatting came from the archive tab, not the source row',
   w.dst._fmt[3].every(c=>c==='DEST'), JSON.stringify(w.dst._fmt[3].slice(0,4)));
ck('copied from the row directly above', w.dst._fmtFrom.join()==='TBR - JNS row 3', w.dst._fmtFrom.join());

suite('Mover: tab WITHOUT a totals row');
w = moverWorld([BASE, jobRow('A',1000,600,550)]); o = move(w);
ck('no error', !o.r.err, o.r.err);
ck('job appended at row 3', w.dst._v[2] && w.dst._v[2][2]==='Job X');
ck('appended row also inherits from the archive tab', w.dst._fmt[2].every(c=>c==='DEST'));
ck('no totals reported', o.r.ok && o.r.ok.totalsRow===0);
ck('no stray formulas written', w.dst._fx.every(r=>r.every(c=>c==='')));

suite('Mover: totals row but no job rows yet');
w = moverWorld([BASE, totalsRow()]); o = move(w);
ck('no error', !o.r.err, o.r.err);
ck('job landed at row 2', w.dst._v[1][2]==='Job X');
ck('empty archive falls back to the source row, never the header',
   w.dst._fmt[1].every(c=>c==='SRC'), JSON.stringify(w.dst._fmt[1].slice(0,4)));
ck('sum spans J2:J2', w.dst._fx[2][COL('Contract Value')]==='=SUM(J2:J2)',
   w.dst._fx[2][COL('Contract Value')]);

suite('Mover: banding follows the data extent');
w = moverWorld([BASE, jobRow('A',1,1,1), totalsRow()], { destBand:'B2:Y3' }); o = move(w);
ck('destination banding resized to B2:Y4', w.dst._bands[0].getRange().getA1Notation()==='B2:Y4',
   w.dst._bands[0].getRange().getA1Notation());
ck('source banding untouched (no data rows left)', w.src._bandSet===0);

suite('Mover: tab with no banding at all');
w = moverWorld([BASE, totalsRow()], { destBand:null }); o = move(w);
ck('no error', !o.r.err, o.r.err);
ck('logged the missing banding', o.logs.some(m=>/expected exactly 1/.test(m)));
ck('job still moved', w.dst._v[1][2]==='Job X');

suite('Mover: helpers');
w = moverWorld([BASE, jobRow('A',1,1,1), totalsRow()]); o = move(w);
const probe = r => makeSheet('t', r, { banding:'B2:Y3' });
ck('finds the TOTAL row', o.api.findTotalsRow_(probe([BASE, jobRow('A',1,1,1), totalsRow()]))===3);
const lower = totalsRow(); lower[0]='total';
ck('label match ignores case', o.api.findTotalsRow_(probe([BASE, jobRow('A',1,1,1), lower]))===3);
ck('returns 0 when absent', o.api.findTotalsRow_(probe([BASE, jobRow('A',1,1,1)]))===0);
ck('last header column is Y', o.api.lastHeaderColumn_(probe([BASE, jobRow('A',1,1,1)]))===25);

suite('Mover: header guard still refuses a mismatched tab');
w = moverWorld([BASE.filter(h=>h!=='Project Manager'), totalsRow()]); o = move(w);
ck('threw', !!o.r.err);
ck('names column D', !!o.r.err && o.r.err.indexOf('Column D')!==-1);
ck('source row NOT deleted', w.src._v.length===2);
ck('nothing written to the archive', w.dst._v.length===2);

suite('Totals: rebuilt after a manual row delete');
// Completed-style tab: three jobs then TOTAL, with sums already spanning them.
w = moverWorld([BASE, jobRow('A',1000,600,550), jobRow('B',2000,1200,1100),
                jobRow('C',3000,1800,1700), totalsRow()]);
let sb2 = sandbox(w.ss); sb2.__ss = w.ss; sb2.__dst = w.dst;
load(MOVER, sb2,
  'globalThis.__api2 = { refreshAllTotals_, onChangeInstallable, refreshTotalsNow, findTotalsRow_ };');
// Someone deletes job B by hand. Sheets would shrink the SUM itself, but a
// totals cell that was typed rather than written by the script would not.
w.dst.deleteRow(3);
w.dst._fx[3][COL('Contract Value')] = '';              // simulate a non-formula totals cell
w.dst._v[3][COL('Contract Value')] = '9999';           // a stale typed number
sb2.__api2.onChangeInstallable({ changeType: 'REMOVE_ROW' });
ck('totals row found after the delete', sb2.__api2.findTotalsRow_(w.dst) === 4);
ck('Contract Value sum rewritten to span J2:J3',
   w.dst._fx[3][COL('Contract Value')] === '=SUM(J2:J3)', w.dst._fx[3][COL('Contract Value')]);
ck('the stale typed number was replaced by a formula',
   w.dst._v[3][COL('Contract Value')] === '<calc>', String(w.dst._v[3][COL('Contract Value')]));
ck('Estimated Cost rebuilt too', w.dst._fx[3][COL('Estimated Cost')] === '=SUM(K2:K3)');
ck('Actual Cost rebuilt too', w.dst._fx[3][COL('Actual Cost')] === '=SUM(L2:L3)');

suite('Totals: the manual rerun function');
w = moverWorld([BASE, jobRow('A',10,5,4), totalsRow()]);
let sb3 = sandbox(w.ss);
load(MOVER, sb3, 'globalThis.__api3 = { refreshTotalsNow };');
sb3.__api3.refreshTotalsNow();
ck('rebuilt the sums', w.dst._fx[2][COL('Contract Value')] === '=SUM(J2:J2)',
   w.dst._fx[2][COL('Contract Value')]);
ck('said which tab it touched', sb3.__logs.some(m => /Totals rebuilt on: TBR - JNS/.test(m)),
   sb3.__logs.join(' | '));

suite('Totals: rerun on a tab with no totals row says so and does not throw');
w = moverWorld([BASE, jobRow('A',10,5,4)]);
let sb4 = sandbox(w.ss);
load(MOVER, sb4, 'globalThis.__api4 = { refreshTotalsNow };');
let threw = false;
try { sb4.__api4.refreshTotalsNow(); } catch (e) { threw = true; }
ck('did not throw', !threw);
ck('explained that no totals row was found',
   sb4.__logs.some(m => /No totals row found/.test(m)), sb4.__logs.join(' | '));
ck('wrote no formulas', w.dst._fx.every(r => r.every(c => c === '')));

suite('Totals: a non-row-count change is ignored');
w = moverWorld([BASE, jobRow('A',10,5,4), totalsRow()]);
let sb5 = sandbox(w.ss);
load(MOVER, sb5, 'globalThis.__api5 = { onChangeInstallable };');
sb5.__api5.onChangeInstallable({ changeType: 'FORMAT' });
ck('FORMAT change does not rewrite totals', w.dst._fx[2][COL('Contract Value')] === '');

/* ================= Lead form: what an appended row inherits ================= */

const RULES = { [COL('Create Job #?')]:'YESNO', [COL('Project Manager')]:'PM_LIST',
                [COL('Job Status')]:'STATUS_LIST' };
function formSheet(nRows, opts) {
  opts = opts || {};
  const rows = [BASE.slice()], dv = [new Array(W).fill(null)];
  for (let i=0;i<nRows;i++) {
    const r = new Array(W).fill(''); r[COL('Create Job #?')]='YES';
    r[COL('Project Name')]='Job '+(i+1); rows.push(r);
    const line = new Array(W).fill(null);
    if (!opts.noRules) for (const k in RULES) line[k]=RULES[k];
    dv.push(line);
  }
  const sh = makeSheet('TBR Job Numbers', rows,
    { validations: dv, banding: opts.band === undefined ? 'B2:Y'+(nRows+1) : opts.band });
  return sh;
}
function append(sheet, vals) {
  const sb = sandbox(sheet); sb.__s = sheet; sb.__v = vals;
  load(FORM, sb, 'try{ appendLeadRow(globalThis.__s, globalThis.__v); globalThis.__e=null; }catch(e){ globalThis.__e=e.message; }');
  return { err: sb.__e, logs: sb.__logs };
}

suite('Form append: inherits dropdowns, formatting and the stripe');
let sh = formSheet(3);
let a = append(sh, { 'Project Name':'New Lead' });
ck('no error', !a.err, a.err);
ck('row appended at row 5', sh.getLastRow()===5);
ck('column A got the YES/NO list', sh._dv[4][COL('Create Job #?')]==='YESNO');
ck('column D got the Project Manager list', sh._dv[4][COL('Project Manager')]==='PM_LIST');
ck('column I got the Job Status list', sh._dv[4][COL('Job Status')]==='STATUS_LIST');
ck('a column with no rule stays null', sh._dv[4][COL('Project Name')]===null);
ck('validation written in one call', sh._dvSetCalls===1, String(sh._dvSetCalls));
ck('inherited borders/fonts/number formats', sh._fmt[4].every(c=>c==='FMT'));
ck('format copied in one call', sh._fmtCopies===1, String(sh._fmtCopies));
ck('format copy did not clobber the values', sh._v[4][COL('Project Name')]==='New Lead');
ck('banding stretched to B2:Y5', sh._bands[0].getRange().getA1Notation()==='B2:Y5',
   sh._bands[0].getRange().getA1Notation());

suite('Form append: very first data row inherits nothing from the header');
sh = formSheet(0, { band:'B2:Y2' }); a = append(sh, { 'Project Name':'First Ever' });
ck('no error', !a.err, a.err);
ck('row landed at row 2', sh.getLastRow()===2);
ck('no validation taken from the header', sh._dv[1].every(c=>c===null));
ck('no formatting taken from the header', sh._fmt[1].every(c=>c==='BARE'));

suite('Form append: row above has no dropdowns');
sh = formSheet(2, { noRules:true }); a = append(sh, { 'Project Name':'No Rules' });
ck('no error', !a.err, a.err);
ck('no setDataValidations call made', !sh._dvSetCalls);

suite('Form append: cosmetics failing must never cost a lead');
sh = formSheet(2);
const realGetRange = sh.getRange.bind(sh);
sh.getRange = (r,c,nr,nc) => {
  const rng = realGetRange(r,c,nr,nc);
  rng.getDataValidations = () => { throw new Error('validation service down'); };
  rng.copyTo = () => { throw new Error('format service down'); };
  return rng;
};
sh.getBandings = () => { throw new Error('banding service down'); };
a = append(sh, { 'Project Name':'Still Must Land' });
ck('appendLeadRow did not throw', !a.err, a.err);
ck('the lead row still landed', sh.getLastRow()===4, String(sh.getLastRow()));
ck('all three failures logged, not swallowed', a.logs.length===3, JSON.stringify(a.logs));

/* ================= Lead form: every collected field reaches a column ================= */

// Type of Job already exists at O. Created Date is the one header still to add.
const HEADERS = BASE.concat(['Created Date']);
const ANSWERS = {
  'Project Name':'Smith, John', 'Project Manager':'Dana Reyes', 'Type of Job':'Reconstruction',
  'Project Category':'MITIGATION', 'Create Project Folder?':'No', 'Create Project Number?':'No',
  'Project Address':'1 Test Way', 'Contact Number':'555-0100', 'Email':'nobody@example.com',
  'Contract Value':'1000', 'Estimated Cost':'600', 'Job Completion Date':'2026-10-01',
  'Insurance Company':'Other', 'Other Insurance Company':'Acme Mutual', 'Claim Number':'CLM-12345',
  'Date of Loss':'2026-09-01', 'Type of Loss':'Water', 'Referral Name':'Jane Ref', 'Referral Number':'555-0199'
};
// Collected for folder routing, deliberately not stored. Anything else showing
// up here is a regression: a field taken from a customer and thrown away.
const KNOWN_HOMELESS = ['Project Category'];

suite('Form submit: no field is silently dropped');
const target = makeSheet('TBR Job Numbers', [HEADERS.slice()], { banding:null });
const items = Object.keys(ANSWERS).map(k => ({ getItem:()=>({getTitle:()=>k}), getResponse:()=>ANSWERS[k] }));
const sb = sandbox(target, {
  DriveApp: { getFolderById(){ throw new Error('DriveApp must not be reached'); } },
  Utilities: { formatDate(d,tz,fmt){ const p=n=>String(n).padStart(2,'0');
    return fmt==='MM/dd/yyyy' ? p(d.getUTCMonth()+1)+'/'+p(d.getUTCDate())+'/'+d.getUTCFullYear() : d.toISOString(); } },
  Session: { getScriptTimeZone:()=>'UTC' }
});
sb.__ev = { response:{ getItemResponses:()=>items, getTimestamp:()=>new Date('2026-09-21T12:00:00Z') } };
load(FORM, sb, 'onLeadFormSubmit(globalThis.__ev);');
const at = n => (target._v[1] || [])[HEADERS.indexOf(n)];
const dropped = sb.__logs.filter(m=>/sheet is missing column/.test(m)).map(m=>m.replace(/.*\[(.*)\].*/,'$1'));
ck('no NEW silently dropped field', dropped.filter(d=>KNOWN_HOMELESS.indexOf(d)===-1).length===0,
   dropped.join(', '));
ck('the deliberately-unstored list is unchanged',
   dropped.length===KNOWN_HOMELESS.length && KNOWN_HOMELESS.every(k=>dropped.indexOf(k)!==-1),
   'saw: ' + dropped.join(', '));

suite('Form submit: values land in the right columns');
ck('Project Name',    at('Project Name')==='Smith, John', String(at('Project Name')));
ck('Project Manager', at('Project Manager')==='Dana Reyes', String(at('Project Manager')));
ck('Type of Job',     at('Type of Job')==='Reconstruction', String(at('Type of Job')));
ck('Created Date',    at('Created Date')==='09/21/2026', String(at('Created Date')));
ck('Claim Number',    at('Claim Number')==='CLM-12345', String(at('Claim Number')));
ck('Type of Loss',    at('Type of Loss')==='Water');
ck('Contract Value',  at('Contract Value')==='1000');
ck('Insurance resolves "Other" to the free-text answer', at('Insurance')==='Acme Mutual', String(at('Insurance')));
ck('Date of Loss reformatted', at('Date of Loss')==='09/01/2026', String(at('Date of Loss')));

/* ================= config sanity, catches deploy landmines ================= */

suite('Config: both scripts point at the same production spreadsheet');
const idOf = f => (fs.readFileSync(f,'utf8').match(/^var TRACKING_SHEET_ID = '([^']+)'/m) || [])[1];
const PROD = '1D6kppfGobZ42vRSN6xVTw7mS8YdmthXh1ICuudCmeFE';
ck('Sheet Script points at production', idOf(MOVER)===PROD, String(idOf(MOVER)));
ck('Form script points at the same spreadsheet', idOf(FORM)===PROD, String(idOf(FORM)));
ck('no test-copy id committed anywhere',
   idOf(MOVER)!=='1OD-IkT6S43YqMwqdPhNNC4kNsPTqcmK7JXKn3hF_aaw' &&
   idOf(FORM)!=='1OD-IkT6S43YqMwqdPhNNC4kNsPTqcmK7JXKn3hF_aaw');

console.log('\n' + (fail===0 ? 'ALL ' + pass + ' CHECKS PASSED' : pass + ' passed, ' + fail + ' FAILED'));
process.exit(fail===0 ? 0 : 1);
