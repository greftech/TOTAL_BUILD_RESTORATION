// Minimal fake of the Apps Script Spreadsheet service: enough to drive the real
// .gs files in Node. Each sheet carries parallel grids so a test can assert on
// values, formulas, number formats, styling and data validation independently.
//
// _fmt is a stand-in for "everything PASTE_FORMAT carries" (borders, fonts,
// alignment). Tagging it per sheet is how the mover tests prove a row took its
// formatting from the archive tab rather than from the source row.

function columnLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
  return s;
}

function makeSheet(name, rows, opts) {
  opts = opts || {};
  const tag = opts.formatTag || 'FMT';
  const width = () => (rows[0] ? rows[0].length : 0);

  const sh = {
    _name: name,
    _v:   rows.map(r => r.slice()),
    _nf:  rows.map(r => r.map(() => 'General')),
    _fx:  rows.map(r => r.map(() => '')),
    _dv:  opts.validations || rows.map(r => r.map(() => null)),
    _fmt: rows.map((r, ri) => r.map(() => (ri === 0 ? 'HEADER' : tag))),
    _bands: [], _bandSet: 0, _fmtCopies: 0, _fmtFrom: [], _dvSetCalls: 0,

    getName() { return this._name; },
    getParent() { return this._ss; },
    getMaxColumns() { return this._v[0] ? this._v[0].length : 0; },
    getMaxRows() { return this._v.length; },
    getBandings() { return this._bands; },
    getLastColumn() {
      let l = 0;
      this._v.forEach(r => r.forEach((c, i) => { if (c !== '' && c != null) l = Math.max(l, i + 1); }));
      return l;
    },
    getLastRow() {
      let l = 0;
      this._v.forEach((r, ri) => { if (r.some(c => c !== '' && c != null)) l = ri + 1; });
      return l;
    },
    _blank() { return new Array(this.getMaxColumns()).fill(''); },
    appendRow(row) {
      const w = this.getMaxColumns(), v = new Array(w).fill('');
      for (let i = 0; i < row.length && i < w; i++) v[i] = row[i];
      this._v.push(v);
      this._nf.push(new Array(w).fill('General'));
      this._fx.push(new Array(w).fill(''));
      this._dv.push(new Array(w).fill(null));
      this._fmt.push(new Array(w).fill('BARE'));
    },
    insertRowsAfter(after, n) {
      const w = this.getMaxColumns();
      for (let i = 0; i < n; i++) {
        this._v.push(new Array(w).fill('')); this._nf.push(new Array(w).fill('General'));
        this._fx.push(new Array(w).fill('')); this._dv.push(new Array(w).fill(null));
        this._fmt.push(new Array(w).fill('BARE'));
      }
    },
    insertRowBefore(row) {
      const w = this.getMaxColumns();
      this._v.splice(row - 1, 0, new Array(w).fill(''));
      this._nf.splice(row - 1, 0, new Array(w).fill('General'));
      this._fx.splice(row - 1, 0, new Array(w).fill(''));
      this._dv.splice(row - 1, 0, new Array(w).fill(null));
      this._fmt.splice(row - 1, 0, new Array(w).fill('BARE'));
    },
    deleteRow(row) {
      [this._v, this._nf, this._fx, this._dv, this._fmt].forEach(g => g.splice(row - 1, 1));
    },

    getRange(row, col, nr, nc) {
      const s = this; nr = nr === undefined ? 1 : nr; nc = nc === undefined ? 1 : nc;
      const grid = (store, dflt) => {
        const out = [];
        for (let r = 0; r < nr; r++) {
          const line = [];
          for (let c = 0; c < nc; c++) {
            const rr = store[row - 1 + r];
            line.push(rr && rr[col - 1 + c] !== undefined ? rr[col - 1 + c] : dflt);
          }
          out.push(line);
        }
        return out;
      };
      const put = (store, vals, dflt) => {
        for (let r = 0; r < nr; r++) {
          if (!store[row - 1 + r]) store[row - 1 + r] = [];
          for (let c = 0; c < nc; c++) {
            store[row - 1 + r][col - 1 + c] = vals[r][c] === undefined ? dflt : vals[r][c];
          }
        }
      };
      return {
        _sh: s, _row: row, _col: col, _nr: nr, _nc: nc,
        getSheet: () => s, getRow: () => row, getColumn: () => col, getNumColumns: () => nc,
        getA1Notation: () => columnLetter(col) + row + ':' + columnLetter(col + nc - 1) + (row + nr - 1),
        getValues: () => grid(s._v, ''),
        getValue:  () => grid(s._v, '')[0][0],
        getFormula: () => grid(s._fx, '')[0][0],
        getNumberFormats: () => grid(s._nf, 'General'),
        getDataValidations: () => grid(s._dv, null),
        getFormats: () => grid(s._fmt, 'BARE'),
        setValues: v => put(s._v, v, ''),
        setValue(v) { put(s._v, [[v]], ''); },
        setFormula(f) { put(s._fx, [[f]], ''); put(s._v, [['<calc>']], ''); },
        setNumberFormats: f => put(s._nf, f, 'General'),
        setDataValidations(rules) { put(s._dv, rules, null); s._dvSetCalls++; },
        copyTo(target, type) {
          if (type !== 'PASTE_FORMAT') throw new Error('unexpected CopyPasteType: ' + type);
          const nf = grid(s._nf, 'General'), st = grid(s._fmt, 'BARE');
          for (let r = 0; r < nr; r++) {
            const tr = target._row - 1 + r;
            if (!target._sh._nf[tr]) target._sh._nf[tr] = [];
            if (!target._sh._fmt[tr]) target._sh._fmt[tr] = [];
            for (let c = 0; c < nc; c++) {
              target._sh._nf[tr][target._col - 1 + c] = nf[r][c];
              target._sh._fmt[tr][target._col - 1 + c] = st[r][c];
            }
          }
          target._sh._fmtCopies++;
          target._sh._fmtFrom.push(s._name + ' row ' + row);
        }
      };
    }
  };

  if (opts.banding) {
    sh._bands.push({
      _r: { getA1Notation: () => opts.banding },
      getRange() { return this._r; },
      setRange(x) { this._r = x; sh._bandSet++; }
    });
  }
  return sh;
}

/** Apps Script globals the two .gs files touch. */
function sandbox(active, extra) {
  const logs = [];
  const sb = Object.assign({
    SpreadsheetApp: {
      flush() {},
      getActive: () => active,
      // Callers pass either a spreadsheet-like object or a bare sheet; openById
      // must always hand back something with getSheetByName on it.
      openById: () => (typeof active.getSheetByName === 'function'
        ? active
        : { getSheetByName: () => active, getId: () => 'fake-id' }),
      CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT' }
    },
    Logger: { log: m => logs.push(String(m)) },
    console
  }, extra || {});
  sb.__logs = logs;
  return sb;
}

module.exports = { columnLetter, makeSheet, sandbox };
