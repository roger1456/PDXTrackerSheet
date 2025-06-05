/* ==========================================================
   CONFIG
   ========================================================== */
const CFG = {
  importSheet:  'FDX Import',
  trackerSheet: 'RDX To Do',
  importStart:  2,   // first data row in FDX Import (row 1 is header)
  trackerStart: 4,   // first data row in tracker (rows 1–3 are headers)
  importCols:   { seq: 0, char: 1, lines: 2, dbStat: 3 }, // A–D
  trCols:       {               // zero-based indexes in tracker
    date: 1, seq: 2, char: 5, lines: 6, dbStat: 7,
    status: 8, notes: 9
  },
  formulaCols: [3, 8, 10, 11, 12, 13], // D, I, K, L, M, N
  green: '#b7e1cd'
};

/* ==========================================================
   1) Stamp Date on manual edits (rows 4+, cols A–J)
   ========================================================== */
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== CFG.trackerSheet) return;
  if (e.range.getRow() < CFG.trackerStart) return;
  if (e.range.getColumn() > 10) return;
  sh.getRange(e.range.getRow(), CFG.trCols.date + 1).setValue(new Date());
}

/* ==========================================================
   2) Add menu
   ========================================================== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Update Tracker')
    .addItem('Import Line Statuses', 'importLinesData')
    .addToUi();
}

/* ==========================================================
   3) Main importer
   ========================================================== */
function importLinesData() {
  const ss   = SpreadsheetApp.getActive();
  const iSh  = ss.getSheetByName(CFG.importSheet);
  const tSh  = ss.getSheetByName(CFG.trackerSheet);
  const now  = new Date();

  /* ---------- STEP 1: read import ---------- */
  const iLast = iSh.getLastRow();
  if (iLast < CFG.importStart) return uiAlert('FDX Import is empty.');

  const iVals = iSh.getRange(CFG.importStart, 1,
                             iLast - CFG.importStart + 1, 4).getValues();
  // keep only rows with Sequence & Character
  const importRows = [];
  const seqSet     = new Set();       // sequences present in this import
  const keySet     = new Set();       // full keys we’ll process

  const makeKey = (s, c, st) =>
      `${s.toString().trim()}|${c.toString().trim()}|${st.toString().trim()}`;
  const makeSC  = (s, c) =>
      `${s.toString().trim()}|${c.toString().trim()}`;

  iVals.forEach((r, idx) => {
    const seq = r[CFG.importCols.seq];
    const chr = r[CFG.importCols.char];
    if (seq === '' || chr === '') return;
    importRows.push({ r, sheetRow: idx + CFG.importStart });
    seqSet.add(seq);
    keySet.add(makeKey(seq, chr, r[CFG.importCols.dbStat]));
  });
  if (importRows.length === 0) return uiAlert('No usable rows in import.');

  /* ---------- STEP 2: snapshot tracker ---------- */
  const tLast = tSh.getLastRow();
  const tVals = tSh.getRange(CFG.trackerStart, 1,
                             tLast - CFG.trackerStart + 1,
                             tSh.getLastColumn()).getValues();

  const keyToRow = Object.create(null);   // full key → row#
  const scToRow  = Object.create(null);   // seq|char → first row#
  const templates= [];                    // template row numbers

  tVals.forEach((row, idx) => {
    const sheetRow = idx + CFG.trackerStart;
    const seq = row[CFG.trCols.seq];
    const chr = row[CFG.trCols.char];
    const st  = row[CFG.trCols.dbStat];
    if (seq === '' && chr === '' && st === '') {
      templates.push(sheetRow);
      return;
    }
    const key = makeKey(seq, chr, st);
    keyToRow[key] = sheetRow;
    const sc = makeSC(seq, chr);
    if (!(sc in scToRow)) scToRow[sc] = sheetRow;
  });

  /* ---------- STEP 3: updates + adds ---------- */
  let upd = 0, add = 0;
  importRows.forEach(({ r: row, sheetRow: iRow }) => {
    const seq = row[CFG.importCols.seq];
    const chr = row[CFG.importCols.char];
    const lin = row[CFG.importCols.lines];
    const st  = row[CFG.importCols.dbStat];
    const key = makeKey(seq, chr, st);
    const sc  = makeSC(seq, chr);

    let tgt;
    if (key in keyToRow) {
      // --- update
      tgt = keyToRow[key];
      tSh.getRange(tgt, CFG.trCols.date + 1).setValue(now);
      tSh.getRange(tgt, CFG.trCols.seq  + 1).setValue(seq);
      tSh.getRange(tgt, CFG.trCols.char + 1).setValue(chr);
      tSh.getRange(tgt, CFG.trCols.lines+ 1).setValue(lin);
      tSh.getRange(tgt, CFG.trCols.dbStat+1).setValue(st);
      upd++;
    } else {
      // --- add
      if (templates.length) {
        tgt = templates.shift();
        tSh.getRange(tgt, CFG.trCols.date + 1).setValue(now);
        tSh.getRange(tgt, CFG.trCols.seq  + 1).setValue(seq);
        tSh.getRange(tgt, CFG.trCols.char + 1).setValue(chr);
        tSh.getRange(tgt, CFG.trCols.lines+ 1).setValue(lin);
        tSh.getRange(tgt, CFG.trCols.dbStat+1).setValue(st);
      } else {
        // real append + copy formulas once
        const lastCol = tSh.getLastColumn();
        const blank   = Array(lastCol).fill('');
        blank[CFG.trCols.date ] = now;
        blank[CFG.trCols.seq  ] = seq;
        blank[CFG.trCols.char ] = chr;
        blank[CFG.trCols.lines] = lin;
        blank[CFG.trCols.dbStat] = st;
        tSh.appendRow(blank);
        tgt = tSh.getLastRow();
        const srcRow = tgt - 1;
        CFG.formulaCols.forEach(col =>
          tSh.getRange(tgt, col + 1)
              .setFormula(tSh.getRange(srcRow, col + 1).getFormula()));
      }
      add++;
    }
    keyToRow[key] = tgt;
    if (!(sc in scToRow)) scToRow[sc] = tgt;

    // mark & clear import row
    iSh.getRange(iRow, 1, 1, 4)
       .setBackground(CFG.green)
       .clearContent();
  });

  /* ---------- STEP 4: delete stale rows (only sequences in import) ---------- */
  const toDel = [];
  for (let seq of seqSet) {
    // quick scan rows where column C == seq
    const colC = tSh.getRange(CFG.trackerStart, CFG.trCols.seq + 1,
                              tSh.getLastRow() - CFG.trackerStart + 1, 1)
                    .getValues();
    colC.forEach((v, offset) => {
      if (v[0] !== seq) return;
      const sheetRow = offset + CFG.trackerStart;
      const chr = tSh.getRange(sheetRow, CFG.trCols.char + 1).getValue();
      const st  = tSh.getRange(sheetRow, CFG.trCols.dbStat + 1).getValue();
      if (chr === '' && st === '') return; // template
      const key = makeKey(seq, chr, st);
      if (keySet.has(key)) return;         // still active

      // migrate note (if any)
      const noteCell = tSh.getRange(sheetRow, CFG.trCols.notes + 1);
      const oldNote  = noteCell.getNote();
      if (oldNote) {
        const destRow = scToRow[makeSC(seq, chr)];
        const destCell = tSh.getRange(destRow, CFG.trCols.notes + 1);
        const merged   =
          (destCell.getNote() ? destCell.getNote() + '\n' : '') +
          `Old Status Note: ${st} – ${oldNote}`;
        destCell.setNote(merged);
      }
      toDel.push(sheetRow);
    });
  }
  // delete bottom-up
  toDel.sort((a, b) => b - a).forEach(r => tSh.deleteRow(r));

  /* ---------- STEP 5: sort ---------- */
  const lastRow = tSh.getLastRow();
  if (lastRow >= CFG.trackerStart) {
    tSh.getRange(CFG.trackerStart, 1,
                 lastRow - CFG.trackerStart + 1,
                 tSh.getLastColumn())
       .sort([
         { column: CFG.trCols.seq  + 1, ascending: true },
         { column: CFG.trCols.char + 1, ascending: true }
       ]);
  }

  uiAlert(`Import finished:\n• ${upd} updated\n• ${add} added\n• ${toDel.length} deleted`);
}

/* ---------- small helper ---------- */
function uiAlert(msg) { SpreadsheetApp.getUi().alert(msg); }
