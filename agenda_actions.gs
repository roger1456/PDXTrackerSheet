/* ===== New & Add buttons for Agenda Builder ===== */

const TAB_DEST        = 'Agenda Builder';
const DEST_FIRST_ROW  = 5;
const SCHED_START_ROW = 5;

/* ---- New Agenda ---- */
function newAgenda() {
  const rows = getQuickRows_();                       // util from quick_view_utils.gs
  const dest = SpreadsheetApp.getActive().getSheetByName(TAB_DEST);

  /* clear old B-E and G-J content (keeps formatting) */
  if (dest.getLastRow() >= DEST_FIRST_ROW)
    dest.getRange(DEST_FIRST_ROW, 2,
                  dest.getLastRow() - DEST_FIRST_ROW + 1, 4).clearContent();
  dest.getRange(SCHED_START_ROW, 7,
                dest.getMaxRows() - SCHED_START_ROW + 1, 4).clearContent();

  /* write agenda rows B-E */
  if (rows.length)
    dest.getRange(DEST_FIRST_ROW, 2, rows.length, 4).setValues(rows);

  /* write fresh actor list in G (schedule table) */
  const actors = [...new Set(rows.map(r => r[3]).filter(Boolean))].sort();
  if (actors.length)
    dest.getRange(SCHED_START_ROW, 7, actors.length, 1)
        .setValues(actors.map(a => [a]));

  /* sort agenda by Actor col E */
  if (rows.length)
    dest.getRange(DEST_FIRST_ROW, 2, rows.length, 4)
        .sort({ column: 5, ascending: true });
}

/* ---- Add to Agenda ---- */
function appendAgenda() {
  const addRows = getQuickRows_();
  if (!addRows.length) return;

  const dest = SpreadsheetApp.getActive().getSheetByName(TAB_DEST);

  /* append B-E */
  const startRow = Math.max(dest.getLastRow() + 1, DEST_FIRST_ROW);
  dest.getRange(startRow, 2, addRows.length, 4).setValues(addRows);

  /* schedule table: keep existing timings, add NEW actors only */
  const sched = dest.getRange(SCHED_START_ROW, 7,
                              dest.getMaxRows() - SCHED_START_ROW + 1, 4).getValues();
  const have  = sched.map(r => r[0]).filter(Boolean);
  const newActs = [...new Set(addRows.map(r => r[3]).filter(Boolean))]
                  .filter(a => !have.includes(a));
  if (newActs.length) {
    let rowPtr = sched.findIndex(r => !r[0]);          // first blank slot
    if (rowPtr < 0) rowPtr = sched.length;
    dest.getRange(SCHED_START_ROW + rowPtr, 7, newActs.length, 1)
        .setValues(newActs.map(a => [a]));
  }

  /* resort agenda */
  dest.getRange(DEST_FIRST_ROW, 2,
                dest.getLastRow() - DEST_FIRST_ROW + 1, 4)
      .sort({ column: 5, ascending: true });
}
