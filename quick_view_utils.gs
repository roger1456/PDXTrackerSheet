/* ===== QUICK-VIEW → agenda helpers ===== */

const TAB_SRC       = 'Quick View';
const TAB_CHARLIST  = 'Character List';
const SRC_FIRST_ROW = 5;

/* Build Map<character-lower, actor> once per run */
function buildActorMap_() {
  const rows = SpreadsheetApp.getActive()
      .getSheetByName(TAB_CHARLIST)
      .getRange('C:D')
      .getValues();

  const map = new Map();
  rows.forEach(([char, actor]) => {
    if (char && actor) map.set(char.toString().toLowerCase(), actor.toString());
  });
  return map;
}

/* Pull A-D from Quick View (row 5↓) and resolve actors */
function getQuickRows_() {
  const sh   = SpreadsheetApp.getActive().getSheetByName(TAB_SRC);
  const last = sh.getLastRow();
  if (last < SRC_FIRST_ROW) return [];

  const raw  = sh.getRange(SRC_FIRST_ROW, 1, last-SRC_FIRST_ROW+1, 4).getValues();
  const map  = buildActorMap_();
  const out  = [];

  raw.forEach(([seq, seqName, charName, lines]) => {
    if (!seq) return;                                // stop at first blank seq
    const seqAnd = `${seq} ${seqName}`.trim();
    const key    = charName.toString().toLowerCase();

    let actor = map.get(key);
    if (!actor) {                                    // partial match fallback
      for (const [k, v] of map) { if (k.includes(key)) { actor = v; break; } }
    }
    out.push([seqAnd, charName, lines, actor || '—']);
  });
  return out;
}
