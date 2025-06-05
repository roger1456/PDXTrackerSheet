
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== TAB_DEST) return;      // not Agenda Builder
  const col = e.range.getColumn();
  if (col !== 8 && col !== 9 && col !== 10) return; // only H/I/J edits

  /* session start time from E1 ("… – 1:00 PM–…") */
  const sessionVal = sh.getRange('E1').getValue();
  const sm = sessionVal.match(/–\s.+?–\s+(\d{1,2}:\d{2}\s*[AP]M)/i);
  if (!sm) return;

  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy ');

  const toDate = tStr => new Date(todayStr + tStr);      // helper

  const sessionStart = toDate(sm[1]);
  if (isNaN(sessionStart)) return;

  /* pull schedule G-J rows */
  const rng   = sh.getRange(SCHED_START_ROW, 7,
                            sh.getLastRow() - SCHED_START_ROW + 1, 4);
  const rows  = rng.getValues();

  /* collect rows with actor + order + duration */
  const tasks = rows.map((r,i) => ({
      idx   : i,
      actor : r[0],
      order : Number(r[1]),
      dur   : r[2],
      range : r[3]            // existing "h:mm a-h:mm a" or blank
  }))
  .filter(t => t.actor && t.order && t.dur)
  .sort((a,b) => a.order - b.order);

  if (!tasks.length) return;

  const toMin = v => {                              // duration → minutes
    if (v instanceof Date)  return v.getHours()*60 + v.getMinutes();
    const n = Number(v);     if (!isNaN(n)) return n;      // "30"
    const mm = String(v).match(/(\d+):(\d+)/);             // "1:15"
    if (mm) return (+mm[1])*60 + (+mm[2]);
    return 0;
  };

  /* iterate in order – honour any user-typed range */
  let cursor = new Date(sessionStart);

  tasks.forEach(t => {
    let start  = cursor;
    let end;

    /* If user already provided a valid range, KEEP it */
    const m = t.range && t.range.match(/(\d{1,2}:\d{2}\s*[AP]M)\s*-\s*(\d{1,2}:\d{2}\s*[AP]M)/i);
    if (m) {
      start = toDate(m[1]);                          // user’s start
      end   = toDate(m[2]);                          // user’s end
      if (isNaN(start) || isNaN(end) || end <= start) { /* ignore bad */ }
      else {
        cursor = new Date(end);                      // advance cursor
        return;                                     // keep existing value
      }
    }

    /* otherwise, compute range from cursor + duration */
    const mins = toMin(t.dur);
    if (!mins) return;                               // skip if bad duration

    const startStr = Utilities.formatDate(start, tz, 'h:mm a');
    end = new Date(start.getTime() + mins*60000);
    const endStr   = Utilities.formatDate(end,   tz, 'h:mm a');

    rows[t.idx][3] = `${startStr}-${endStr}`;        // write into col J
    cursor = end;                                    // advance cursor
  });

  /* write back only col J */
  const jCol = rows.map(r => [r[3]]);
  sh.getRange(SCHED_START_ROW, 10, jCol.length, 1).setValues(jCol);
}
