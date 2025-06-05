/* ==========================================================
   UPDATE CAL  — validates the schedule block (G-J),
                 checks the “prep” calendar against it,
                 notifies the user what’s missing / wrong, and
                 (if everything lines up) rewrites the AGENDA
                 section in the main session event’s description.
   ========================================================== */

const PREP_CAL_ID  = 'c_be742fbaaf9fe3ebac7b2d13338356c18d3bb4fcb01b6e602b4a1257287ee030@group.calendar.google.com';
const MAIN_CAL_ID  = 'c_cff91665badcd301942da4bf691d83dfecba2fad5027a030db9587d8e7bb3bef@group.calendar.google.com';
const TZ           = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

/* ---------------- PUBLIC button -------------------------- */
function updateCalendar() {
  const ss  = SpreadsheetApp.getActive();
  const sh  = ss.getSheetByName(TAB_DEST);           // from earlier file
  const ui  = SpreadsheetApp.getUi();

  /* 1 — validate schedule table completeness (G-J) */
  const sched = sh.getRange(SCHED_START_ROW, 7,
                            sh.getLastRow() - SCHED_START_ROW + 1, 4)
                  .getValues();

  const missing = [];
  sched.forEach((r, i) => {
    if (!r[0]) return;                               // no actor → ignore row
    if (!r[1] || !r[2] || !r[3])
      missing.push(`Row ${SCHED_START_ROW + i}: ${r[0] || '(blank actor)'}`);
  });
  if (missing.length) {
    ui.alert('Fill in every column for these rows first:\n\n' + missing.join('\n'));
    return;
  }

  /* 2 — locate the session event from E1 */
  const e1 = sh.getRange('E1').getValue();
  const sm = e1.match(/–\s+(.*?)\s+–/);
  if (!sm) { ui.alert('Bad session string in E1.'); return; }
  const sessCode = (sm[1].match(/(RDX|PDX|CDX)\d{3}/i) || [])[0];
  if (!sessCode) { ui.alert('Cannot parse session code in E1.'); return; }

  const evMain = CalendarApp.getCalendarById(MAIN_CAL_ID)
                 .getEvents(new Date(), new Date(Date.now() + 28*864e5))
                 .find(e => e.getTitle().includes(sessCode));
  if (!evMain) { ui.alert('Main session event not found on calendar.'); return; }

  const isSJRS = /Steve Jobs/i.test(evMain.getLocation() || '');
  const studioCode = isSJRS ? 'SJRS' : 'BKRS';

  /* 3 — build expected prep-events list from schedule rows */
  const dayStamp   = Utilities.formatDate(evMain.getStartTime(), TZ, 'MM/dd/yyyy ');
  const toDate = tStr => new Date(dayStamp + tStr);
  const expected = [];                                // [{start,end,title}]
  sched.forEach(r => {
    if (!r[0]) return;                                // blank actor row
    const m = r[3].match(/(.+)-(.+)/);                // "4:45 PM-5:00 PM"
    if (!m) return;
    expected.push({
      actor : r[0],
      start : toDate(m[1].trim()),
      end   : toDate(m[2].trim()),
      title : `${r[0]} as ${findCharForActor_(r[0])} @ ${studioCode}`
    });
  });

  /* 4 — scan prep calendar for matching events */
  const prepEvents = CalendarApp.getCalendarById(PREP_CAL_ID)
                     .getEvents(evMain.getStartTime(), evMain.getEndTime());

  const toAdd = [], toAdjust = [];

  expected.forEach(exp => {
    /* any prep event that mentions this actor? */
    const byName = prepEvents.filter(evt =>
        evt.getTitle().toLowerCase().includes(exp.actor.toLowerCase()));

    if (!byName.length) {                       // none at all → MISSING
      toAdd.push(exp);
      return;
    }

    /* check if at least one has the exact time */
    const good = byName.find(evt =>
        Math.abs(evt.getStartTime() - exp.start) < 60000 &&
        Math.abs(evt.getEndTime()   - exp.end)   < 60000);

    if (!good) toAdjust.push(exp);              // name exists but time off → WRONG
  });

  if (toAdd.length || toAdjust.length) {
    const msg =
      (toAdd.length    ? ('Missing events:\n'   + toAdd   .map(fmt_).join('\n') + '\n\n') : '') +
      (toAdjust.length ? ('Wrong-time events:\n'   + toAdjust.map(fmt_).join('\n'))          : '');
    ui.alert('Prep calendar needs attention:\n\n' + msg);
    return;
  }

  /* 5 — everything matches: refresh AGENDA block in main event */
  const agendaHtml = buildAgendaHtml_(sh, studioCode);
  const cleaned = (evMain.getDescription() || '')
                  .replace(/<b><u>AGENDA[\s\S]*$/i, '');  // drop old agenda
  evMain.setDescription(cleaned + agendaHtml);

  ui.alert('Calendar updated – AGENDA refreshed in session event.');
}

/* helper: actor → char (first match in agenda table) */
function findCharForActor_(actor) {
  const rows = SpreadsheetApp.getActive()
      .getSheetByName(TAB_DEST)
      .getRange(DEST_FIRST_ROW, 2,
                SpreadsheetApp.getActive().getSheetByName(TAB_DEST).getLastRow() - DEST_FIRST_ROW + 1,
                4)
      .getValues();
  const hit = rows.find(r => r[3] === actor);
  return hit ? hit[1] : '';
}

/* helper: pretty string */
function fmt_(o) {
  const s = Utilities.formatDate(o.start, TZ, 'h:mm a');
  const e = Utilities.formatDate(o.end  , TZ, 'h:mm a');
  return `• ${s}-${e}  ${o.title}`;
}

/* helper: rebuild AGENDA block */
function buildAgendaHtml_(sheet, studioCode) {
  const br = '<br>', b = s=>'<b>'+s+'</b>', i=s=>'<i>'+s+'</i>';

  /* timing rows sorted by start-time */
  const sched = sheet.getRange(SCHED_START_ROW, 7,
                                sheet.getLastRow() - SCHED_START_ROW + 1, 4)
                     .getValues()
                     .filter(r=>r[0] && r[3])
                     .map(r=>{
                       const m = r[3].match(/(.+)-(.+)/);
                       return {
                         actor    : r[0],
                         startStr : m?m[1].trim():'',
                         endStr   : m?m[2].trim():''
                       };
                     })
                     .sort((a,b)=>{
                       const d1 = new Date('1/1/2000 '+a.startStr);
                       const d2 = new Date('1/1/2000 '+b.startStr);
                       return d1 - d2;
                     });

  /* actor → char & lines */
  const rows = sheet.getRange(DEST_FIRST_ROW,2,
                              sheet.getLastRow()-DEST_FIRST_ROW+1,4).getValues()
                    .filter(r=>r[0]);
  const amap=new Map();
  rows.forEach(([seq,char,lines,actor])=>{
    if(!amap.has(actor))amap.set(actor,{char:char.toUpperCase(),total:0,items:[]});
    const o=amap.get(actor);
    o.total+=+lines;
    o.items.push({seq,lines:+lines});
  });

  /* build HTML */
  let html = '<b><u>AGENDA</u></b>'+br;
  sched.forEach(t=>{
    html += '• ' + i(t.startStr + ' - ' + t.endStr) + ' : ';
    const d = amap.get(t.actor) || {char:'',total:0};
    html += t.actor + ' as ' + b(d.char) + br;         
  });
  return html;
}
