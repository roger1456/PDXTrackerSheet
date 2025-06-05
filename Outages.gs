/**
 * Pulls full-day “OOO” events from your outage calendar (today → +1 month),
 * looks up each username in columns G:H of the same sheet, sorts by date,
 * and writes [Full Name, MM/dd – MM/dd] into A3:B…
 */
function importTalentOutages() {
  const CAL_ID = 'c_07e77fc5cdccab79cce60bffbc32a9f20114411c0098ed39cdf89acc7b04af55@group.calendar.google.com';
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheet  = ss.getSheetByName('RDX Talent Outages');
  const cal    = CalendarApp.getCalendarById(CAL_ID);

  // 1️⃣ Build date window: today @ 00:00 → one month out @ 23:59:59
  const now   = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const end   = new Date(start);
  end.setMonth(end.getMonth() + 1);
  end.setHours(23, 59, 59);

  // 2️⃣ Fetch events with “OOO” in title, filter to all-day
  const events = cal.getEvents(start, end, { search: 'OOO' })
                    .filter(evt => evt.isAllDayEvent());

  // 3️⃣ Map → { start, name, dates }
  const outages = events.map(evt => {
    const parts    = evt.getTitle().split(' ');
    const username = parts[0];
    const fullName = fetchFullName_(username, sheet);

    const sDate = evt.getAllDayStartDate();
    const eRaw  = evt.getAllDayEndDate();                    // exclusive
    const eDate = new Date(eRaw.getTime() - 24*60*60*1000);  // subtract a day

    const fmt  = Utilities.formatDate;
    const tz   = Session.getScriptTimeZone();
    const rangeText = fmt(sDate, tz, 'MM/dd') + ' – ' + fmt(eDate, tz, 'MM/dd');

    return { start: sDate, name: fullName, dates: rangeText };
  });

  // 4️⃣ Sort by start date
  outages.sort((a, b) => a.start - b.start);

  // 5️⃣ Write into A3:B… (clear old first)
  sheet.getRange('A3:B').clearContent();
  if (outages.length) {
    const rows = outages.map(o => [o.name, o.dates]);
    sheet.getRange(3, 1, rows.length, 2).setValues(rows);
  }
}

/**
 * Looks up username → full name from columns G (7) and H (8)
 * on the provided sheet. Falls back to username if no match.
 */
function fetchFullName_(username, sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return username;

  // Read G2:H[lastRow]
  const data = sheet.getRange(2, 7, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === username) {
      return data[i][1] || username;
    }
  }
  return username;
}
