/**
 * Refreshes the dropdown of upcoming sessions.
 */
function populateRecordingSessionsDropdown() {
  const CALENDAR_ID = 'c_cff91665badcd301942da4bf691d83dfecba2fad5027a030db9587d8e7bb3bef@group.calendar.google.com';
  const SHEET_NAME  = 'Agenda Builder';
  const CELL_ADDR   = 'E1';

  const today     = new Date();
  const fourWeeks = new Date(today);
  fourWeeks.setDate(today.getDate() + 28);

  const events = CalendarApp
    .getCalendarById(CALENDAR_ID)
    .getEvents(today, fourWeeks)
    .filter(ev => {
      const t = ev.getTitle();
      return /(RDX|PDX|CDX)/i.test(t) && !/PREP/i.test(t);
    })
    .sort((a,b) => a.getStartTime() - b.getStartTime());

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const items = events.map(ev => {
    const start = ev.getStartTime(), end = ev.getEndTime();
    return Utilities.formatDate(start,tz,'MM/dd') +
      ' – ' + ev.getTitle() +
      ' – ' + Utilities.formatDate(start,tz,'h:mm a') +
      '–' + Utilities.formatDate(end,  tz,'h:mm a');
  });

  if (!items.length) {
    Logger.log('No matching events found.');
    return;
  }
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const cell  = ss.getSheetByName(SHEET_NAME).getRange(CELL_ADDR);
  cell.clearDataValidations()
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(items,true)
          .setAllowInvalid(false)
          .build()
      )
      .setValue(items[0]);
}

