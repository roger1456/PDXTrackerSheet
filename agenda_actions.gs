// Original functions newAgenda() and appendAgenda() are now commented out (on 2023-10-27 by AI Assistant)
// as their functionality is replaced by the "Manage Agenda Item(s)" dialog
// and direct interaction with the "AgendaBackend" sheet.
// The "Quick View" sheet is no longer the direct source for populating "Agenda Builder".

/* ===== New & Add buttons for Agenda Builder ===== */

/*
const TAB_DEST        = 'Agenda Builder';
const DEST_FIRST_ROW  = 5;
const SCHED_START_ROW = 5;

/* ---- New Agenda ---- * /
function newAgenda() {
  const rows = getQuickRows_();                       // util from quick_view_utils.gs
  const dest = SpreadsheetApp.getActive().getSheetByName(TAB_DEST);

  /* clear old B-E and G-J content (keeps formatting) * /
  if (dest.getLastRow() >= DEST_FIRST_ROW)
    dest.getRange(DEST_FIRST_ROW, 2,
                  dest.getLastRow() - DEST_FIRST_ROW + 1, 4).clearContent();
  dest.getRange(SCHED_START_ROW, 7,
                dest.getMaxRows() - SCHED_START_ROW + 1, 4).clearContent();

  /* write agenda rows B-E * /
  if (rows.length)
    dest.getRange(DEST_FIRST_ROW, 2, rows.length, 4).setValues(rows);

  /* write fresh actor list in G (schedule table) * /
  const actors = [...new Set(rows.map(r => r[3]).filter(Boolean))].sort();
  if (actors.length)
    dest.getRange(SCHED_START_ROW, 7, actors.length, 1)
        .setValues(actors.map(a => [a]));

  /* sort agenda by Actor col E * /
  if (rows.length)
    dest.getRange(DEST_FIRST_ROW, 2, rows.length, 4)
        .sort({ column: 5, ascending: true });
}

/* ---- Add to Agenda ---- * /
function appendAgenda() {
  const addRows = getQuickRows_();
  if (!addRows.length) return;

  const dest = SpreadsheetApp.getActive().getSheetByName(TAB_DEST);

  /* append B-E * /
  const startRow = Math.max(dest.getLastRow() + 1, DEST_FIRST_ROW);
  dest.getRange(startRow, 2, addRows.length, 4).setValues(addRows);

  /* schedule table: keep existing timings, add NEW actors only * /
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

  /* resort agenda * /
  dest.getRange(DEST_FIRST_ROW, 2,
                dest.getLastRow() - DEST_FIRST_ROW + 1, 4)
      .sort({ column: 5, ascending: true });
}
*/

// ===== Code for Manage Agenda Item Dialog =====

const AGENDA_BACKEND_SHEET_NAME = "AgendaBackend";
// TAB_DEST from above is 'Agenda Builder', so we can use that or this new constant.
// For consistency within these new functions, AGENDA_BUILDER_SHEET_NAME is used.
// Note: The original TAB_DEST, DEST_FIRST_ROW, SCHED_START_ROW are now commented out above.
// If any active code below still relies on them being defined at the top level of this file,
// they might need to be re-declared or those functions updated to use AGENDA_BUILDER_SHEET_NAME etc.
// For now, assuming the below code uses its own constants or ones passed/defined globally elsewhere.
// The constants AGENDA_BUILDER_SHEET_NAME, EVENT_CELL_IN_BUILDER are defined here.
// TAB_DEST, DEST_FIRST_ROW, SCHED_START_ROW are used by deleteSelectedItems, mapping to these new constants.

const AGENDA_BUILDER_SHEET_NAME = "Agenda Builder";
const EVENT_CELL_IN_BUILDER = "E1";

// --- Helper function to generate UniqueID ---
function generateUniqueID_() {
  return Utilities.getUuid();
}

// --- Function to add menu item ---
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Agenda Tools')
      .addItem('Manage Agenda Item(s)', 'showManageAgendaDialog')
      .addItem('Delete Selected Item(s)', 'deleteSelectedItems') // New item
      .addSeparator()
      .addItem('Populate Event Dropdown', 'populateRecordingSessionsDropdown') // Assuming this is still useful
      .addToUi();
}

// --- Function to show the dialog ---
function showManageAgendaDialog() {
  const htmlOutput = HtmlService.createTemplateFromFile('ManageAgendaDialog')
      .evaluate()
      .setWidth(600)
      .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Manage Agenda Item(s)');
}

// --- Function to get events for the dialog dropdown ---
function getEventsForDialog() {
  const CALENDAR_ID = 'c_cff91665badcd301942da4bf691d83dfecba2fad5027a030db9587d8e7bb3bef@group.calendar.google.com';
  const today = new Date();
  const fourWeeks = new Date(today);
  fourWeeks.setDate(today.getDate() + 28);
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  try {
    const events = CalendarApp
      .getCalendarById(CALENDAR_ID)
      .getEvents(today, fourWeeks)
      .filter(ev => {
        const t = ev.getTitle();
        return /(RDX|PDX|CDX)/i.test(t) && !/PREP/i.test(t);
      })
      .sort((a,b) => a.getStartTime() - b.getStartTime());

    const items = events.map(ev => {
      const start = ev.getStartTime(), end = ev.getEndTime();
      return Utilities.formatDate(start,tz,'MM/dd') +
        ' – ' + ev.getTitle() +
        ' – ' + Utilities.formatDate(start,tz,'h:mm a') +
        '–' + Utilities.formatDate(end,  tz,'h:mm a');
    });

    Logger.log(`getEventsForDialog: Found ${items.length} events.`);
    return items.length > 0 ? items : ["No events found in the next 4 weeks"];
  } catch (e) {
    Logger.log(`Error in getEventsForDialog: ${e.toString()}`);
    return [`Error fetching events: ${e.toString()}`];
  }
}

// --- Function to process the form submission ---
function processManageAgendaForm(formData) {
  try {
    Logger.log(`Processing form data: ${JSON.stringify(formData)}`);
    const eventID = formData.eventSelect;
    const action = formData.actionType; // 'add' or 'replace'

    const items = [];
    if (formData.topic) { // Basic check for a single item's main field
        items.push({
            sequence: formData.sequence,
            topic: formData.topic,
            details: formData.details,
            actor: formData.actor
        });
    } else if (formData.items && typeof formData.items === 'string') {
        try {
            const parsedItems = JSON.parse(formData.items);
            if (Array.isArray(parsedItems)) {
                items.push(...parsedItems);
            }
        } catch (e) {
            Logger.log(`Error parsing items JSON: ${e.toString()}`);
            return { success: false, message: "Error processing items. Invalid format." };
        }
    }

    if (!eventID || !action || (items.length === 0 && action === 'add')) { // Items needed for 'add'
      Logger.log("Missing data: eventID, action, or items array is empty for 'add' action.");
      return { success: false, message: "Missing event, action, or item details for 'add' action." };
    }
    if (!eventID || !action ) { // General check
      Logger.log("Missing data: eventID or action.");
      return { success: false, message: "Missing event or action type." };
    }


    const backendSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AGENDA_BACKEND_SHEET_NAME);
    if (!backendSheet) {
      Logger.log(`Error: Sheet ${AGENDA_BACKEND_SHEET_NAME} not found.`);
      return { success: false, message: `Backend sheet ${AGENDA_BACKEND_SHEET_NAME} not found.` };
    }

    const backendHeadersRange = backendSheet.getRange(1, 1, 1, backendSheet.getMaxColumns());
    const backendHeaders = backendHeadersRange.getValues()[0];
    const eventIDColIdx = backendHeaders.indexOf("EventID");
    const uniqueIDColIdx = backendHeaders.indexOf("UniqueID");
    const sequenceColIdx = backendHeaders.indexOf("Sequence");
    const topicColIdx = backendHeaders.indexOf("Topic/Action");
    const detailsColIdx = backendHeaders.indexOf("Details/Lines");
    const actorColIdx = backendHeaders.indexOf("Actor/Owner");
    const originalOrderColIdx = backendHeaders.indexOf("OriginalOrder");

    if (eventIDColIdx === -1 || uniqueIDColIdx === -1 || sequenceColIdx === -1 || topicColIdx === -1 || detailsColIdx === -1 || actorColIdx === -1 || originalOrderColIdx === -1) {
       Logger.log(`Error: One or more required columns not found in ${AGENDA_BACKEND_SHEET_NAME}. Missing: EventID, UniqueID, Sequence, Topic/Action, Details/Lines, Actor/Owner, or OriginalOrder.`);
      return { success: false, message: `One or more key columns are missing in ${AGENDA_BACKEND_SHEET_NAME}. Setup is incomplete.` };
    }

    if (action === "replace") {
      Logger.log(`Action: Replace. Deleting items for event: ${eventID}`);
      const data = backendSheet.getDataRange().getValues();
      const rowsToDelete = [];
      for (let i = data.length - 1; i >= 1; i--) { // Skip header, iterate backwards
        if (data[i][eventIDColIdx] === eventID) {
          rowsToDelete.push(i + 1);
        }
      }
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach(rowIndex => {
        backendSheet.deleteRow(rowIndex);
        Logger.log(`Deleted row ${rowIndex} from ${AGENDA_BACKEND_SHEET_NAME} for event ${eventID}`);
      });
    }

    let currentMaxOriginalOrder = 0;
    if (action === "add" || action === "replace") { // Need to calculate for both, for 'replace' it starts from 0.
        const allData = backendSheet.getDataRange().getValues();
        allData.slice(1).forEach(row => { // Skip header
            if (row[eventIDColIdx] === eventID) { // Only for the current event
                const orderVal = parseInt(row[originalOrderColIdx]);
                if (!isNaN(orderVal) && orderVal > currentMaxOriginalOrder) {
                    currentMaxOriginalOrder = orderVal;
                }
            }
        });
    }

    const newRowsData = [];
    items.forEach(item => {
      if (!item.topic && action === 'add') return; // Skip adding empty items
      currentMaxOriginalOrder++;
      newRowsData.push([
        eventID,
        generateUniqueID_(),
        item.sequence || "", // Sequence from form
        item.topic || "",    // Topic from form
        item.details || "",  // Details from form
        item.actor || "",    // Actor from form
        currentMaxOriginalOrder // Calculated OriginalOrder
      ]);
    });

    if (newRowsData.length > 0) {
      backendSheet.getRange(backendSheet.getLastRow() + 1, 1, newRowsData.length, newRowsData[0].length).setValues(newRowsData);
      Logger.log(`Added ${newRowsData.length} new rows to ${AGENDA_BACKEND_SHEET_NAME}.`);
    }

    const agendaBuilderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AGENDA_BUILDER_SHEET_NAME);
    if (agendaBuilderSheet) {
      const currentEventInBuilder = agendaBuilderSheet.getRange(EVENT_CELL_IN_BUILDER).getValue();
      if (currentEventInBuilder === eventID) {
        Logger.log(`Current event in builder matches modified event (${eventID}). Attempting to trigger refresh.`);
        const e1Cell = agendaBuilderSheet.getRange(EVENT_CELL_IN_BUILDER);
        const originalValueE1 = e1Cell.getValue();
        e1Cell.setValue(Utilities.getUuid());
        SpreadsheetApp.flush();
        e1Cell.setValue(originalValueE1);
        SpreadsheetApp.flush();
        Logger.log(`Attempted to refresh Agenda Builder by re-setting ${EVENT_CELL_IN_BUILDER} for event ${eventID}.`);
      }
    }

    return { success: true, message: `Agenda for '${eventID}' updated. ${newRowsData.length} items processed.` };
  } catch (e) {
    Logger.log(`Error in processManageAgendaForm: ${e.toString()} Stack: ${e.stack}`);
    return { success: false, message: `An error occurred: ${e.toString()}` };
  }
}

// --- Function to delete selected items ---
function deleteSelectedItems() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  // Constants for deleteSelectedItems - these will map to the global ones defined or available
  const AGENDA_BUILDER_SHEET_NAME_CONST = AGENDA_BUILDER_SHEET_NAME; // Uses 'Agenda Builder' from this file
  const DEST_FIRST_ROW_CONST = 5; // Assuming 5, needs to be consistent if used by other functions
  const HELPER_UNIQUE_ID_COL_CONST = 11; // Column K
  const BACKEND_SHEET_NAME_CONST = AGENDA_BACKEND_SHEET_NAME; // Uses 'AgendaBackend' from this file
  const EVENT_CELL_IN_BUILDER_CONST = EVENT_CELL_IN_BUILDER; // Uses 'E1' from this file
  const SCHED_START_ROW_CONST = 5; // Assuming 5


  if (activeSheet.getName() !== AGENDA_BUILDER_SHEET_NAME_CONST) {
    ui.alert("Please select items on the 'Agenda Builder' sheet to delete.");
    return;
  }

  const currentEventID = activeSheet.getRange(EVENT_CELL_IN_BUILDER_CONST).getValue();
  if (!currentEventID) {
    ui.alert("No event selected in cell E1. Cannot determine which event's items to delete from the backend.");
    return;
  }

  const selection = activeSheet.getSelection();
  const selectedRanges = selection.getActiveRangeList().getRanges();

  if (!selectedRanges || selectedRanges.length === 0) {
    ui.alert("Please select one or more rows of agenda items to delete.");
    return;
  }

  const uniqueIDsToDelete = new Set();
  const rowsToDeleteInBuilder = new Set();

  for (const range of selectedRanges) {
    const startRow = range.getRow();
    const endRow = range.getLastRow();

    for (let r = startRow; r <= endRow; r++) {
      if (r < DEST_FIRST_ROW_CONST) continue;

      const uniqueID = activeSheet.getRange(r, HELPER_UNIQUE_ID_COL_CONST).getValue();
      if (uniqueID) {
        uniqueIDsToDelete.add(uniqueID);
        rowsToDeleteInBuilder.add(r);
      } else {
        Logger.log(`Row ${r} selected for deletion but no UniqueID found in column K. Skipping.`);
      }
    }
  }

  if (uniqueIDsToDelete.size === 0) {
    ui.alert("No valid agenda items selected for deletion (could not find UniqueIDs for selected rows).");
    return;
  }

  const confirm = ui.alert(
    'Confirm Deletion',
    `Are you sure you want to delete ${uniqueIDsToDelete.size} selected agenda item(s) from event '${currentEventID}'? This action cannot be undone.`,
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }

  Logger.log(`Proceeding with deletion. UniqueIDs to delete: ${Array.from(uniqueIDsToDelete).join(', ')} for event ${currentEventID}`);

  const backendSheet = ss.getSheetByName(BACKEND_SHEET_NAME_CONST);
  if (!backendSheet) {
    ui.alert(`Error: Backend sheet '${BACKEND_SHEET_NAME_CONST}' not found.`);
    Logger.log(`Error: Backend sheet '${BACKEND_SHEET_NAME_CONST}' not found during deletion.`);
    return;
  }

  const backendDataRange = backendSheet.getDataRange();
  const backendAllData = backendDataRange.getValues();
  const backendHeaders = backendAllData.shift();
  const backendUniqueIDColIdx = backendHeaders.indexOf("UniqueID");
  const backendEventIDColIdx = backendHeaders.indexOf("EventID");
  const backendActorColIdx = backendHeaders.indexOf("Actor/Owner");

  if (backendUniqueIDColIdx === -1 || backendEventIDColIdx === -1 || backendActorColIdx === -1) {
    ui.alert("Error: Critical columns (UniqueID, EventID, or Actor/Owner) not found in backend sheet.");
    Logger.log("Error: Critical columns (UniqueID, EventID, or Actor/Owner) not found in backend sheet during deletion.");
    return;
  }

  let backendRowsDeletedCount = 0;
  const remainingBackendDataForEvent = [];
  const rowsToDeleteFromBackendSheet = [];

  for (let i = 0; i < backendAllData.length; i++) {
    const row = backendAllData[i];
    const rowUniqueID = row[backendUniqueIDColIdx];
    const rowEventID = row[backendEventIDColIdx];

    if (rowEventID === currentEventID && uniqueIDsToDelete.has(rowUniqueID)) {
      rowsToDeleteFromBackendSheet.push(i + 2);
      backendRowsDeletedCount++;
    } else if (rowEventID === currentEventID) {
      remainingBackendDataForEvent.push(row);
    }
  }

  rowsToDeleteFromBackendSheet.sort((a,b) => b-a).forEach(rowIndex => {
      backendSheet.deleteRow(rowIndex);
      Logger.log(`Deleted row ${rowIndex} from backend.`);
  });

  Logger.log(`Deleted ${backendRowsDeletedCount} rows from ${BACKEND_SHEET_NAME_CONST}.`);

  const sortedRowsToDeleteInBuilder = Array.from(rowsToDeleteInBuilder).sort((a, b) => b - a);
  sortedRowsToDeleteInBuilder.forEach(rowNum => {
    activeSheet.deleteRow(rowNum);
    Logger.log(`Deleted row ${rowNum} from ${AGENDA_BUILDER_SHEET_NAME_CONST}.`);
  });

  Logger.log("Refreshing actor list and schedule after deletion.");
  const uniqueActors = new Set();
  remainingBackendDataForEvent.forEach(itemRow => {
      if (itemRow[backendActorColIdx]) {
          uniqueActors.add(itemRow[backendActorColIdx]);
      }
  });

  const numRowsToClearInG = Math.max(0, activeSheet.getLastRow() - SCHED_START_ROW_CONST + 1);
  if (numRowsToClearInG > 0) {
      activeSheet.getRange(SCHED_START_ROW_CONST, 7, numRowsToClearInG, 1).clearContent();
      Logger.log(`Cleared column G from row ${SCHED_START_ROW_CONST} for actor refresh.`);
  }

  const sortedActors = Array.from(uniqueActors).sort().map(actor => [actor]);
  if (sortedActors.length > 0) {
    activeSheet.getRange(SCHED_START_ROW_CONST, 7, sortedActors.length, 1).setValues(sortedActors);
    Logger.log(`Refreshed column G with ${sortedActors.length} actors after deletion for event ${currentEventID}.`);
  } else {
    Logger.log(`No actors found for event ${currentEventID} after deletion. Column G is empty.`);
  }

  if (typeof recalculateScheduleTimingsForSheet === "function") {
    recalculateScheduleTimingsForSheet(activeSheet);
    Logger.log("Called recalculateScheduleTimingsForSheet after deletions.");
  } else {
    Logger.log("Warning: recalculateScheduleTimingsForSheet function not found. Schedule may be stale.");
    ui.alert("Items deleted, but schedule recalculation function was not found. The schedule timings might be stale.");
  }

  ui.alert(`${uniqueIDsToDelete.size} item(s) deleted successfully from event '${currentEventID}'.`);
}
