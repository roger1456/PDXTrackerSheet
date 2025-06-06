// Constants for the "Agenda Builder" sheet operations
const TAB_DEST = 'Agenda Builder';
const DEST_FIRST_ROW = 5; // Starting row for agenda item population and clearing
const SCHED_START_ROW = 5;  // Starting row for schedule timing calculations (G-J)

/**
 * Main function triggered on edits in the spreadsheet.
 * @param {Object} e The event object.
 */
function onEdit(e) {
  // Call the primary logic handler
  newOnEditLogic(e);
}

/**
 * Handles the logic when an edit occurs, particularly for event selection
 * in E1 or timing adjustments in H, I, J.
 * @param {Object} e The event object.
 */
function newOnEditLogic(e) {
  const sh = e.range.getSheet();
  const editedCell = e.range;
  const editedRow = editedCell.getRow();
  const editedCol = editedCell.getColumn();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // Constants from the file (assuming they are globally available in schedule_onedit.gs)
  // const TAB_DEST = 'Agenda Builder'; // Already global
  // const DEST_FIRST_ROW = 5; // Already global
  // const SCHED_START_ROW = 5; // Already global
  const BACKEND_SHEET_NAME = 'AgendaBackend'; // Added for this logic
  const HELPER_UNIQUE_ID_COL = 11; // Column K

  if (sh.getName() !== TAB_DEST) return;
  Logger.log(`onEdit triggered in ${TAB_DEST} at ${editedCell.getA1Notation()} with value '${e.value}' (old value '${e.oldValue}')`);

  // Get current EventID from E1
  const currentEventID = sh.getRange('E1').getValue();
  if (!currentEventID) {
    Logger.log("No EventID selected in E1. Cannot process edit for agenda item persistence.");
    if (editedRow >= DEST_FIRST_ROW && editedCol >= 2 && editedCol <= 5) { // If edit is in agenda item rows
        ui.alert("No event selected in cell E1. Cannot save changes to this agenda item. Please select an event first.");
    }
    return; // Exit if no event is selected
  }

  // Scenario 1: Event Selection Change (E1)
  if (editedRow === 1 && editedCol === 5) { // E1 changed
    Logger.log(`Event selection changed in E1 to: ${currentEventID}`);
    const lastAgendaRow = sh.getLastRow();
    if (lastAgendaRow >= DEST_FIRST_ROW) {
      sh.getRange(DEST_FIRST_ROW, 2, lastAgendaRow - DEST_FIRST_ROW + 1, 4).clearContent(); // B-E
      sh.getRange(DEST_FIRST_ROW, 7, lastAgendaRow - DEST_FIRST_ROW + 1, 1).clearContent(); // G (Actors)
      sh.getRange(DEST_FIRST_ROW, 10, lastAgendaRow - DEST_FIRST_ROW + 1, 1).clearContent(); // J (Calculated Times)
      sh.getRange(DEST_FIRST_ROW, HELPER_UNIQUE_ID_COL, lastAgendaRow - DEST_FIRST_ROW + 1, 1).clearContent(); // K (Helper UniqueID)
    }

    const backendSheet = ss.getSheetByName(BACKEND_SHEET_NAME);
    if (!backendSheet) {
      Logger.log(`Error: ${BACKEND_SHEET_NAME} sheet not found.`);
      ui.alert(`Error: ${BACKEND_SHEET_NAME} sheet not found. Please set it up.`);
      return;
    }

    const backendDataRange = backendSheet.getDataRange();
    const backendAllData = backendDataRange.getValues(); // Get all data at once

    if (backendAllData.length === 0) {
        Logger.log(`${BACKEND_SHEET_NAME} sheet is empty.`);
        ui.alert(`${BACKEND_SHEET_NAME} sheet is empty. No data to load.`);
        return;
    }
    const backendHeaders = backendAllData.shift() || []; // Remove header row, get headers

    const eventIDColIdx_backend = backendHeaders.indexOf("EventID");
    const uniqueIDColIdx_backend = backendHeaders.indexOf("UniqueID");
    const sequenceColIdx_backend = backendHeaders.indexOf("Sequence");
    const topicColIdx_backend = backendHeaders.indexOf("Topic/Action");
    const detailsColIdx_backend = backendHeaders.indexOf("Details/Lines");
    const actorColIdx_backend = backendHeaders.indexOf("Actor/Owner");

    if (eventIDColIdx_backend === -1 || uniqueIDColIdx_backend === -1 || sequenceColIdx_backend === -1 || topicColIdx_backend === -1 || detailsColIdx_backend === -1 || actorColIdx_backend === -1 ) {
      Logger.log(`Error: Key columns not found in ${BACKEND_SHEET_NAME}. Required: EventID, UniqueID, Sequence, Topic/Action, Details/Lines, Actor/Owner.`);
      ui.alert(`Error: Key columns not found in ${BACKEND_SHEET_NAME}. Please check its structure.`);
      return;
    }

    const agendaItemsForEvent = backendAllData.filter(row => row[eventIDColIdx_backend] === currentEventID);
    Logger.log(`Found ${agendaItemsForEvent.length} items for event '${currentEventID}' in ${BACKEND_SHEET_NAME}.`);

    if (agendaItemsForEvent.length > 0) {
      const itemsToDisplay = [];
      const uniqueActors = new Set();
      const uniqueIDsForHelperColumn = [];

      agendaItemsForEvent.forEach(itemRow => {
        itemsToDisplay.push([
          itemRow[sequenceColIdx_backend] || "", itemRow[topicColIdx_backend] || "",
          itemRow[detailsColIdx_backend] || "", itemRow[actorColIdx_backend] || ""
        ]);
        if (itemRow[actorColIdx_backend]) uniqueActors.add(itemRow[actorColIdx_backend]);
        uniqueIDsForHelperColumn.push([itemRow[uniqueIDColIdx_backend] || ""]);
      });

      sh.getRange(DEST_FIRST_ROW, 2, itemsToDisplay.length, 4).setValues(itemsToDisplay);
      if (uniqueIDsForHelperColumn.length > 0) {
        sh.getRange(DEST_FIRST_ROW, HELPER_UNIQUE_ID_COL, uniqueIDsForHelperColumn.length, 1).setValues(uniqueIDsForHelperColumn);
      }
      const sortedActors = Array.from(uniqueActors).sort().map(actor => [actor]);
      if (sortedActors.length > 0) {
         // Clear existing actors in G before populating for new event
        const lastActorRowG = sh.getRange("G:G").getValues().filter(String).length;
        if (lastActorRowG >= SCHED_START_ROW) {
            sh.getRange(SCHED_START_ROW, 7, lastActorRowG - SCHED_START_ROW + 1, 1).clearContent();
        }
        sh.getRange(SCHED_START_ROW, 7, sortedActors.length, 1).setValues(sortedActors);
      } else { // No actors for this event, clear column G
        const lastActorRowG = sh.getRange("G:G").getValues().filter(String).length;
        if (lastActorRowG >= SCHED_START_ROW) {
            sh.getRange(SCHED_START_ROW, 7, lastActorRowG - SCHED_START_ROW + 1, 1).clearContent();
        }
      }
    } else { // No items for this event, ensure K and G are cleared below headers
        const lastAgendaRowK = sh.getRange("K:K").getValues().filter(String).length;
        if(lastAgendaRowK >= DEST_FIRST_ROW) {
            sh.getRange(DEST_FIRST_ROW, HELPER_UNIQUE_ID_COL, lastAgendaRowK - DEST_FIRST_ROW + 1, 1).clearContent();
        }
        const lastActorRowG = sh.getRange("G:G").getValues().filter(String).length;
        if (lastActorRowG >= SCHED_START_ROW) {
            sh.getRange(SCHED_START_ROW, 7, lastActorRowG - SCHED_START_ROW + 1, 1).clearContent();
        }
    }
    recalculateScheduleTimingsForSheet(sh);
    Logger.log('Agenda loaded for new event and schedule recalculated.');
    return;
  }

  // Scenario 2: Edit in Agenda Item Columns (B-E)
  if (editedRow >= DEST_FIRST_ROW && editedCol >= 2 && editedCol <= 5) {
    const uniqueIDInHelperCol = sh.getRange(editedRow, HELPER_UNIQUE_ID_COL).getValue();
    Logger.log(`Edit in agenda item area. Row: ${editedRow}, Col: ${editedCol}, UniqueID in K: '${uniqueIDInHelperCol}'`);

    if (uniqueIDInHelperCol) { // Existing item is being edited
      const backendSheet = ss.getSheetByName(BACKEND_SHEET_NAME);
      if (!backendSheet) { ui.alert(`${BACKEND_SHEET_NAME} not found.`); return; }

      const backendDataRange = backendSheet.getDataRange();
      const backendAllData = backendDataRange.getValues();
      const backendHeaders = backendAllData.shift(); // Get and remove header
      const uniqueIDColIdx_backend = backendHeaders.indexOf("UniqueID");
      const eventIDColIdx_backend = backendHeaders.indexOf("EventID");

      let targetBackendRowIndex = -1; // This will be 0-based for backendAllData array
      for(let i = 0; i < backendAllData.length; i++) {
        if (backendAllData[i][uniqueIDColIdx_backend] === uniqueIDInHelperCol && backendAllData[i][eventIDColIdx_backend] === currentEventID) {
          targetBackendRowIndex = i;
          break;
        }
      }

      if (targetBackendRowIndex !== -1) {
        let columnToUpdateInBackend = -1; // This is 0-based index for backendHeaders
        if (editedCol === 2) columnToUpdateInBackend = backendHeaders.indexOf("Sequence");
        else if (editedCol === 3) columnToUpdateInBackend = backendHeaders.indexOf("Topic/Action");
        else if (editedCol === 4) columnToUpdateInBackend = backendHeaders.indexOf("Details/Lines");
        else if (editedCol === 5) columnToUpdateInBackend = backendHeaders.indexOf("Actor/Owner");

        if (columnToUpdateInBackend !== -1) {
          // Update backend (targetBackendRowIndex is 0-based for data, so +2 for sheet row)
          backendSheet.getRange(targetBackendRowIndex + 2, columnToUpdateInBackend + 1).setValue(e.value);
          Logger.log(`Updated backend. Sheet Row: ${targetBackendRowIndex + 2}, Sheet Col: ${columnToUpdateInBackend + 1} ('${backendHeaders[columnToUpdateInBackend]}'), New Value: '${e.value}'`);

          if (editedCol === 5) { // Actor/Owner changed
            Logger.log("Actor/Owner changed. Refreshing actor list in Col G and recalculating schedule.");

            // Collect all unique actors for the current event from the backend
            const currentEventActors = new Set();
            backendAllData.forEach(row => {
                if (row[eventIDColIdx_backend] === currentEventID) {
                    // For the edited row, use the new value
                    if (row[uniqueIDColIdx_backend] === uniqueIDInHelperCol) {
                        if (e.value) currentEventActors.add(e.value);
                    } else { // For other rows, use their existing actor value
                        if (row[actorColIdx_backend]) currentEventActors.add(row[actorColIdx_backend]);
                    }
                }
            });

            const sortedActors = Array.from(currentEventActors).sort().map(actor => [actor]);

            const lastActorRowG = sh.getRange("G:G").getValues().filter(String).length;
            if (lastActorRowG >= SCHED_START_ROW) {
                sh.getRange(SCHED_START_ROW, 7, lastActorRowG - SCHED_START_ROW + 1, 1).clearContent();
            }
            if (sortedActors.length > 0) {
              sh.getRange(SCHED_START_ROW, 7, sortedActors.length, 1).setValues(sortedActors);
              Logger.log(`Refreshed column G with ${sortedActors.length} unique actors for the current event.`);
            }
            recalculateScheduleTimingsForSheet(sh);
          }
        }
      } else {
        Logger.log(`Error: Could not find item with UniqueID '${uniqueIDInHelperCol}' for event '${currentEventID}' in backend to update.`);
        // ui.alert("Error: Could not sync this change. Item not found in backend or event mismatch. Consider refreshing.");
        // e.range.setValue(e.oldValue); // Potentially revert? Risky.
      }
    } else { // No UniqueID in helper column K - This is a new item being entered directly
      const rowValues = sh.getRange(editedRow, 2, 1, 4).getValues()[0]; // B to E
      const hasContent = rowValues.some(cell => cell && cell.toString().trim() !== "");
      const currentEditAddsContent = e.value !== undefined && e.value !== null && e.value.toString().trim() !== "";

      if (hasContent || currentEditAddsContent) {
        Logger.log(`New item detected in row ${editedRow}. Content: ${rowValues.join(', ')}. Current edit: '${e.value}'`);
        const backendSheet = ss.getSheetByName(BACKEND_SHEET_NAME);
        if (!backendSheet) { ui.alert(`${BACKEND_SHEET_NAME} not found.`); return; }

        const backendDataRange = backendSheet.getDataRange();
        const backendAllDataWithHeaders = backendDataRange.getValues();
        const backendHeaders = backendAllDataWithHeaders[0].slice(); // Get a copy of headers
        const backendData = backendAllDataWithHeaders.slice(1); // Data without headers

        const eventIDColIdx_backend = backendHeaders.indexOf("EventID");
        const originalOrderColIdx_backend = backendHeaders.indexOf("OriginalOrder");
        const actorColIdx_backend = backendHeaders.indexOf("Actor/Owner"); // For actor list refresh

        let maxOrder = 0;
        backendData.forEach(r => {
            if (r[eventIDColIdx_backend] === currentEventID) {
                const order = parseInt(r[originalOrderColIdx_backend]);
                if (!isNaN(order) && order > maxOrder) maxOrder = order;
            }
        });

        const newUniqueID = Utilities.getUuid();
        const newItemData = [];
        newItemData[eventIDColIdx_backend] = currentEventID;
        newItemData[backendHeaders.indexOf("UniqueID")] = newUniqueID;
        newItemData[backendHeaders.indexOf("Sequence")] = rowValues[0]; // Col B
        newItemData[backendHeaders.indexOf("Topic/Action")] = rowValues[1]; // Col C
        newItemData[backendHeaders.indexOf("Details/Lines")] = rowValues[2]; // Col D
        newItemData[backendHeaders.indexOf("Actor/Owner")] = rowValues[3]; // Col E
        newItemData[originalOrderColIdx_backend] = maxOrder + 1;

        // Ensure newItemData is a flat array in the correct column order for appendRow
        const finalNewRow = backendHeaders.map((_, idx) => newItemData[idx] !== undefined ? newItemData[idx] : "");

        backendSheet.appendRow(finalNewRow);
        sh.getRange(editedRow, HELPER_UNIQUE_ID_COL).setValue(newUniqueID);
        Logger.log(`Added new item to backend. UniqueID: ${newUniqueID}, Sheet Row: ${editedRow}, Data: ${finalNewRow.join(', ')}`);

        const actorValue = rowValues[3]; // Actor from Col E (index 3 of rowValues)
        if (actorValue) {
            const currentEventActors = new Set();
            backendData.forEach(row => { // Add existing actors for the event
                if (row[eventIDColIdx_backend] === currentEventID && row[actorColIdx_backend]) {
                    currentEventActors.add(row[actorColIdx_backend]);
                }
            });
            currentEventActors.add(actorValue); // Add the new actor

            const sortedActors = Array.from(currentEventActors).sort().map(actor => [actor]);
            const lastActorRowG = sh.getRange("G:G").getValues().filter(String).length;
            if (lastActorRowG >= SCHED_START_ROW) {
                 sh.getRange(SCHED_START_ROW, 7, lastActorRowG - SCHED_START_ROW + 1, 1).clearContent();
            }
            if (sortedActors.length > 0) {
              sh.getRange(SCHED_START_ROW, 7, sortedActors.length, 1).setValues(sortedActors);
            }
            Logger.log(`Actor list updated for new item. Recalculating schedule.`);
        }
        recalculateScheduleTimingsForSheet(sh);
      }
    }
    return;
  }

  // Scenario 3: Edit in Schedule Table Columns (G-J) for timing/actor assignment
  if (editedRow >= SCHED_START_ROW && (editedCol >= 7 && editedCol <= 10)) {
    Logger.log(`Timing-related cell edit detected at ${editedCell.getA1Notation()} in schedule table.`);
    recalculateScheduleTimingsForSheet(sh);
    Logger.log('Schedule recalculated due to edit in G, H, I, or J.');
    return;
  }

  Logger.log(`onEdit in ${TAB_DEST} completed without specific action for cell ${editedCell.getA1Notation()}.`);
}

/**
 * Recalculates and updates schedule timings in column J of the given sheet.
 * Reads session start time from E1, and task details from G, H, I.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetObject The sheet to process.
 */
function recalculateScheduleTimingsForSheet(sheetObject) {
  const sh = sheetObject;
  Logger.log(`Recalculating schedule timings for sheet: ${sh.getName()}`);

  const sessionVal = sh.getRange('E1').getValue();
  const sm = sessionVal.match(/–\s.+?–\s+(\d{1,2}:\d{2}\s*[AP]M)/i);
  if (!sm || sm.length < 2) {
    Logger.log(`Could not parse session start time from E1 value: '${sessionVal}'. Existing times in J might be stale.`);
    // Optional: Clear column J if time is invalid, but consider user-entered data.
    // sh.getRange(SCHED_START_ROW, 10, sh.getMaxRows() - SCHED_START_ROW + 1, 1).clearContent();
    return;
  }
  Logger.log(`Parsed session time string from E1: ${sm[1]}`);

  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'MM/dd/yyyy ');
  const toDate = tStr => new Date(todayStr + tStr.replace(/\s+/g, ' ')); // Normalize spaces in time string

  const sessionStart = toDate(sm[1]);
  if (isNaN(sessionStart.getTime())) {
    Logger.log(`Parsed session start time '${sm[1]}' is not a valid date. Existing times in J might be stale.`);
    return;
  }
  Logger.log(`Session start time successfully parsed: ${sessionStart}`);

  const lastRowData = sh.getLastRow();
  let schedRangeValues;
  if (lastRowData >= SCHED_START_ROW) {
    // Read G:J (Actor, Order, Duration, Manual/Calculated Time)
    schedRangeValues = sh.getRange(SCHED_START_ROW, 7,
                                   lastRowData - SCHED_START_ROW + 1, 4).getValues();
  } else {
    Logger.log("No data in schedule range (G${SCHED_START_ROW} onwards) to process.");
    // Ensure J is cleared if there's no data
    const maxRows = sh.getMaxRows();
    if (maxRows >= SCHED_START_ROW) {
        sh.getRange(SCHED_START_ROW, 10, maxRows - SCHED_START_ROW + 1, 1).clearContent();
    }
    return;
  }

  // Prepare an array to hold the new values for column J. Initialize with current values.
  let newJColumnValues = schedRangeValues.map(row => [row[3]]); // row[3] is current column J value

  const tasks = schedRangeValues.map((r, i) => ({
    idx: i, // Original index in schedRangeValues, corresponds to row number - SCHED_START_ROW
    actor: r[0],  // Column G
    order: Number(r[1]), // Column H
    dur: r[2],    // Column I
    manualRange: String(r[3] || "") // Column J
  }))
  .filter(t => t.actor && typeof t.order === 'number' && !isNaN(t.order) && t.order > 0) // Must have actor and valid positive order
  .sort((a, b) => a.order - b.order);

  if (!tasks.length) {
    Logger.log("No valid tasks found to schedule after filtering by actor and order.");
    // Clear column J for all rows in the schedRangeValues if no tasks are valid
    if (schedRangeValues.length > 0) {
       const clearedJValues = schedRangeValues.map(() => [""]);
       sh.getRange(SCHED_START_ROW, 10, clearedJValues.length, 1).setValues(clearedJValues);
       Logger.log("Cleared column J as no valid tasks were found.");
    }
    return;
  }
  Logger.log(`Processing ${tasks.length} tasks for schedule timing.`);

  const toMin = v => {
    if (v instanceof Date) return v.getHours() * 60 + v.getMinutes();
    if (!isNaN(Number(v)) && Number(v) > 0) return Number(v); // "30" (duration in minutes)
    let totalMinutes = 0;
    const timePattern = String(v).match(/(\d+)\s*:\s*(\d+)/); // "1:15"
    if (timePattern) {
        totalMinutes = (parseInt(timePattern[1]) * 60) + parseInt(timePattern[2]);
    } else {
        const hMatch = String(v).match(/(\d+(?:\.\d+)?)\s*h/i); // "1h" or "1.5h"
        const mMatch = String(v).match(/(\d+)\s*m/i); // "30m"
        if (hMatch) totalMinutes += parseFloat(hMatch[1]) * 60;
        if (mMatch) totalMinutes += parseInt(mMatch[1]);
    }
    return totalMinutes > 0 ? Math.round(totalMinutes) : 0; // Return rounded total minutes, or 0 if invalid/zero
  };

  let cursor = new Date(sessionStart.getTime());

  // First pass: Apply all user-defined times and sort tasks by these times if they exist
  tasks.forEach(t => {
    const m = t.manualRange.match(/(\d{1,2}:\d{2}\s*[AP]M)\s*-\s*(\d{1,2}:\d{2}\s*[AP]M)/i);
    if (m && m.length === 3) {
      const userStart = toDate(m[1]);
      const userEnd = toDate(m[2]);
      if (!isNaN(userStart.getTime()) && !isNaN(userEnd.getTime()) && userEnd > userStart) {
        t.calculatedStart = userStart; // Store these for sorting and cursor advancement
        t.calculatedEnd = userEnd;
        Logger.log(`Task for actor '${t.actor}' (Order ${t.order}): Using user-defined range ${Utilities.formatDate(userStart, tz, 'h:mm a')} - ${Utilities.formatDate(userEnd, tz, 'h:mm a')}`);
      } else {
        Logger.log(`Task for actor '${t.actor}' (Order ${t.order}): Invalid user-defined range '${t.manualRange}'. Will attempt to calculate.`);
        t.manualRange = ""; // Invalidate for calculation purposes
      }
    }
  });

  // Re-sort tasks: those with fixed times first, then by order.
  // This is complex if fixed times interleave with ordered items.
  // Simpler: honor fixed times and let them dictate cursor, then fill gaps.
  // For now, the provided logic iterates through sorted-by-order tasks.
  // If a task has a valid user-defined time, it uses it and advances the main cursor.

  tasks.forEach(t => {
    let start, end;

    if (t.calculatedStart && t.calculatedEnd) { // User-defined time is valid
        start = t.calculatedStart;
        end = t.calculatedEnd;
        // If this task's user-defined start is earlier than current cursor, it might indicate an overlap or out-of-order fixed time.
        // For simplicity, we let user-defined times override the cursor.
        cursor = new Date(end.getTime());
    } else { // Calculate time
        const mins = toMin(t.dur);
        if (!mins) {
          Logger.log(`Task for actor '${t.actor}' (Order ${t.order}): Invalid or zero duration '${t.dur}'. Clearing time for this task.`);
          newJColumnValues[t.idx] = [""]; // Clear time in J
          return; // Skip to next task
        }
        start = new Date(cursor.getTime());
        end = new Date(start.getTime() + mins * 60000);
        cursor = new Date(end.getTime());
    }

    const startStr = Utilities.formatDate(start, tz, 'h:mm a');
    const endStr = Utilities.formatDate(end, tz, 'h:mm a');
    newJColumnValues[t.idx] = [`${startStr}-${endStr}`];
    Logger.log(`Task for actor '${t.actor}' (Order ${t.order}): Set range ${startStr}-${endStr}`);
  });

  // Update all relevant J column cells in one go
  // This ensures that even if tasks are sparse, only corresponding J cells are updated.
  if (newJColumnValues.length > 0) {
      // This needs to write to the correct rows in the sheet, not just a contiguous block
      // The `tasks` array is sorted and filtered. `newJColumnValues` is dense.
      // We need to map results from `tasks` back to their original positions in `schedRangeValues`

      // Create a temporary array matching the original `schedRangeValues` length, filled with blanks or original values
      let finalJOutputValues = schedRangeValues.map(row => [row[3]]); // Default to existing values

      tasks.forEach(task => {
        // task.idx is the original index within schedRangeValues
        // newJColumnValues contains the computed results for `tasks` list, in the same sorted order.
        // This direct mapping is wrong. The result for tasks[i] should go to newJColumnValues[tasks[i].idx]
        // The current newJColumnValues is already built with the correct indices in mind before filtering/sorting for `tasks`
        // The update to newJColumnValues[t.idx] inside the loop correctly places results.
      });

      // So, newJColumnValues should be correctly mapped if t.idx was used properly.
      sh.getRange(SCHED_START_ROW, 10, newJColumnValues.length, 1).setValues(newJColumnValues);
      Logger.log(`Batch updated column J with ${newJColumnValues.length} time calculations based on current data.`);
  } else {
    Logger.log("No values to update in column J.");
  }
}

Logger.log("schedule_onedit.gs processing: New onEdit logic and recalculateScheduleTimingsForSheet defined. Main onEdit is set.");
// End of script
