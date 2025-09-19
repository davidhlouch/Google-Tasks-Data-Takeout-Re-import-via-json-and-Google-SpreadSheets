/**
 * This script provides a unified utility for managing tasks between Google Sheets and Google Tasks.
 * It includes functions to:
 * - Import tasks from a JSON file into a Google Sheet.
 * - Export tasks from a Google Sheet to a Google Tasks list.
 *
 * It uses advanced techniques like batch processing with time-based triggers and a persistent cache
 * to handle large datasets and API propagation delays without hitting quotas or execution time limits.
 *
 * @author Gemini (with optimizations)
 */

// Define constants for the script's properties and triggers.
const SCRIPT_PROPERTY_ROW = 'lastProcessedRow';
const SCRIPT_PROPERTY_SHEET = 'lastProcessedSheetIndex';
const SCRIPT_PROPERTY_CACHE = 'taskListCache';
const SCRIPT_PROPERTY_LIST = 'lastKnownListTitle'; // Remembers the list title across batches
const SCRIPT_TRIGGER_HANDLER = 'continueProcess';
const API_DELAY_MS = 500; // Delay in milliseconds between each API call to prevent quota errors.
const BATCH_SIZE = 100; // Process 100 rows at a time.

/**
 * Creates a custom menu in the Google Sheet for running the script.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Data Utilities')
    .addItem('Import JSON (from Prompt)', 'importJsonFromPrompt')
    .addItem('Create Tasks from All Sheets', 'createTasksFromAllSheets')
    .addToUi();
}

/**
 * Initiates the JSON import process by prompting the user for data.
 */
function importJsonFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const promptResponse = ui.prompt(
    'Import JSON Data',
    'Paste the JSON data here:',
    ui.ButtonSet.OK_CANCEL
  );

  if (promptResponse.getSelectedButton() === ui.Button.OK) {
    const rawJson = promptResponse.getResponseText();
    if (rawJson) {
      importJsonToSheet(rawJson);
    } else {
      ui.alert('No JSON data was entered. Please try again.');
    }
  }
}

/**
 * Main function to import JSON data into the active spreadsheet.
 * @param {string} rawJson The raw JSON string to be parsed.
 */
function importJsonToSheet(rawJson) {
  const ui = SpreadsheetApp.getUi();
  try {
    const data = JSON.parse(rawJson);
    if (!Array.isArray(data) || data.length === 0) {
      ui.alert('Invalid JSON data. Please ensure it is a non-empty array of objects.');
      return;
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.clear(); // Clear existing content before importing.

    const allHeaders = new Set();
    data.forEach(item => {
      Object.keys(item).forEach(key => allHeaders.add(key));
      if (Array.isArray(item.items)) {
        item.items.forEach(task => {
          Object.keys(task).forEach(key => allHeaders.add(key));
        });
      }
    });

    const orderedHeaders = [
      'list_title', 'list_id', 'id', 'title', 'status', 'created', 'updated', 'due', 'links', 'task_type', 'kind', 'selfLink'
    ];
    const finalHeaders = orderedHeaders.filter(header => allHeaders.has(header));

    const outputData = [finalHeaders];

    data.forEach(taskList => {
      const listTitle = taskList.title;
      const listId = taskList.id;

      if (Array.isArray(taskList.items) && taskList.items.length > 0) {
        taskList.items.forEach(task => {
          const row = finalHeaders.map(header => {
            if (header === 'list_title') return listTitle;
            if (header === 'list_id') return listId;
            return task[header] !== undefined ? task[header] : '';
          });
          outputData.push(row);
        });
      } else {
        const emptyRow = finalHeaders.map(header => {
          if (header === 'list_title') return listTitle;
          if (header === 'list_id') return listId;
          return '';
        });
        outputData.push(emptyRow);
      }
    });

    sheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
    ui.alert(`Import complete! Successfully imported ${outputData.length - 1} records to the sheet.`);

  } catch (e) {
    ui.alert(`An error occurred: ${e.message}`);
    Logger.log(e);
  }
}

/**
 * Main function to initiate the task creation process for all sheets.
 */
function createTasksFromAllSheets() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(SCRIPT_PROPERTY_ROW);
  properties.deleteProperty(SCRIPT_PROPERTY_SHEET);
  properties.deleteProperty(SCRIPT_PROPERTY_CACHE);
  properties.deleteProperty(SCRIPT_PROPERTY_LIST); // Clear the last known list title
  deleteExistingTriggers();

  SpreadsheetApp.getUi().alert('Starting task creation. This may take a while for large sheets. The script will run in batches to avoid timeouts.');
  continueProcess();
}

/**
 * The core function that processes all sheets in batches.
 * This version "remembers" the last seen list title and uses the 'id' as a fallback for an empty 'title'.
 */
function continueProcess() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = spreadsheet.getSheets();

  const properties = PropertiesService.getScriptProperties();
  let currentSheetIndex = parseInt(properties.getProperty(SCRIPT_PROPERTY_SHEET), 10) || 0;
  let startRow = parseInt(properties.getProperty(SCRIPT_PROPERTY_ROW), 10) || 2;
  
  if (currentSheetIndex >= allSheets.length) {
    Logger.log('All tasks have been successfully created! Cleaning up triggers.');
    properties.deleteProperty(SCRIPT_PROPERTY_ROW);
    properties.deleteProperty(SCRIPT_PROPERTY_SHEET);
    properties.deleteProperty(SCRIPT_PROPERTY_CACHE);
    properties.deleteProperty(SCRIPT_PROPERTY_LIST);
    deleteExistingTriggers();
    return;
  }
  
  const sheet = allSheets[currentSheetIndex];
  const lastRow = sheet.getLastRow();
  
  if (startRow > lastRow) {
    Logger.log(`Processing complete for sheet "${sheet.getName()}". Moving to the next sheet.`);
    properties.setProperty(SCRIPT_PROPERTY_SHEET, (currentSheetIndex + 1).toString());
    properties.deleteProperty(SCRIPT_PROPERTY_ROW);
    properties.deleteProperty(SCRIPT_PROPERTY_LIST); // Reset list title for new sheet
    createTrigger();
    return;
  }

  Logger.log(`Starting process. Sheet: "${sheet.getName()}", Starting row: ${startRow}`);
  
  let cacheString = properties.getProperty(SCRIPT_PROPERTY_CACHE);
  let taskListCache = cacheString ? JSON.parse(cacheString) : {};
  let lastKnownListTitle = properties.getProperty(SCRIPT_PROPERTY_LIST) || null;

  try {
    const taskLists = Tasks.Tasklists.list().items;
    if (taskLists) {
      taskLists.forEach(list => {
        taskListCache[list.title] = list.id;
      });
    }
  } catch (e) {
    Logger.log(`Could not retrieve task lists to update cache: ${e.message}. Proceeding with stored cache.`);
  }

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const titleIndex = headers.indexOf('title');
  const idIndex = headers.indexOf('id'); // Get the index for the ID column
  const listTitleIndex = headers.indexOf('list_title');
  const statusIndex = headers.indexOf('status');
  const dueIndex = headers.indexOf('due');
  const linksIndex = headers.indexOf('links');

  if (titleIndex === -1 || listTitleIndex === -1) {
    Logger.log(`Skipping sheet "${sheet.getName()}". Required headers 'title' or 'list_title' not found.`);
    properties.setProperty(SCRIPT_PROPERTY_SHEET, (currentSheetIndex + 1).toString());
    properties.deleteProperty(SCRIPT_PROPERTY_ROW);
    createTrigger();
    return;
  }

  const rowsToProcess = Math.min(BATCH_SIZE, lastRow - startRow + 1);
  const data = sheet.getRange(startRow, 1, rowsToProcess, lastColumn).getValues();

  for (let j = 0; j < data.length; j++) {
    const currentRowNumber = startRow + j;
    const row = data[j];
    
    // Remember the last seen list title
    const listTitleFromCell = row[listTitleIndex] ? row[listTitleIndex].toString().trim() : null;
    if (listTitleFromCell) {
      lastKnownListTitle = listTitleFromCell;
    }
    const taskListName = lastKnownListTitle;
    
    // *** MODIFIED LOGIC: Use ID as fallback for Title ***
    const taskTitleFromCell = row[titleIndex] ? row[titleIndex].toString().trim() : null;
    const taskIdFromCell = idIndex !== -1 && row[idIndex] ? row[idIndex].toString().trim() : null;
    const finalTaskTitle = taskTitleFromCell || taskIdFromCell; // Prioritize title, fallback to ID
    // *** END OF MODIFIED LOGIC ***

    if (!taskListName) {
      Logger.log(`Row ${currentRowNumber}: Skipped because no 'list_title' has been found yet in this sheet.`);
      continue;
    }
    if (!finalTaskTitle) {
      Logger.log(`Row ${currentRowNumber}: Skipped because both 'title' and 'id' are empty.`);
      continue;
    }
    if (finalTaskTitle.toLowerCase().startsWith('list:')) {
        Logger.log(`Row ${currentRowNumber}: Skipped because it appears to be a list definition row.`);
        continue;
    }

    const currentTaskListId = getTaskListIdByName(taskListName, taskListCache);

    if (currentTaskListId) {
      const task = Tasks.newTask();
      task.title = finalTaskTitle; // Use the determined title

      if (statusIndex !== -1 && row[statusIndex] && row[statusIndex].toString().toLowerCase().trim() === 'completed') {
        task.status = 'completed';
      }

      if (dueIndex !== -1 && row[dueIndex]) {
        try {
          const dueDate = new Date(row[dueIndex]);
          if (!isNaN(dueDate.getTime())) {
            task.due = dueDate.toISOString();
          }
        } catch (e) { /* Ignore invalid dates */ }
      }

      if (linksIndex !== -1 && row[linksIndex]) {
        task.notes = `Link: ${row[linksIndex]}`;
      }

      try {
        Tasks.Tasks.insert(task, currentTaskListId);
        Logger.log(`Row ${currentRowNumber}: Successfully created task "${finalTaskTitle}" in list "${taskListName}".`);
        Utilities.sleep(API_DELAY_MS);
      } catch (e) {
        Logger.log(`Row ${currentRowNumber}: FAILED to create task "${finalTaskTitle}". Error: ${e.message}`);
      }
    } else {
      Logger.log(`Row ${currentRowNumber}: Skipped because an ID could not be found or created for task list "${taskListName}".`);
    }
  }

  const nextStartRow = startRow + rowsToProcess;
  properties.setProperty(SCRIPT_PROPERTY_ROW, nextStartRow.toString());
  properties.setProperty(SCRIPT_PROPERTY_SHEET, currentSheetIndex.toString());
  properties.setProperty(SCRIPT_PROPERTY_CACHE, JSON.stringify(taskListCache));
  if(lastKnownListTitle) {
      properties.setProperty(SCRIPT_PROPERTY_LIST, lastKnownListTitle); // Save the last title for the next batch
  }
  createTrigger();
  Logger.log(`Processing complete for batch. Next batch will start at row ${nextStartRow}.`);
}

/**
 * Helper function to find or create a task list by name using a persistent cache.
 * This version includes a fallback mechanism to handle API race conditions.
 * @param {string} name The name of the task list to find or create.
 * @param {Object<string, string>} cache A map of task list titles to their IDs, passed by reference.
 * @return {string | null} The ID of the task list, or null if creation fails.
 */
function getTaskListIdByName(name, cache) {
  if (cache[name]) {
    return cache[name];
  }

  try {
    const newTaskList = Tasks.newTaskList();
    newTaskList.title = name;
    const createdList = Tasks.Tasklists.insert(newTaskList);
    
    if (createdList && createdList.id) {
        cache[name] = createdList.id; 
        Logger.log(`Created new task list and added to cache: "${name}"`);
        return createdList.id;
    }
    Logger.log(`Failed to create task list "${name}" - API did not return an ID.`);
    return null;
  } catch (e) {
    Logger.log(`Could not create list "${name}" (it may already exist). Error: ${e.message}. Attempting to find it by re-fetching all lists.`);
    
    Utilities.sleep(2000); 

    try {
        const taskLists = Tasks.Tasklists.list().items;
        if (taskLists) {
            for (const list of taskLists) {
                if (list.title === name) {
                    Logger.log(`Found existing list "${name}" after re-fetching.`);
                    cache[name] = list.id;
                    return list.id;
                }
            }
        }
        Logger.log(`Still could not find task list "${name}" after re-fetching. No tasks will be created for this list in this batch.`);
        return null;
    } catch (e2) {
        Logger.log(`Failed to re-fetch task lists. Error: ${e2.message}`);
        return null;
    }
  }
}

/**
 * Helper function to create a new time-based trigger.
 */
function createTrigger() {
  deleteExistingTriggers();
  ScriptApp.newTrigger(SCRIPT_TRIGGER_HANDLER)
    .timeBased()
    .at(new Date(Date.now() + 60 * 1000)) // 1 minute from now
    .create();
}

/**
 * Helper function to delete any existing triggers for this script.
 */
function deleteExistingTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === SCRIPT_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

