/**
 * This script provides a unified utility for managing tasks between Google Sheets and Google Tasks.
 * It includes functions to:
 * - Import tasks from a JSON file into a Google Sheet.
 * - Export tasks from the current Google Sheet to a Google Tasks list.
 *
 * It uses advanced techniques like batch processing with time-based triggers and a persistent cache
 * to handle large datasets and API propagation delays without hitting quotas or execution time limits.
 *
 * @author Gemini (with optimizations)
 */

// Define constants for the script's properties and triggers.
const SCRIPT_PROPERTY_ROW = 'lastProcessedRow';
const SCRIPT_PROPERTY_SHEET_NAME = 'sourceSheetName'; // Now stores the name of the single sheet being processed
const SCRIPT_PROPERTY_CACHE = 'taskListCache';
const SCRIPT_PROPERTY_LIST = 'lastKnownListTitle'; // Remembers the list title across batches
const SCRIPT_PROPERTY_REPORT = 'executionReport'; // Stores summary data
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
    .addItem('Create Tasks from Current Sheet', 'createTasksFromCurrentSheet') // Updated menu item
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
 * Main function to initiate the task creation process for the CURRENTLY ACTIVE sheet.
 */
function createTasksFromCurrentSheet() {
  const properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(SCRIPT_PROPERTY_ROW);
  properties.deleteProperty(SCRIPT_PROPERTY_SHEET_NAME);
  properties.deleteProperty(SCRIPT_PROPERTY_CACHE);
  properties.deleteProperty(SCRIPT_PROPERTY_LIST);
  properties.deleteProperty(SCRIPT_PROPERTY_REPORT);
  deleteExistingTriggers();

  // Get the active sheet to process
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = activeSheet.getName();
  properties.setProperty(SCRIPT_PROPERTY_SHEET_NAME, sheetName);

  SpreadsheetApp.getUi().alert(`Starting task creation for sheet "${sheetName}". This may take a while. A new sheet with a summary report will be created upon completion.`);
  continueProcess();
}

/**
 * The core function that processes the selected sheet in batches.
 */
function continueProcess() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const properties = PropertiesService.getScriptProperties();

  const sourceSheetName = properties.getProperty(SCRIPT_PROPERTY_SHEET_NAME);
  if (!sourceSheetName) {
      Logger.log('Error: No source sheet name found in properties. Aborting.');
      deleteExistingTriggers();
      return;
  }
  const sheet = spreadsheet.getSheetByName(sourceSheetName);
  if (!sheet) {
      Logger.log(`Error: Could not find sheet with name "${sourceSheetName}". Aborting.`);
      deleteExistingTriggers();
      return;
  }

  let startRow = parseInt(properties.getProperty(SCRIPT_PROPERTY_ROW), 10) || 2;
  const lastRow = sheet.getLastRow();
  
  const reportDataString = properties.getProperty(SCRIPT_PROPERTY_REPORT);
  let reportData = reportDataString ? JSON.parse(reportDataString) : {};

  // Check if processing for the current sheet is complete
  if (startRow > lastRow) {
    Logger.log('All tasks have been successfully created! Cleaning up triggers and generating report.');

    // --- GENERATE REPORT SHEET (FIXED) ---
    const timestamp = new Date().toLocaleString('sv-SE'); // YYYY-MM-DD HH:MM:SS format
    const reportSheetName = `${sourceSheetName} Report ${timestamp}`;
    
    const reportSheet = spreadsheet.insertSheet(reportSheetName);
    
    // Write header and format
    reportSheet.getRange('A1').setValue('Task Import Summary').setFontWeight('bold').setFontSize(14);
    reportSheet.getRange('A1:D1').merge().setHorizontalAlignment('center');

    reportSheet.getRange('A2').setValue('Generated on:');
    reportSheet.getRange('B2').setValue(new Date());

    // Write table headers
    reportSheet.getRange('A4:D4').setValues([['Task List', 'Total Imported', 'Completed', 'Needs Action']]).setFontWeight('bold');

    // Prepare and write table content
    const tableContent = [];
    const listNames = Object.keys(reportData);
    if (listNames.length > 0) {
        for (const listName of listNames) {
            const stats = reportData[listName];
            tableContent.push([listName, stats.total, stats.completed, stats.needsAction]);
        }
    } else {
        tableContent.push(['No new tasks were imported.', '', '', '']);
    }

    if (tableContent.length > 0) {
        reportSheet.getRange(5, 1, tableContent.length, tableContent[0].length).setValues(tableContent);
    }

    // Auto-resize columns for readability
    reportSheet.autoResizeColumns(1, 4);
    reportSheet.activate(); // Make the report sheet visible to the user

    // Final cleanup of all script properties
    properties.deleteProperty(SCRIPT_PROPERTY_ROW);
    properties.deleteProperty(SCRIPT_PROPERTY_SHEET_NAME);
    properties.deleteProperty(SCRIPT_PROPERTY_CACHE);
    properties.deleteProperty(SCRIPT_PROPERTY_LIST);
    properties.deleteProperty(SCRIPT_PROPERTY_REPORT);
    deleteExistingTriggers();
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
  const idIndex = headers.indexOf('id');
  const listTitleIndex = headers.indexOf('list_title');
  const statusIndex = headers.indexOf('status');
  const dueIndex = headers.indexOf('due');
  const linksIndex = headers.indexOf('links');

  if (titleIndex === -1 || listTitleIndex === -1) {
    Logger.log(`Skipping sheet "${sheet.getName()}". Required headers 'title' or 'list_title' not found.`);
    deleteExistingTriggers(); // Stop if headers are missing
    return;
  }

  const rowsToProcess = Math.min(BATCH_SIZE, lastRow - startRow + 1);
  const data = sheet.getRange(startRow, 1, rowsToProcess, lastColumn).getValues();

  for (let j = 0; j < data.length; j++) {
    const currentRowNumber = startRow + j;
    const row = data[j];
    
    const listTitleFromCell = row[listTitleIndex] ? row[listTitleIndex].toString().trim() : null;
    if (listTitleFromCell) {
      lastKnownListTitle = listTitleFromCell;
    }
    const taskListName = lastKnownListTitle;
    
    const taskTitleFromCell = row[titleIndex] ? row[titleIndex].toString().trim() : null;
    const taskIdFromCell = idIndex !== -1 && row[idIndex] ? row[idIndex].toString().trim() : null;
    const finalTaskTitle = taskTitleFromCell || taskIdFromCell;

    if (!taskListName) {
      Logger.log(`Row ${currentRowNumber}: Skipped because no 'list_title' has been found yet.`);
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
      task.title = finalTaskTitle;

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
        
        // --- UPDATE REPORT DATA ---
        if (!reportData[taskListName]) {
            reportData[taskListName] = { completed: 0, needsAction: 0, total: 0 };
        }
        if (task.status === 'completed') {
            reportData[taskListName].completed++;
        } else {
            reportData[taskListName].needsAction++;
        }
        reportData[taskListName].total++;
        
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
  properties.setProperty(SCRIPT_PROPERTY_CACHE, JSON.stringify(taskListCache));
  properties.setProperty(SCRIPT_PROPERTY_REPORT, JSON.stringify(reportData));
  if(lastKnownListTitle) {
      properties.setProperty(SCRIPT_PROPERTY_LIST, lastKnownListTitle);
  }
  createTrigger();
  Logger.log(`Processing complete for batch. Next batch will start at row ${nextStartRow}.`);
}

/**
 * Helper function to find or create a task list by name using a persistent cache.
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
        Logger.log(`Still could not find task list "${name}" after re-fetching.`);
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

