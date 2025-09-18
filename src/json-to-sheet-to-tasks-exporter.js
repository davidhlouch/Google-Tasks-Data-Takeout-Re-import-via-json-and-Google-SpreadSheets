/**
 * This script provides a simple utility to manage data between a Google Sheet
 * and Google Tasks. It imports JSON data into a sheet and exports tasks
 * from the active sheet to Google Tasks.
 *
 * @author Gemini
 */

/**
 * Creates a custom menu in the Google Sheet for running the script.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Data Utilities');
  
  // Add only the two requested menu items.
  menu.addItem('Create Tasks from Sheet', 'createTasksFromSheet');
  menu.addSeparator();
  menu.addItem('Import JSON (from Prompt)', 'importJsonToSheet');
  menu.addToUi();
}

/**
 * Main function to initiate the task creation process for the active sheet.
 */
function createTasksFromSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(`Starting task creation for sheet "${activeSheet.getName()}".`);
  
  // Get the headers to dynamically find the correct columns.
  const headers = activeSheet.getRange(1, 1, 1, activeSheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header.toString().trim().toLowerCase()] = index + 1;
  });

  // Dynamically get the column numbers based on header names.
  const getCol = (headerName) => {
    const col = headerMap[headerName.toLowerCase()];
    if (!col) {
      Logger.log(`Error: Missing required header column "${headerName}".`);
      SpreadsheetApp.getUi().alert(`Error: Missing required header column "${headerName}". Please check your sheet.`);
      return null;
    }
    return col;
  };

  const taskTitleColumn = getCol('title');
  const taskDueDateColumn = getCol('due');
  const taskStatusColumn = getCol('status');
  const taskStarredColumn = getCol('links'); // The JSON export puts starred status in the links field.
  const taskLinkColumn = getCol('links');
  const taskListNameColumn = getCol('list_title'); // The JSON import puts the task list name in the list_title field.

  if (!taskTitleColumn || !taskListNameColumn) {
    return;
  }

  const lastRow = activeSheet.getLastRow();
  const data = activeSheet.getRange(2, 1, lastRow - 1, activeSheet.getLastColumn()).getValues();

  let currentTaskList = null;

  // Loop through each row of data.
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    // Check if the row represents a task list.
    const rowTitle = row[taskTitleColumn - 1];
    if (typeof rowTitle === 'string' && rowTitle.trim().toLowerCase().startsWith("list")) {
      const listName = row[taskListNameColumn - 1];
      if (listName) {
        currentTaskList = getTaskListByName(listName.toString().trim());
      }
      continue; // Skip to the next row as this is a list row, not a task.
    }

    const taskTitle = row[taskTitleColumn - 1];
    if (taskTitle && currentTaskList) { // Only create a task if the title is not empty and a task list is defined.
      const task = Tasks.newTask();
      
      // Ensure the task title is a string.
      const taskStarred = row[getCol('links') - 1];
      if (taskStarred && typeof taskStarred === 'string' && taskStarred.toString().toLowerCase().includes('starred')) {
        task.title = `⭐ ${taskTitle.toString()}`;
      } else {
        task.title = taskTitle.toString();
      }
      
      // Add a check for a valid date before setting the 'due' property.
      const taskDueDate = row[getCol('due') - 1];
      if (taskDueDate instanceof Date && !isNaN(taskDueDate.getTime())) {
        task.due = taskDueDate.toISOString();
      } else {
        try {
          const parsedDate = new Date(taskDueDate);
          if (!isNaN(parsedDate.getTime())) {
            task.due = parsedDate.toISOString();
          }
        } catch (e) {
          Logger.log(`Skipping due date for "${taskTitle}" due to invalid value: "${taskDueDate}". Error: ${e.message}`);
        }
      }

      // Add the link to the task's notes (description) if it exists.
      const taskLink = row[getCol('links') - 1];
      if (typeof taskLink === 'string' && taskLink.trim() !== '' && taskLink.toString().toLowerCase() !== 'starred') {
        task.notes = `Link: ${taskLink.trim()}`;
      }

      // Check if the status is 'completed' and set the task status.
      const taskStatus = row[getCol('status') - 1];
      if (taskStatus && taskStatus.toString().toLowerCase().trim() === 'completed') {
        task.status = 'completed';
      }
      
      Tasks.Tasks.insert(task, currentTaskList.id);
      
      // Add a small delay to avoid hitting the API rate limit.
      Utilities.sleep(500);
    }
  }

  ui.alert('All tasks have been successfully created!');
}

/**
 * Prompts the user for a JSON string and imports it into the active sheet.
 */
function importJsonToSheet() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Import JSON', 'Paste your JSON data below:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) { return; }

  let data;
  try {
    data = JSON.parse(response.getResponseText());
  } catch (e) {
    ui.alert('Error', 'Invalid JSON data. Please check the format and try again.', ui.ButtonSet.OK);
    return;
  }
  
  importDataToSheet(data, 'JSON data has been successfully imported.');
}

/**
 * A helper function to handle the core logic of importing data to the sheet.
 * @param {Array} data The data array to import.
 * @param {string} successMessage The message to display on success.
 */
function importDataToSheet(data, successMessage) {
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!Array.isArray(data) || data.length === 0) {
    SpreadsheetApp.getUi().alert('Invalid data format. Please ensure it is a non-empty array of objects.');
    return;
  }
  
  targetSheet.clear();
  const headers = Object.keys(data[0]);
  const outputData = [headers];
  
  data.forEach(item => {
    const row = headers.map(header => {
      const value = item[header];
      if (typeof value === 'object' && value !== null) {
        if (value.url) return value.url;
        if (value.href) return value.href;
        if (value.link) return value.link;
        return JSON.stringify(value);
      }
      return value;
    });
    outputData.push(row);
  });
  
  const range = targetSheet.getRange(1, 1, outputData.length, outputData[0].length);
  range.setValues(outputData);
  targetSheet.autoResizeColumns(1, targetSheet.getLastColumn());
  SpreadsheetApp.getUi().alert(successMessage);
}

/**
 * Helper function to find or create a task list by name.
 * @param {string} name The name of the task list to find or create.
 * @return {Object} The task list object.
 */
function getTaskListByName(name) {
  let taskList = null;
  const taskLists = Tasks.Tasklists.list().items;
  
  if (taskLists) {
    for (let i = 0; i < taskLists.length; i++) {
      if (taskLists[i].title === name) {
        taskList = taskLists[i];
        break;
      }
    }
  }
  
  if (!taskList) {
    try {
      const newTaskList = Tasks.newTaskList();
      newTaskList.title = name;
      taskList = Tasks.Tasklists.insert(newTaskList);
    } catch (e) {
      Logger.log(`Failed to create task list: ${e.message}`);
    }
  }
  
  return taskList;
}
