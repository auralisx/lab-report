// Global variable to store the last processed row
var LAST_PROCESSED_ROW_KEY = "lastProcessedRow";

function getLastProcessedRow() {
  var props = PropertiesService.getScriptProperties();
  var lastProcessedRow = props.getProperty(LAST_PROCESSED_ROW_KEY);
  return lastProcessedRow ? parseInt(lastProcessedRow) : 1; // Start from row 2 if it's the first run
}

function setLastProcessedRow(rowNumber) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(LAST_PROCESSED_ROW_KEY, rowNumber.toString());
}

function checkForNewSubmissions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastProcessedRow = getLastProcessedRow();

  if (lastRow != lastProcessedRow) {
    Logger.log("Form Submission Detected: Proceed pdf generation");
    // New rows have been added so proceed pdf generation
    createDynamicLabReport();

    // Update the last processed row
    setLastProcessedRow(lastRow);
  } else {
    Logger.log("No New Appended Row detected");
  }
}

function onSheetChange(e) {
  const lock = LockService.getScriptLock();

  // Lock the script
  lock.waitLock(100000); // wait 100 seconds for others' use of the code section and lock to stop and then proceed

  try {
    const changeType = e.changeType;
    if (changeType === "OTHER") {
      Logger.log("Other Change: " + changeType);
    } else {
      // Logger.log('Form Submission Detected: Proceed pdf generation');
      checkForNewSubmissions();
    }
  } catch (error) {
    Logger.log("Error in onSheetChange: " + error.toString());
  } finally {
    // Release the lock
    lock.releaseLock();
  }
}

function setupTrigger() {
  // Delete any existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  // Create a new installable onChange trigger
  ScriptApp.newTrigger("onSheetChange")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  Logger.log("Trigger set up successfully");
}
