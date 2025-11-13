/** Utility Functions used in main.gs */

// Get the current date
const currentDate = new Date();
const formattedDate = Utilities.formatDate(
  currentDate,
  Session.getScriptTimeZone(),
  "MM/dd/yyyy hh:mm a"
);

// Save the Doc file and return PDF
function saveAsPdf(docId, fileName) {
  // Logger.log(docId);
  var doc = DocumentApp.openById(docId);
  var pdfFile = DriveApp.createFile(doc.getAs("application/pdf"));
  pdfFile.setName("Lab_Report_" + fileName);

  // DriveApp.getFileById(docId).setTrashed(true);  // Clean up temporary Google Docs file
  return pdfFile;
}

// To check for empty test value
function checkIfEmpty(value) {
  return (
    value == null ||
    value === "" ||
    (typeof value === "string" && value.trim() === "") ||
    (Array.isArray(value) && value.length === 0) ||
    (typeof value === "object" && Object.keys(value).length === 0)
  );
}

// create lab number
function generateLabNumber() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const formDataSheet = sheet.getSheetByName("Data");
  const lastRow = formDataSheet.getLastRow();

  // Get the current year and month
  var date = new Date();
  var year = date.getFullYear().toString().slice(-2); // Get last two digits of the year
  var month = ("0" + (date.getMonth() + 1)).slice(-2); // Get the month and ensure it's two digits

  var serialNum = lastRow - 1; //exclude the header row

  // Format the serial number to be three digits, starting from 001
  var formattedSerialNumber = ("000" + serialNum).slice(-3);

  // Generate the full report number
  var labNumber = year + month + formattedSerialNumber;

  // Log and return the report number
  // Logger.log(reportNumber);
  return labNumber;
}

// Function to check if there’s enough space for a table on the current page
function hasEnoughSpace(body, estimatedHeight) {
  const contentHeight = body.getText().length * 0.7; // Estimate current content height
  const remainingHeight = PAGE_HEIGHT - (contentHeight % PAGE_HEIGHT); // Approximation
  return remainingHeight >= estimatedHeight;
}

// Function to calculate table height based on row count
function calculateTableHeight(rowCount) {
  return rowCount * ROW_HEIGHT;
}
