// IDs
const TEMPLATE_DOC_ID = "id for the template document";
const SIGNATURE_TEMPLATE_DOC_ID = "id for the template document";
const REPORT_FOLDER_ID = "id for the template document";
const REPORT_TEMPLATE_ID = "id for the template document";
const REPORT_TEMPLATE_WITHOUT_HF_ID = "id for the template document";
const COMMENT_TEMPLATE_ID = "id for the template document";

// Estimated page height of the doc and row height of the table
const ROW_HEIGHT = 17;
const PAGE_HEIGHT = 792;

// Custom menu creation
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Custom Menu")
    .addItem("Generate Dynamic Report", "createDynamicLabReport")
    .addToUi();
}

function createReportWithHeaderAndFooter(formData, patientFolder) {
  const withHF = patientFolder.createFolder("Report_With_Header_&_Footer");

  // Create a copy of report template doc to edit
  const docCopyId = DriveApp.getFileById(REPORT_TEMPLATE_ID).makeCopy().getId();
  // Logger.log(docCopyId);
  DriveApp.getFileById(docCopyId).moveTo(withHF);
  const doc = DocumentApp.openById(docCopyId);
  doc.setName("Lab_Report_" + formData[0]);

  const body = doc.getBody();

  replacePlaceholdersInHeader(doc, formData); // Replace placeholders in header
  processTableInsertion(body, formData); // Insert tables

  // appendSignatureTable(doc, 0, formData[9]);

  doc.saveAndClose();
  // Utilities.sleep(1000); // wait for 1s for document generation

  const pdfFile = saveAsPdf(docCopyId, formData[0]);
  Logger.log("PDF with header and footer created: " + pdfFile.getUrl());
  const pdfFileId = pdfFile.getId();
  DriveApp.getFileById(pdfFileId).moveTo(withHF);

  return {
    pdf: pdfFile,
    docURL: doc.getUrl(),
    folder: withHF,
  };
}

function createReportWithoutHeaderAndFooter(formData, patientFolder) {
  const withoutHF = patientFolder.createFolder(
    "Report_Without_Header_&_Footer"
  );

  // Create a copy of report template doc to edit
  const docCopyId = DriveApp.getFileById(REPORT_TEMPLATE_WITHOUT_HF_ID)
    .makeCopy()
    .getId();
  // Logger.log(docCopyId);
  DriveApp.getFileById(docCopyId).moveTo(withoutHF);
  const doc = DocumentApp.openById(docCopyId);
  doc.setName("Lab_Report_" + formData[0]);

  const body = doc.getBody();

  replacePlaceholdersInHeader(doc, formData); // Replace placeholders in header
  processTableInsertion(body, formData); // Insert tables

  // appendSignatureTable(doc, 0, formData[9]);

  doc.saveAndClose();
  // Utilities.sleep(1000); // wait for 1s for document generation

  const pdfFile = saveAsPdf(docCopyId, formData[0]);
  Logger.log("PDF without header and footer created: " + pdfFile.getUrl());
  const pdfFileId = pdfFile.getId();
  DriveApp.getFileById(pdfFileId).moveTo(withoutHF);

  return {
    pdf: pdfFile,
    docURl: doc.getUrl(),
    folder: withoutHF,
  };
}

/** Main funtion: Create folder, file and save them */
function createDynamicLabReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const formDataSheet = sheet.getSheetByName("Data");
  const lastRow = formDataSheet.getLastRow();
  const formData = formDataSheet
    .getRange(lastRow, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  try {
    // Create a new folder with Patient Name
    const reportFolder = DriveApp.getFolderById(REPORT_FOLDER_ID);
    const patientFolder = reportFolder.createFolder(formData[0] + "_Report");

    var pdfWithHF = createReportWithHeaderAndFooter(formData, patientFolder);
    var pdfWithoutHF = createReportWithoutHeaderAndFooter(
      formData,
      patientFolder
    );

    // Optionally, email the PDF
    MailApp.sendEmail({
      to: "your gmail address",
      subject: "Lab Report Of Patient: " + formData[0],
      body:
        "Lab Report Are Ready. Please refer to following folder link if needed: \n\nPatient Folder Link: " +
        patientFolder.getUrl() +
        "\n\nYou can find the lab report pdf attached.",
      attachments: [pdfWithHF.pdf, pdfWithoutHF.pdf],
    });
  } catch (error) {
    Logger.log("Error in createDynamicLabReport: " + error.toString());
  }
}
