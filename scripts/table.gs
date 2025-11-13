// Return indexed table from table template doc
function getTableFromTemplate(tableIndex, templateDocId) {
  try {
    const tableTemplateDoc = DocumentApp.openById(templateDocId);
    const tables = tableTemplateDoc.getBody().getTables();

    if (tableIndex < 0 || tableIndex >= tables.length) {
      throw new Error("Invalid table index: " + tableIndex);
    }

    return tables[tableIndex].copy();
  } catch (e) {
    Logger.log("Error in getTableFromTemplate: " + e.message);
    return null;
  }
}

// process table insertion according to the test selected
function processTableInsertion(body, formData) {
  const selectedTests = JSON.parse(formData[10]);
  const labTests = getLabTests(formData);
  const departmentTests = {};

  // Categorize tests by department
  selectedTests.forEach((selectedTest) => {
    const labTest = labTests[selectedTest];
    if (labTest) {
      const department = labTest.department;
      // Logger.log(department);
      if (!departmentTests[department]) {
        departmentTests[department] = [];
      }
      departmentTests[department].push(labTest);
    }
  });

  // Process each department
  const departments = Object.keys(departmentTests);
  let totalTestsProcessed = 0;
  const totalTests = selectedTests.length;

  departments.forEach((department, deptIndex) => {
    if (deptIndex > 0) {
      // Start new department on a new page (except first department)
      body.appendPageBreak();
    }
    const labTestsInDepartment = departmentTests[department];

    // Sort tests - put tests with comments at the end of their department
    labTestsInDepartment.sort((a, b) => {
      if (a.commentTableIndex && !b.commentTableIndex) return 1;
      if (!a.commentTableIndex && b.commentTableIndex) return -1;
      return 0;
    });

    // Process each test in the department
    labTestsInDepartment.forEach((labTestInDepartment, testIndex) => {
      const rowCount = labTestInDepartment.tests.length + 1; // Include header row
      const tableHeight = calculateTableHeight(rowCount);
      const estimatedSpaceNeeded = tableHeight + 200; // 200 extra space for signature

      // Flag to track if we need to add signature after this test
      let needsSignature = false;

      // Check if test has comment
      const hasComment =
        labTestInDepartment.commentTableIndex &&
        (formData[76] === "Yes" || formData[76] === "Custom");

      // If test has a comment, append signature and always start on a new page
      if (hasComment && testIndex > 0) {
        appendSignatureTable(body, 0, formData[9]); // Add signature before page break
        Logger.log(
          "Signature  and pagebreak inserted because the test has comment and is not the first test in the department"
        );
        body.appendPageBreak();
      }

      // Check if there's enough space for the current table
      if (!hasEnoughSpace(body, estimatedSpaceNeeded)) {
        appendSignatureTable(body, 0, formData[9]); // Add signature before page break
        Logger.log(
          "Signature  and pagebreak inserted because there is not enough space for the table"
        );
        body.appendPageBreak();
      }

      // Check if there’s enough space; insert page break and signature if not
      // if (!hasEnoughSpace(body, estimatedSpaceNeeded)) {
      //    Logger.log("Inserted signature due to insufficient space for " + labTestInDepartment.tableHeading);
      //   appendSignatureTable(body, 0, formData[9]);
      //   Logger.log("signature after test inserted before page break");
      //   body.appendPageBreak();
      // } else if (hasEnoughSpace(body, estimatedSpaceNeeded) && labTestInDepartment.commentTableIndex) {
      //   if (currentTestCount>0) {
      //     appendSignatureTable(body, 0, formData[9]);
      //     Logger.log("has comment so, signature inserted before page break");
      //     body.appendPageBreak();
      //   }
      // }

      insertTable(body, labTestInDepartment, formData);
      Logger.log(labTestInDepartment.tableHeading + " table inserted");
      totalTestsProcessed++;

      // Determine if we need to add a signature after this test
      const isLastTestInDept = testIndex === labTestsInDepartment.length - 1;
      const isLastDepartment = deptIndex === departments.length - 1;
      const isLastTest = totalTestsProcessed === totalTests;

      // Add signature if:
      // 1. This test has a comment
      // 2. Or this is the last test in its department
      needsSignature = hasComment || isLastTestInDept;

      if (needsSignature) {
        // Only append page break if this isn't the last test and we have more content coming
        const shouldAddPageBreak =
          !isLastTest && (hasComment || isLastTestInDept);

        appendSignatureTable(body, 0, formData[9], isLastTest);
        Logger.log(
          "Signature inserted because this is the last test and has comment, is last test in department or is last department"
        );

        if (shouldAddPageBreak) {
          body.appendPageBreak();
        }
      }
    });

    // // Check if this is the last test
    // const isLastTest = currentTestCount === totalTests;
    // Logger.log("Is " + department + " last department and test: " + isLastTest);

    // // Append signature to the table
    // appendSignatureTable(body, 0, formData[9], isLastTest);
    // Logger.log("signature after department inserted");

    // // Only insert a page break if this is not the last department
    // if (index < departments.length - 1) {
    //   body.appendPageBreak();
    //   Logger.log("page break inserted because it is not the last test");
    // }
  });

  // const selectedTests = JSON.parse(formData[10]);
  // const labTests = getLabTests(formData);

  // selectedTests.forEach((selectedTest, index) => {
  //   if (labTests[selectedTest]) {
  //     const labTest = labTests[selectedTest];
  //     labTest.isLastTest = index === selectedTests.length - 1;
  //     insertTable(doc, labTest, formData);
  //   }
  // });
}

// For inserting test table
function insertTable(body, labTest, formData) {
  // const body = doc.getBody();

  // Insert heading before the table
  const heading = body.insertParagraph(
    body.getNumChildren(),
    labTest.tableHeading
  );
  heading
    .setHeading(DocumentApp.ParagraphHeading.HEADING2)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setBold(true)
    .setFontSize(12);

  const testTable = getTableFromTemplate(
    labTest.templateTableIndex,
    TEMPLATE_DOC_ID
  );

  // Check if the table was successfully retrieved
  if (!testTable) {
    Logger.log("Failed to retrieve the ${testType} table.");
    return;
  }

  // Iterate over table cells and replace placeholders with formData
  for (let row = 1; row < testTable.getNumRows(); row++) {
    const cell = testTable.getCell(row, 1);
    const cellText = cell.getText();

    // Find the corresponding test from the test objects
    const test = labTest.tests.find((t) => t.placeholder === cellText);
    if (test) {
      if (checkIfEmpty(test.result)) {
        if (row === testTable.getNumRows() - 1) {
          // If it's the last row, don't delete it, just clear its contents
          var numCells = testTable.getRow(row).getNumCells();
          for (var col = 0; col < numCells; col++) {
            testTable.getRow(row).getCell(col).clear();
          }
        } else {
          testTable.removeRow(row);
          row--; // Adjust the row index after deletion
        }
      } else {
        let result = test.result;
        // Format single-digit numbers as two digits
        if (!isNaN(result) && Number.isInteger(result) && result < 10) {
          result = "0" + result;
        }
        // Replace placeholder with result
        cell.setText(result);

        // const rangeCol = testTable.getCell(row, 3).getText();
        // const rangeColSplit = rangeCol.split('-');
        // const refLow = parseFloat(rangeColSplit[0].trim());
        // const refHigh = parseFloat(rangeColSplit[1].trim());

        // // Apply bold if the result is out of the reference range
        // if (test.result < refLow || test.result > refHigh) {
        //   cell.setBold(true);
        // }
        applyFormatting(testTable, row, result);
      }
    }
  }

  // Append the updated table to the document
  body.appendTable(testTable);

  // Append comment
  // if (formData[76] === 'Yes') {
  //     insertComment(doc, formData);
  // }

  // Append comment
  if (formData[76] === "Custom") {
    insertComment(body, formData);
  } else if (formData[76] === "Yes") {
    if (labTest.commentTableIndex) {
      appendCommentTable(body, labTest.commentTableIndex);
    } else {
      Logger.log("No comment index");
    }
  } else {
    Logger.log("No comment");
  }

  // If this isn't the last test type, add a page break
  // if (labTest.isLastTest !== true) {
  //   body.appendPageBreak();
  // }
}

// apply formatting based on reference range
function applyFormatting(testTable, row, result) {
  const headerRow = testTable.getRow(0);
  let referenceColIndex = null;

  // Search for the "Reference Range" column in the header row
  for (let col = 0; col < headerRow.getNumCells(); col++) {
    const headerText = headerRow.getCell(col).getText().trim();
    if (headerText === "Reference Range") {
      referenceColIndex = col;
      break;
    }
  }

  // Proceed with formatting only if the "Reference Range" column is found
  if (referenceColIndex !== null) {
    const rangeCol = testTable.getCell(row, referenceColIndex).getText();
    if (!checkIfEmpty(rangeCol)) {
      const [refLow, refHigh] = rangeCol
        .split("-")
        .map((v) => parseFloat(v.trim()));
      if (result < refLow || result > refHigh) {
        testTable.getCell(row, 1).setBold(true);
      }
    }
  }
}

// append signature table from signature template
function appendSignatureTable(
  body,
  signatureTableIndex,
  isSpecialData,
  isLastTest
) {
  // const body = doc.getBody();

  // If this is the last test type, add END OF REPORT
  if (isLastTest) {
    body
      .appendParagraph("** END OF REPORT **")
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setBold(true)
      .setFontSize(10);
  }

  const signatureTable = getTableFromTemplate(
    signatureTableIndex,
    SIGNATURE_TEMPLATE_DOC_ID
  );
  body.appendTable(signatureTable);
  if (isSpecialData === "Yes") {
    body
      .appendParagraph(
        "*This sample was processed at Medi Quest Laboratory Clinic."
      )
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setBold(false)
      .setFontSize(10);
  }
}

// append comment table from comment template
function appendCommentTable(body, commentTableIndex) {
  // const body = doc.getBody();

  const commentTable = getTableFromTemplate(
    commentTableIndex,
    COMMENT_TEMPLATE_ID
  );
  body.appendTable(commentTable);
}

// for comment
function insertComment(body, formData) {
  // const body = doc.getBody();
  const commentTable = getTableFromTemplate(0, COMMENT_TEMPLATE_ID);
  body.appendTable(commentTable);
  body.replaceText("{{Comment}}", formData[77]);
}
