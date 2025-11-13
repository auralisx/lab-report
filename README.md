# Lab Report Automation Using Google Apps Script

This project automates the generation of lab reports from form submissions into a Google Sheet, using Google Apps Script to streamline the process. Once test data is entered through a form, the script organizes, formats, and exports the data into a structured lab report in Google Docs. The script is built to maintain readability by grouping tests by department and preventing page breaks within individual test tables.

> [!NOTE]
> This project was intended to serve as a temporary automation solution until a complete Lab Report Management System is implemented. It provides a functional and effective way to handle lab report generation for the interim.

## Project Overview

This Google Apps Script project simplifies the lab reporting process by automating the following steps:

1. **Trigger Script on Form Submission**
   - Detects changes in the Google Sheet and only triggers if the new data row is appended from a form submission, ignoring manual data edits.

2. **Create Patient-Specific Folders**
   - Generates a folder named after the patient within the specified REPORT_FOLDER_ID, organizing reports for easy retrieval.

3. **Generate and Customize Report Document**
   - Duplicates a template document (REPORT_TEMPLATE_ID) and fills in placeholder data in the header using data from the form submission (`formData`), then begins table insertion.

4. **Categorize Tests by Department**
   - Groups tests under specific departments and inserts department-specific tables, maintaining an organized layout for better readability.

5. **Check and Prevent Page Breaks within Tables**
   - Ensures each table fully displays on a single page by checking space availability. If needed, it inserts a page break between departments instead of within tables.

6. **Format Report for Readability**
   - Structures the report with department-specific sections and ensures clean layout transitions, including adding a department header.

7. **Dynamic Placeholder Replacement**
   - Replaces placeholders (enclosed by `<< >>` inside templates) with `formData` values as defined in the `testObjects.gs` file.

8. **Insert Signature**
   - Inserts a signature in a table with invisible borders to complete the report.

9. **Export Report as DOC and PDF**
   - Generates both DOC and PDF versions, saved in separate folders with headers and without headers as required.

10. **Send Report via Email**
   - Emails the report files and a folder link to the specified Gmail address, simplifying report delivery.

### Folder Structure

- **docs**: Contains simple documentation and guidelines for using the script and how it works.
- **templates**: Contains DOC templates used for table creation in the report.
- **testing**: Holds experimental functions and logic for feature enhancements and testing purposes. Only there for my reference, Not essential for primary automation.

## Known Issues

- **Limited Page Break Control**
  - Google Docs API has limitations on precise layout control, which may lead to minor inconsistencies in the table layout.

- **Estimated Table Height**
  - Table height is estimated based on average row size, which can vary slightly depending on font style or size, affecting spacing.

---

This script provides a practical solution for temporary lab report automation and is a stepping stone towards a comprehensive Lab Report Management System.
