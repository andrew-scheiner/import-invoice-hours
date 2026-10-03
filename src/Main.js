/*
GASLibrary.copyDynamicRangeToTargetSheet
  instead of alert, add an update to a sell
copyFormInvoiceHoursToInvoicesHours
  add a completion notification to the cell (see above) 
copyLearnerNamesWithEndDate
  add contractor initials to the sheet copyset
  add number of hours to sheet copyset
  see if this can be simplified by using  
    GASLibrary.getNumberOfDataRowsInSingleColumnRange and
    GASLibrary.copyDynamicRangeToTargetSheet
  ?change learnerName in function name and elsewhere to serviceUserName
  ?include ContractorInitials
  ?more namedRange variables to Config
  get other data items for Report Monthly Invoice
    Contract Name
    contractor Initials
    Contract type 
*/

function createMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom')
    .addItem('Backup Spreadsheet', 'backupSpreadsheet')
    .addItem('Sort Sheet', 'sortActiveSheet')
    .addItem('Update Data Source(s)', 'updateDataSources')
    .addItem('Update Invoicee List', 'copyLearnerNamesWithInvoiceDate')
    .addToUi();
}

function getGCalInvoiceEvents() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName('Form');

  // Get filter values
  const calendarId = spreadsheet.getRangeByName('form__calendar').getValue();
  const invoiceTime = spreadsheet.getRangeByName('form__invoiceTime').getValue();
  const reportFrom = new Date(spreadsheet.getRangeByName('form__startDate').getValue());
  const reportTo = new Date(spreadsheet.getRangeByName('form__endDate').getValue());

  // Adjust report end date/time to end of the day
  reportTo.setHours(23, 59, 59);

  // Get events from the specified calendar
  const cal = CalendarApp.getCalendarById(calendarId);
  const events = cal.getEvents(reportFrom, reportTo, { search: invoiceTime });

  // Prepare data for batch insertion
  const data = events.map((event) => {
    const startTime = event.getStartTime();
    const endTime = event.getEndTime();
    const duration = (endTime - startTime) / 60000 / 60; // Duration in hours

    return [
      startTime, // Column 1: Start date
      calendarId, // Column 2: Calendar ID
      event.getTitle(), // Column 3: Event title
      startTime, // Column 4: Start time
      endTime, // Column 5: End time
      duration, // Column 6: Duration
      event.getDescription(), // Column 7: Description
    ];
  });

  // Determine number of columns dynamically
  const numCols = data.length > 0 ? data[0].length : 0;

  // Clear previous content if there are columns to clear
  if (numCols > 0) {
    const reportStartRow = INVOICE_HOURS_START_ROW;
    const reportStartColumn = INVOICE_HOURS_START_COL;
    const numRows = sheet.getLastRow() - reportStartRow + 1;
    sheet.getRange(reportStartRow, reportStartColumn, numRows, numCols).clearContent();
  }

  // Write data to the sheet
  if (data.length > 0) {
    const reportStartRow = INVOICE_HOURS_START_ROW;
    const reportStartColumn = INVOICE_HOURS_START_COL;
    const range = sheet.getRange(reportStartRow, reportStartColumn, data.length, numCols);
    range.setValues(data);

    // Apply formatting
    sheet.getRange(reportStartRow, reportStartColumn).setNumberFormat('dd/mm/yyyy'); // Start date
    sheet.getRange(reportStartRow, reportStartColumn + 3, data.length, 1).setNumberFormat('hh:mm'); // Start time
    sheet.getRange(reportStartRow, reportStartColumn + 4, data.length, 1).setNumberFormat('hh:mm'); // End time
    sheet.getRange(reportStartRow, reportStartColumn + 5, data.length, 1).setNumberFormat('0.00'); // Duration
  }
}

function submitFormValues() {
  copyFormInvoiceHoursToInvoicesHours();
  copyLearnerNamesWithInvoiceDate();
}

function copyFormInvoiceHoursToInvoicesHours() {
  const sourceSpreadsheetId = SS_ID;
  const sourceSheetName = FORM_SHEET_NAME;
  const sourceRangeStartRow = INVOICE_HOURS_START_ROW;
  const sourceRangeStartColumn = INVOICE_HOURS_START_COL;
  const sourceRangeColumnWidth = INVOICE_HOURS_COL_WIDTH;

  GASLibrary.copyDynamicRangeToTargetSheet(
    sourceSpreadsheetId,
    sourceSheetName,
    sourceRangeStartRow,
    sourceRangeStartColumn,
    sourceRangeColumnWidth,
    INVOICE_HOURS_ID,
    INVOICE_HOURS
  );
}

function copyLearnerNamesWithInvoiceDate() {
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceRange = sourceSpreadsheet.getRangeByName('form__learnerNames');
  const sourceValues = sourceRange
    .getValues()
    .flat()
    .filter((name) => name); // Flatten & clean
  const targetSpreadsheetId = REPORT_MONTHLY_INVOICE_LIST_ID; // Report Monthly Invoice List

  let invoiceDate;

  if (INVOICE_TIME === '*CPE') {
    invoiceDate = END_DATE;
  }
  if (INVOICE_TIME === '*CPB') {
    invoiceDate = START_DATE;
  }

  let contractorInitials;

  if (CALENDAR === 'andrew.scheiner@brightpath.services') {
    contractorInitials = 'AAS';
  }
  if (CALENDAR === 'helen.sender@brightpath.services') {
    contractorInitials = 'HRS';
  }

  const outputArray = sourceValues.map((name) => [invoiceDate, '', name, contractorInitials]);

  const targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  const targetRange = targetSpreadsheet.getRangeByName('invoiceList__invoiceDate');
  const sheet = targetRange.getSheet();
  const targetStartColumn = targetRange.getColumn();

  // Get the last non-empty row in the first column of the named range
  const dataColumnRange = sheet.getRange(
    targetRange.getRow(),
    targetStartColumn,
    sheet.getLastRow() - targetRange.getRow() + 1
  );
  const values = dataColumnRange.getValues().flat();
  let lastRowOffset = values.map(String).filter((v) => v.trim()).length;

  const targetStartRow = targetRange.getRow() + lastRowOffset;
  sheet.getRange(targetStartRow, targetStartColumn, outputArray.length, 4).setValues(outputArray);

  const targetUrl = targetSpreadsheet.getUrl();
  const statusCell = sourceSpreadsheet.getRangeByName('form__invoiceListStatus');
  statusCell.setFormula(`=HYPERLINK("${targetUrl}", "Updated")`);
}

function clearForm() {
  const formData = SS.getRangeByName('form__formInputData');
  formData.clearContent();

  const invoiceHoursReportNumRows = GASLibrary.getNumberOfDataRowsInSingleColumnRange(
    FORM_SHEET,
    INVOICE_HOURS_START_ROW,
    INVOICE_HOURS_START_COL
  );
  const invoiceHoursReport = FORM_SHEET.getRange(
    INVOICE_HOURS_START_ROW,
    INVOICE_HOURS_START_COL,
    invoiceHoursReportNumRows,
    INVOICE_HOURS_COL_WIDTH
  );
  invoiceHoursReport.clearContent();

  const targetCell = SS.getRangeByName('form__calendar');
  targetCell.activate();
}
