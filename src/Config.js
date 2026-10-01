// =================Global constants =================================

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SS_ID = SS.getId();

const TODAYS_DATE = new Date();

// Declare Form constants
const FORM_SHEET_NAME = 'Form';
const FORM_SHEET = SS.getSheetByName(FORM_SHEET_NAME);

const START_DATE = GASLibrary.getNamedRangeValue(SS, 'form__startDate');
const END_DATE = GASLibrary.getNamedRangeValue(SS, 'form__endDate');

const CALENDAR = GASLibrary.getNamedRangeValue(SS, 'form__calendar');
const INVOICE_TIME = GASLibrary.getNamedRangeValue(SS, 'form__invoiceTime');

//  Declare Invoice Hours report parameters
const INVOICE_HOURS_START_ROW = 2;
const INVOICE_HOURS_START_COL = 7;
const INVOICE_HOURS_COL_WIDTH = 7;

const REPORT_MONTHLY_INVOICE_LIST_ID = '1KRhirjLfWGtBqG8dTE-wKZbTmFk4SdrRak-eGxcKCV4';

const { INVOICE_HOURS_ID } = GASConfigLibrary.getSpreadsheetIds();
const { INVOICE_HOURS } = GASConfigLibrary.getSheetNames();

const SORT_CONFIGS = {
  Cpi: {
    sortColumns: [
      { column: 3, ascending: true },
      { column: 4, ascending: true },
      { column: 5, ascending: true },
      { column: 2, ascending: true },
    ],
    headerRows: 1,
  },
};
