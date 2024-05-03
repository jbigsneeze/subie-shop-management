// These constant values may be configured as desired to
// match their corresponding values in the spreadsheet.

const INPUT_SHEET_NAME         = "1 - Input";
const WORK_ORDER_SHEET_NAME    = "2 - Work Orders";
const CUSTOMER_SHEET_NAME      = "3 - Customers";
const PARTS_SERVICE_SHEET_NAME = "4 - Parts & Services";
const INVOICE_SHEET_NAME       = "6 - Printable Invoice";

const INPUT_CUSTOMER_COLUMN  = 4;
const INPUT_WORK_ITEM_COLUMN = 8;
const INPUT_AMOUNT_COLUMN    = 9;

const INPUT_NAME_ROW     =  3;
const INPUT_PHONE_ROW    =  4;
const INPUT_ADDRESS_ROW  =  5;
const INPUT_YEAR_ROW     =  7;
const INPUT_MODEL_ROW    =  8;
const INPUT_TRIM_ROW     =  9;
const INPUT_PLATE_ROW    = 10;
const INPUT_ODOMETER_ROW = 11;
const INPUT_VIN_ROW      = 12;
const INPUT_MECHANIC_ROW = 14;
const INPUT_1ST_ITEM_ROW =  4;

const WO_ID_COLUMN        =  1;
const WO_DATE_COLUMN      =  2;
const WO_NAME_COLUMN      =  3;
const WO_YEAR_COLUMN      =  4;
const WO_MODEL_COLUMN     =  5;
const WO_TRIM_COLUMN      =  6;
const WO_PLATE_COLUMN     =  7;
const WO_ODOMETER_COLUMN  =  8;
const WO_MECHANIC_COLUMN  =  9;
const WO_WORK_DESC_COLUMN = 10;
const WO_AMOUNT_COLUMN    = 11;
const WO_TAX_COLUMN       = 12;
const WO_SUBTOTAL_COLUMN  = 13;
const WO_TOTAL_COLUMN     = 14;
const WO_PHONE_COLUMN     = 15;
const WO_ADDRESS_COLUMN   = 16;
const WO_VIN_COLUMN       = 17;

const CUSTOMER_1ST_DATA_ROW      =  2;
const CUSTOMER_NAME_COLUMN       =  1;
const CUSTOMER_FIRST_NAME_COLUMN =  2;
const CUSTOMER_YEAR_COLUMN       =  4;
const CUSTOMER_MODEL_COLUMN      =  5;
const CUSTOMER_TRIM_COLUMN       =  6;
const CUSTOMER_PLATE_COLUMN      =  7;
const CUSTOMER_PHONE_COLUMN      =  8;
const CUSTOMER_ADDRESS_COLUMN    =  9;
const CUSTOMER_VIN_COLUMN        = 10;

const PARTS_SERV_1ST_DATA_ROW = 2;
const PARTS_SERV_DESC_COLUMN  = 1;
const PARTS_SERV_PRICE_COLUMN = 2;

const INVOICE_1ST_ITEM_ROW  = 13;
const INVOICE_WORK_ITEM_COL =  3;
const INVOICE_PRICE_COL     =  8;
const INVOICE_TAX_COL       =  9;
const INVOICE_LINE_ITEMS    = 25;

const INVOICE_NAME_CELL      = "E3";
const INVOICE_ADDRESS_CELL   = "E4";
const INVOICE_PHONE_CELL     = "E5";
const INVOICE_VEHICLE_CELL   = "H3";
const INVOICE_PLATE_CELL     = "H4";
const INVOICE_MILES_CELL     = "H5";
const INVOICE_VIN_CELL       = "H6";
const INVOICE_ID_CELL        = "D8";
const INVOICE_MECHANIC_CELL  = "F8";
const INVOICE_DATE_CELL      = "H8";
const INVOICE_SUBTOTAL_CELL  = "H38";
const INVOICE_TOTAL_TAX_CELL = "H39";
const INVOICE_TOTAL_CELL     = "H40";

const SUCCESS_COLOR = "#00b900";
const PROBLEM_COLOR = "red";

const IOWA_TAX = 0.07;

const CELL_INPUT_WARNING = "(Don't forget to press enter after typing in a cell.)";

// Leave these alone
const doc = SpreadsheetApp.getActiveSpreadsheet();
const ui = SpreadsheetApp.getUi();
const input_sheet = doc.getSheetByName(INPUT_SHEET_NAME);
const wo_sheet = doc.getSheetByName(WORK_ORDER_SHEET_NAME);
const cust_sheet = doc.getSheetByName(CUSTOMER_SHEET_NAME);
const ps_sheet = doc.getSheetByName(PARTS_SERVICE_SHEET_NAME);
const inv_sheet = doc.getSheetByName(INVOICE_SHEET_NAME);

// Useful utility function
function getLastRowOfColumn(worksheet,column)
{
  let row_index = worksheet.getLastRow();
  let last_cell_maybe = worksheet.getRange(row_index,column);

  if(last_cell_maybe.getValue() === "")
  {
    // Another column has more rows filled out than this one. Find the last filled cell in this column
    row_index = last_cell_maybe.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }

  return row_index;
}
