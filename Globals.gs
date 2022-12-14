/**
 * All sheets and ranges used in the PTO Tracker.
 * Referenced throughout scripts.
 */

var spreadsheet = SpreadsheetApp.getActive();

//Sheets
var mainSheet = spreadsheet.getSheetByName('MainSheet');
var updateSheet = spreadsheet.getSheetByName('UpdateSheet');
var directorySheet = spreadsheet.getSheetByName("EmployeeDirectory");
var reportSheet = spreadsheet.getSheetByName('Details');
var breakdownSheet = spreadsheet.getSheetByName('Breakdown');

//Directory ranges
var directoryN = spreadsheet.getRangeByName("DirectoryNames");
var directoryNames = directoryN.getDisplayValues();
var directoryHireDate = spreadsheet.getRangeByName("DirectoryHireDate").getDisplayValues();

//Breakdown ranges
var breakdownNames = spreadsheet.getRangeByName("BreakdownNames").getValues();
var breakdownType = spreadsheet.getRangeByName("BreakdownType").getValues();
var breakdownBalance = spreadsheet.getRangeByName("BreakdownBalance").getValues();

//Details ranges
var reportNames = spreadsheet.getRangeByName("ReportNames").getDisplayValues();
var reportPolicy = spreadsheet.getRangeByName("ReportPolicy").getValues();
var reportDates = spreadsheet.getRangeByName("ReportDates").getValues();
var reportStatus = spreadsheet.getRangeByName("ReportStatus").getValues();
var reportHours = spreadsheet.getRangeByName("ReportHours").getValues();
var reportReviewedBy = spreadsheet.getRangeByName("ReportReviewedBy").getValues();
var reportSubmittedDate = spreadsheet.getRangeByName("ReportSubmittedDate").getValues();

//Update ranges
var fullEmployeeChart = spreadsheet.getRangeByName("FullEmployeeChart");

//Change every year
const CURRENT_YEAR = '2022';

//Default name for safety sheets
const SAFETY_COPY_NAME = 'SafetyCopy';








