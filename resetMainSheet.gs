/**
 * @OnlyCurrentDoc
 * Create an initial empty calendar for every employee listed in the
 * employee directory. Useful for removing terminated employee.
 * @customfunction
 */
function resetMainSheet() {
  makeSafetyCopy();

  if(mainSheet.getMaxColumns() < fullEmployeeChart.getWidth()) {
    mainSheet.insertColumnsAfter(24, fullEmployeeChart.getWidth() - mainSheet.getMaxColumns());
  }
  fullEmployeeChart.copyTo(mainSheet.getRange(1,1), SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);

  let lastRow = directorySheet.getLastRow() - 1;
  for(let i = 1, count = 1; i <= lastRow; ++i) {
    fullEmployeeChart.copyTo(mainSheet.getRange(count,1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
    mainSheet.getRange(count + 2,2).setValue(directoryNames[i][0]);
    count += fullEmployeeChart.getHeight();
  }
  mainSheet.setFrozenColumns(10);
}


/**
 * @OnlyCurrentDoc
 * Replace values in 'MainSheet' with values in sheet.
 * @param sheet - default grabs the safety sheet created by makeSafetyCopy()
 */
function restoreMainSheet(sheet=spreadsheet.getSheetByName(SAFETY_COPY_NAME)) {
  let copyValues = sheet.getDataRange();
  let rows = copyValues.getLastRow();
  let cols = copyValues.getNumColumns();
  copyValues.copyTo(mainSheet.getRange(1,1,rows,cols),SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
}


/**
 * @OnlyCurrentDoc
 * SLOW
 * Create a copy of 'MainSheet'. Used to restore sheet upon error.
 * @return - copy of sheet
 */
function makeSafetyCopy() {
  if(spreadsheet.getSheetByName(SAFETY_COPY_NAME)) {
    spreadsheet.deleteSheet(spreadsheet.getSheetByName(SAFETY_COPY_NAME));
  }
  let copy = mainSheet.copyTo(spreadsheet);
  copy.setName(SAFETY_COPY_NAME);
  copy.hideSheet();

  let destFolder = DriveApp.getFolderById('********************');
  let date = Utilities.formatDate(new Date(), 'GMT-8', 'dd_MM_yyyy_HH_mm');
  DriveApp.getFileById('**********************').makeCopy(`SafetyCopy_${date}`, destFolder);
}















