/** @OnlyCurrentDoc
 * To be run to update the PTO banks of each employee. Import the
 * .xlsx as a new sheet and run this.
 * @customfunction
 */
function updateBreakdownSheet() {
  let bRange = spreadsheet.getRangeByName("FullBreakdown");
  let newBreakdown = spreadsheet.getSheetByName("Sheet");
  let newBreakdownRange = newBreakdown.getRange(1,1,bRange.getHeight(),bRange.getWidth());
  newBreakdownRange.copyTo(breakdownSheet.getRange(1,1,bRange.getHeight(),bRange.getWidth()));
  spreadsheet.deleteSheet(newBreakdown);
  removeCommas();
  refreshMainSheet();
}


/** @OnlyCurrentDoc
 * Helper function - called by 'updateBreakdownSheet'
 * Removes commas from names in imported 'Breakdown' sheet, necessary
 * as all other reports are without commas. Prevents false negatives
 * in string comparison operations.
 * @customfunction
 */
function removeCommas() {
  let stop = 0;
  for(let i = 2; i < breakdownSheet.getLastRow(); ++i) {
    if(breakdownNames[i][0] === "") {
      ++stop;
      if(stop === 3) {
        break;
      }
      continue;
    }
    if(breakdownNames[i][0].indexOf(",") > -1) {
      let range = breakdownSheet.getRange(i + 1,1);
      let change = range.getValue().replace(",","");
      range.setValue(change);
      stop = 0;
    }
  }
}


/** @OnlyCurrentDoc
 * Helper function - called by 'updateBreakdownSheet'
 * Silly way of refreshing sheet information
 * Copies all calendars in bulk down one row then
 * deletes row 1
 * @customfunction
 */
function refreshMainSheet() {
  let count = directoryN.getLastRow();
  let mRange = mainSheet.getRange(1,1,count * fullEmployeeChart.getHeight(),fullEmployeeChart.getWidth());
  mRange.copyTo(mainSheet.getRange(2,1,count * fullEmployeeChart.getHeight(),fullEmployeeChart.getWidth()));
  mainSheet.deleteRow(1);
}
