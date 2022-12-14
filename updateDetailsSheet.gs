/** @OnlyCurrentDoc
 * Used for replacing 'Details' sheet with updated details
 * while preserving named ranges integral to program funciton
 * @customfunction
 */
function updateDetailsSheet() {
  spreadsheet.deleteSheet(spreadsheet.getSheetByName("Summary"));
  let newSheet = spreadsheet.getSheetByName("Details (1)");
  let newSheetRange = newSheet.getRange(1,1,newSheet.getLastRow(),26);
  reportSheet.clear();
  newSheetRange.copyTo(reportSheet.getRange(1,1,newSheet.getLastRow(),26));
  spreadsheet.deleteSheet(newSheet);
}
