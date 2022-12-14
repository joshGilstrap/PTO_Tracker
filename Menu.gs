function onOpen(e) {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('PTO Runner')
    .addItem('Run new report', 'updateMainSheet')
    .addSeparator()
    .addItem('Restore last change', 'restoreMainSheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('Refresh sheets')
      .addItem('Breakdown', 'updateBreakdownSheet')
      .addSeparator()
      .addItem('Details', 'updateDetailsSheet'))
    .addSeparator()
    .addItem('Undo', 'restoreMainSheet')
    .addSeparator()
    .addItem('Reset', 'resetMainSheet')
    .addToUi();
}
