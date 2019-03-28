function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Message Creator')
  .addItem('Create Messages', 'create')
  .addToUi();
}