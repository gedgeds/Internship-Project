function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {

  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Add-on Menu')
  .addSubMenu(ui.createMenu('Ðablonai')
              .addItem('Media plano ðablonas', 'openMediaPlanModalDialog')
              .addItem('Uþsakymo ðablonas', 'openOrderModalDialog'))
  .addSeparator()
  .addSubMenu(ui.createMenu('Parankiniai')
              .addItem('Atnaujinti duomenø bazæ', 'updateDatabase')
              .addItem('Generuoti áraðus', 'generateRecords')
              .addItem('Generuoti kalendoriø', 'generateCalendar')
              .addItem('Uþbaigti ðablonà', 'finalizeMediaPlanTemplate'))
  .addToUi();
  
}
