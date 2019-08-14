function onOpen() {
  const ui = DocumentApp.getUi();
  ui.createMenu('Calculate Score')
  .addItem('Outcome', 'calculateOutcomeScore')
  .addToUi();
}

function getTables(tableTitle: string) {
  const tables = DocumentApp.getActiveDocument().getBody().getTables();
  console.log(tables);
}