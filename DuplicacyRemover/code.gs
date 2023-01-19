let ui = SpreadsheetApp.getUi();

/**
 * Event Listener for onOpen event of Spreadsheet
 */
function onOpen(){
  SpreadsheetApp.getUi()
                .createMenu("MyLab")
                .addItem("Launch Duplicacy Remover", "launchDuplicacyRemover")
                .addToUi();
}


function launchDuplicacyRemover() {
  let modalContent = HtmlService.createHtmlOutputFromFile("dialog")
    .setWidth(900)
    .setHeight(510);
  ui.showModalDialog(modalContent, "Launch Duplicate Remover");
}
