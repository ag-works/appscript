const CELL_WIDTH = 40;
const CELL_HEIGHT = 40;
const NUMBER_OF_MINES = 50;
const MINE_IMAGE = "https://raymanpc.com/wiki/script-en/images/a/a0/Mine.png";
const GAME_PLAY_TIME = 300;
const mineImage = SpreadsheetApp
                    .newCellImage()
                    .setSourceUrl(MINE_IMAGE)
                    .setAltTextTitle("")
                    .setAltTextDescription("")
                    .build();

function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("Play")
    .addItem("Start", "showRules")
    .addItem("Start Over", "showRules")
    .addToUi();

  setup();
  showRules();
}

function setup(){
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let rows = sheet.getMaxRows();
  let columns = sheet.getMaxColumns();
  sheet.setTabColor("00ee00")
  sheet.setRowHeights(1, rows, CELL_HEIGHT);
  sheet.setColumnWidths(1, columns, CELL_WIDTH);
  sheet.getRange(1,1,rows, columns).setBackground("#aaaaaa").clearContent();
}

function showRules(){
  let rulesInterface = HtmlService.createHtmlOutputFromFile("rules")
                    .setWidth(320)
                    .setHeight(100);
  
  // SpreadsheetApp.getUi().showModalDialog(rulesInterface, "Rules");
  let ui = SpreadsheetApp.getUi();
  var response = ui.alert('Rules', rulesInterface.getContent(), ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.YES) {
    start();
  } else if (response == ui.Button.NO) {
    Logger.log('The user didn\'t want to play.');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function gameOver(mines){
  let ui = SpreadsheetApp.getUi();
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let rows = sheet.getMaxRows();
  let columns = sheet.getMaxColumns();
  sheet.getRange(1,1,rows, columns).setBackground("#aa0000");


  // Setting mines
  for(let mine in mines){
    let [i,j] = mine.split(":");
    i = parseInt(i);
    j = parseInt(j);
    // sheet.getRange(i, j).setValue(mineImage);
    let iconsSheet = SpreadsheetApp.getActive().getSheetByName("Icons")
    iconsSheet.getRange(1,1).copyTo(sheet.getRange(i, j));
    Utilities.sleep(50);
    SpreadsheetApp.flush();
  }
  let interface = HtmlService.createHtmlOutput(`<img src="${MINE_IMAGE}" width="300" height="300">`)
                    .setWidth(320)
                    .setHeight(320);
  ui.showModalDialog(interface, "Game Over");
}

function start(){
  setup();
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let rows = sheet.getMaxRows();
  let columns = sheet.getMaxColumns();
  let mines = {};
  for(let i=1; i<= NUMBER_OF_MINES; i++){
    let i = Math.ceil( Math.random() * rows );
    let j = Math.ceil( Math.random() * columns );
    mines[`${i}:${j}`] = true;
  }
  let documentCache = CacheService.getDocumentCache();
  documentCache.put('MINES', JSON.stringify(mines));
  documentCache.put('GAME_STARTED', JSON.stringify(true));
  documentCache.put('GAME_OVER', JSON.stringify(false));

  let timerStartTime = Date.now();
  while(true){
    let gameEnded = documentCache.get('GAME_OVER');
    if(gameEnded === 'true'){
      break;
    }
    let timeChange = parseInt((Date.now() - timerStartTime) / 1000);
    sheet.getRange(1,1).setValue(`${ GAME_PLAY_TIME - timeChange } s`);
    SpreadsheetApp.flush();
  }
  // console.log(mines);
}

function getSelection(){
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let column = sheet.getActiveCell().getColumn();
  let row = sheet.getActiveCell().getRow();
  return [row, column];
  // console.log(sheet.getSelection().getCurrentCell().getA1Notation());
}


function onSelectionChange(e){
  let documentCache = CacheService.getDocumentCache();
  let gameStarted = documentCache.get('GAME_STARTED');
  // console.log(gameStarted, typeof(gameStarted));
  if(gameStarted !== 'true'){
    return;
  }

  let mines = JSON.parse(documentCache.get("MINES"));
  let sheet = SpreadsheetApp.getActive().getActiveSheet();
  let row = e.range.rowStart;
  let column = e.range.columnStart;
  if(`${row}:${column}` in mines){
    documentCache.put('GAME_STARTED', "false");
    documentCache.put('GAME_OVER', "true");
    sheet.getRange(row, column).setValue(mineImage);
    gameOver(mines);
  } else {
    // console.log("Clearing format for Cell", `${row}:${column}`);
    sheet.getRange(row, column).clearFormat();
  }
}