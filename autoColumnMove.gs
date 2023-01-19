function onEdit(e) {
  const src = e.source.getActiveSheet();
  const r = e.range;
  if (r.columnStart != 13 || r.rowStart == 1 || e.value == src.getName()) return;
  if (r.columnStart != r.columnEnd) return;

  const targetSheetName = e.value;
  const allSheetNames = SpreadsheetApp.getActiveSpreadsheet().getSheets().map(sheet => sheet.getName());
  if(!allSheetNames.includes(targetSheetName)){
    return;
  }

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
  const targetSheetLastRowNumber = targetSheet.getDataRange().getLastRow() + 1;
  const targetSheetLastColumnNumber = targetSheet.getDataRange().getLastColumn() + 1;
  console.log(targetSheetLastRowNumber, targetSheetLastColumnNumber);
  src.getRange(r.rowStart, 1, 1, targetSheetLastColumnNumber).moveTo(targetSheet.getRange(targetSheetLastRowNumber, 1, 1, targetSheetLastColumnNumber));
  // src.deleteRow(r.rowStart);
}
