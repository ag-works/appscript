function getData(rowNumber) {
  let ss = SpreadsheetApp.getActiveSpreadsheet()
  let dataSheet = ss.getSheetByName("Quotation (Response)");
  let lastColumnNumber = dataSheet.getDataRange().getLastColumn();
  let headers = (dataSheet.getRange("1:1").getDisplayValues()[0] || []).filter(value => !!value);
  
  let rows = dataSheet.getRange(rowNumber, 1, 1, lastColumnNumber).getDisplayValues();
  let lastRow = rows.length > 0 ? rows[0] : [];
  let data = headers.reduce( (obj, header, idx) => {
    obj[header] = lastRow[idx] || "";
    return obj;
  }, {});
  return data;
}
  

function makePdf(rowNumber=null) {
  let settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings")
  let dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Quotation (Response)")
  let tempFolderUrl = settingsSheet.getRange("B1").getDisplayValue();
  let pdfFolderUrl = settingsSheet.getRange("B2").getDisplayValue();
  let docFileUrl = settingsSheet.getRange("B3").getDisplayValue();
  let identifierColumn = settingsSheet.getRange("B4").getDisplayValue();
  let urlColumn = settingsSheet.getRange("B5").getDisplayValue();
  let headers = (dataSheet.getRange("1:1").getDisplayValues()[0] || []).filter(value => !!value);

  if(!rowNumber){
    let dataRows = dataSheet.getDataRange().getDisplayValues();
    for(let i=1; i < dataRows.length; i++){
      let row = dataRows[i];
      if(row[0] == ""){
        break
      }
      rowNumber = i + 1;
    }
  }

  let tempFolderId = tempFolderUrl.split("/")[5];
  let pdfFolderId = pdfFolderUrl.split("/")[5];
  let docFileId = docFileUrl.split("/")[5];

  let data = getData(rowNumber);
  let filename = data[identifierColumn];
  let dummyFile = DriveApp.getFileById(docFileId).makeCopy(tempFolderId).setName(filename);
  let openFile = DocumentApp.openById(dummyFile.getId());

  let body = openFile.getBody();
  for(let field in data) {
    let escapedField = field.replace("(", "\\(").replace(")", "\\)")
    body.replaceText(`{${escapedField}}`, data[field]);
  }
  openFile.saveAndClose();

  let blob = dummyFile.getAs("application/pdf").setName(filename);
  let pdfFile = DriveApp.createFile(blob).moveTo(DriveApp.getFolderById(pdfFolderId));
  DriveApp.getFolderById(tempFolderId).removeFile(dummyFile);

  let urlColumnNumber = headers.indexOf(urlColumn) + 1;
  if(urlColumnNumber == -1){
    console.error("Could not find URL column in headers!");
    return;
  }
  dataSheet.getRange(rowNumber, urlColumnNumber).setValue(pdfFile.getUrl());
}

function generatePdf(){
  makePdf(7);
}