function onOpen(){
  let ui = SpreadsheetApp.getUi();
  ui.createMenu("My Labs").addItem("Set API Key", "set_api_key").addToUi();
}

let SPREADSHEET_URL_INDEX = 0;
let SHEET_NAME_COLUMN_INDEX = 1;
let RANGE_COLUMN_INDEX = 2;
let MESSAGE_COLUMN_INDEX = 3;
let RECEIVER_COLUMN_INDEX = 4;
let STATUS_COLUMN_INDEX = 5;


function exportRangesAndSendWhatsappMessage(){
  let contentSheetName = 'export-ranges';
  let contentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(contentSheetName);
  let rows = contentSheet.getDataRange().getDisplayValues();
  for(let row_number=1; row_number < rows.length; ++row_number){
    let row = rows[row_number];
    let sheetUrl = row[SPREADSHEET_URL_INDEX];
    let sheetName = row[SHEET_NAME_COLUMN_INDEX];
    let range = row[RANGE_COLUMN_INDEX];
    let message = row[MESSAGE_COLUMN_INDEX];
    let receiver = row[RECEIVER_COLUMN_INDEX];
    if(!sheetName){
      continue;
    }
    let pdfId = exportRange(sheetUrl, sheetName, range);
    let pdfContent = getGoogleFileAsBase64(pdfId)
    var status = send_whatsapp_message(receiver, message, [], get_api_key(), pdfContent);
    Logger.log("Got " + status + " in sending whatsapp message to " + String(receiver));
    contentSheet.getRange(row_number + 1, STATUS_COLUMN_INDEX + 1).setValue(status.toUpperCase() + ", " + (new Date()).toString());

    DriveApp.getFileById(pdfId).setTrashed(true);
  }
}


function exportRange(sheetUrl, sheetName, namedRange){
  let sheetToCopy = SpreadsheetApp.openByUrl(sheetUrl).getSheetByName(sheetName);

  /** Creating new spreadsheet for range export */
  var spreadsheetForExport = SpreadsheetApp.create("export" + "-" + sheetToCopy.getName() + "-" + String(Date.now()));
  var exportSpreadsheetId = spreadsheetForExport.getId();
  var sheets = spreadsheetForExport.getSheets();
  var destinationSheet = sheets[0];
  var copiedSheet = sheetToCopy.copyTo(spreadsheetForExport);
  var range = copiedSheet.getRange(namedRange);
  range.setValues(sheetToCopy.getRange(namedRange).getDisplayValues());
  range.copyTo(destinationSheet.getRange(1, 1, range.getLastRow(), range.getLastColumn()));
  spreadsheetForExport.deleteSheet(copiedSheet);

  // DriveApp.getFileById(exportSpreadsheetId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var authToken = ScriptApp.getOAuthToken();
  var folder = DriveApp.getRootFolder();
  var pdfContent = getBlob(exportSpreadsheetId, authToken)
  var pdfFile = folder.createFile(pdfContent).setName(spreadsheetForExport.getName() + '.pdf');

  DriveApp.getFileById(exportSpreadsheetId).setTrashed(true);
  // Browser.msgBox('copied');
  return  pdfFile.getId();
}


function getBlob(spreadsheetId, token){
  var url = 'https://docs.google.com/spreadsheets/d/';
  var id = spreadsheetId;
  var url_ext = '/export?'
  +'format=pdf'
  +'&size=A4'                      //A3/A4/A5/B4/B5/letter/tabloid/legal/statement/executive/folio
  +'&portrait=false'                //true= Potrait / false= Landscape
  +'&scale=2'                      //1= Normal 100% / 2= Fit to width / 3= Fit to height / 4= Fit to Page
  +'&top_margin=0.50'              //All four margins must be set!
  +'&bottom_margin=0.50'           //All four margins must be set!
  +'&left_margin=0.50'             //All four margins must be set!
  +'&right_margin=0.50'            //All four margins must be set!
  +'&gid=0';
  console.log(url+id+url_ext);
  var response = UrlFetchApp.fetch(url+id+url_ext, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
  });

  var blob = response.getBlob().getAs('application/pdf');
  return blob;
}

/** Calling WhatsApp API for sending message */
function send_whatsapp_message(
  mobile_no,
  message,
  media_link,
  apiKey,
  pdfContent) {
  Logger.log([mobile_no, message, media_link]);
  var messages = [].concat(message || []);
  var media_links = [];
  if(typeof(media_link) === 'string' && media_link.length > 0) {
    media_links = media_link.split(',');
  }
  var mobile_nos = [].concat(mobile_no || []);
  var recipient_ids = mobile_nos.filter(function (number) {
    return String(number).endsWith("@c.us") || String(number).endsWith("@g.us")
  });
  var receiver_numbers = mobile_nos.filter(function (number) {
    return !recipient_ids.includes(number);
  });

  var payload = {
    receiverMobileNo: receiver_numbers.join(","),
    recipientIds: recipient_ids,
    message: messages,
    filePathUrl: media_links,
    base64File: [pdfContent]
    // caption: caption
  };
  var fetch_options = {
    method: "post",
    contentType: "application/json",
    headers: {
      ...(apiKey && { "x-api-key": apiKey })
    },
    payload: JSON.stringify(payload),
  };

  try {
    var response = UrlFetchApp.fetch("https://app.messageautosender.com/api/v1/message/create", fetch_options);
    var status_code = response.getResponseCode();
    if(Math.floor(status_code/100) == 2) 
      return "success";
    else
      return "failed";
  } catch (err) {
    console.log("Error occurred while sending whatsapp message");
    console.error(err);
  }
  return "failed";
}

function set_api_key(){
  var ui = SpreadsheetApp.getUi();
  var promptResponse = ui.prompt("Enter API Key", "", ui.ButtonSet.OK);
  var api_key = promptResponse.getResponseText();
  setDocumentProperty("API_KEY", api_key);
}

function get_api_key(){
  return getDocumentProperty("API_KEY");
}

function reset_api_key(){
  return resetDocumentProperty("API_KEY");
}

function getDocumentProperty(key){
  var documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperty(key);
}

function setDocumentProperty(key, value){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(key, value);
}

function resetDocumentProperty(key){
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.deleteProperty(key);
}

/** Function to get google drive file  */
function getGoogleFileAsBase64(fileId) {
  var result = {};
  var file = DriveApp.getFileById(fileId);
  result = {
    name: file.getName(),
    body: Utilities.base64Encode(file.getBlob().getBytes()),
  };
  return result;
}
