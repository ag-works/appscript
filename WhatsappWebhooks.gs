
function doPost(request){
  let headers = ["Timestamp", "Type", "Receiver Number", "Group Name", "Sender Name", "Sender Number", 
                 "Message", "Attachment", "Caption"];
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Messages");
  if(!sheet){
    sheet = ss.insertSheet("Messages");
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  let lock = LockService.getScriptLock(); 
  lock.waitLock(5000); 

  let lastRowNumber = sheet.getDataRange().getLastRow();

  let jsonString = request.postData.getDataAsString();
  let jsonData = JSON.parse(jsonString);
  let timestamp = new Date(jsonData.time).toString();
  let values = [
    timestamp,                                               // Timestamp of received message
    jsonData.itemType || "-",                                // Type of message
    jsonData.receiverNumber, 
    jsonData.authorId ? jsonData.senderName : "-",           // Mobile Number of the message receiver
    jsonData.authorName || jsonData.senderName,              // Sender Name
    jsonData.authorId || jsonData.senderNumber,              // Mobile Number of Sender
    jsonData.value || "-",                                   // Text Message 
    jsonData.filePath || "-",                                // Link of the received document, image or file
    jsonData.caption || "-"                                  // Caption of the received document, image or file
  ];
  sheet.getRange(lastRowNumber + 1, 1, 1, values.length).setValues([values]);
  lock.releaseLock();

  return ContentService.createTextOutput(JSON.stringify(jsonData)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(request) {
  Logger.log("GET", request);
  return ContentService.createTextOutput(JSON.stringify({"message": "I'm Awesome too!"}));
}