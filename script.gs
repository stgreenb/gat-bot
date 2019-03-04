/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "Start Poll", functionName: "startPoll"});
  menuEntries.push(null); // line separator
  menuEntries.push({name: "Stop Poll", functionName: "stopPoll"});

  ss.addMenu("Poll", menuEntries);
}
function startPoll(){
  // open spreadsheet and get tabs
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("logs");
  var userInputSheet = ss.getSheetByName("userInput");
  var dataStorageSheet = ss.getSheetByName("dataStorage");
  var recipientsSheet = ss.getSheetByName("recipients");
  
  // Read user's input
  // getSheetValues(startRow, startColumn, numRows, numColumns)
  
  var userData = userInputSheet.getSheetValues(1, 1, 6, 2);
  
  // keep track of used lines in the logs spreadsheet
  var ssLines = 1;

  // log the userInput to logs sheet

for (var i = 0; i < 6; i++) { 
  logsSheet.getRange(ssLines, 1).setValue(userData[i][0]);
  logsSheet.getRange(ssLines, 2).setValue(userData[i][1]);
  ssLines++;
}
  // read current webhook id
  // getSheetValues(startRow, startColumn, numRows, numColumns)
  var webhookID = dataStorageSheet.getSheetValues(1,1,1,1);
  logsSheet.getRange(ssLines, 1).setValue("webhookID-length");
  logsSheet.getRange(ssLines, 2).setValue(webhookID[0][0].length);
  ssLines++;
  // if user has selected the start checkbox and there is no current webhook then create a new webhook
  if (webhookID[0][0].length<1){ 
  
  var createWebhookOptions = {
      "method": "POST",
      "headers": {
      "Authorization": "Bearer "+ userData[0][1],
      "contentType": "application/json",
      },
      "payload": {
        "name" : 'gsheet webhook',
        "targetUrl" : userData[1][1],
        "resource" : 'messages',
        "event": "created",
      }
    };

    var createWebhook = UrlFetchApp.fetch('https://api.ciscospark.com/v1/webhooks/', createWebhookOptions);
    // parse response, loop through and log results. 
    var jsonCreateWebhook = JSON.parse(createWebhook);
    for (var property in jsonCreateWebhook) {
      if (jsonCreateWebhook.hasOwnProperty(property)) {
        logsSheet.getRange(ssLines, 1).setValue('jsonCreateWebhook: '+ property);
        logsSheet.getRange(ssLines, 2).setValue(jsonCreateWebhook[property]);
        ssLines++;
      }
    }
    //store the webhook id 
    dataStorageSheet.getRange(1,1).setValue(jsonCreateWebhook.id);
   
  // This represents ALL the data on the recipientsSheet
  var recipientsRange = recipientsSheet.getDataRange();
  var recipientsValues = recipientsRange.getValues();

  // Create a baseline header
      var postMessageOptions = {
      "method": "POST",
      "headers": {
      "Authorization": "Bearer "+ userData[0][1],
      "contentType": "application/json",
      },
      "payload": {
        "markdown" : userData[4][1] + ' (' + userData[5][1] + ') has started a poll. Please answer the following:<BR> **' + userData[3][1] + '**',
      }
    };
  
  //loop through all users and send them the poll info. 
  for (var i = 1; i < recipientsValues.length; i++) {
    postMessageOptions.payload.toPersonEmail = recipientsValues[i][0];
    var postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);
    // parse response, loop through and log results. 
    var jsonPostMessage = JSON.parse(postMessage);
    for (var property3 in jsonPostMessage) {
      if (jsonPostMessage.hasOwnProperty(property3)) {
        logsSheet.getRange(ssLines, 1).setValue('jsonPostMesage: '+ property3);
        logsSheet.getRange(ssLines, 2).setValue(jsonPostMessage[property3]);
        ssLines++;
      }
    }  
} 
} else{
    SpreadsheetApp.getUi().alert('Poll already running!');
  }
}


function stopPoll(){

  
  // open spreadsheet and get tabs
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("logs");
  var userInputSheet = ss.getSheetByName("userInput");
  var dataStorageSheet = ss.getSheetByName("dataStorage");
  var recipientsSheet = ss.getSheetByName("recipients");
  
  // Read user's input
  // getSheetValues(startRow, startColumn, numRows, numColumns)
  
  var userData = userInputSheet.getSheetValues(1, 1, 6, 2);
  
  // keep track of used lines in the logs spreadsheet
  var ssLines = 1;

  // log the userInput to logs sheet

  for (var i = 0; i < 6; i++) { 
    logsSheet.getRange(ssLines, 1).setValue(userData[i][0]);
    logsSheet.getRange(ssLines, 2).setValue(userData[i][1]);
    ssLines++;
  } 
  // read current webhook id
  // getSheetValues(startRow, startColumn, numRows, numColumns)
  var webhookID = dataStorageSheet.getSheetValues(1,1,1,1);
  logsSheet.getRange(ssLines, 1).setValue("webhookID-length");
  logsSheet.getRange(ssLines, 2).setValue(webhookID[0][0].length);
  ssLines++;
  // if user has selected the unselected checkbox and there is a current webhook then delte the webhook
  if (webhookID[0][0].length > 1){ 
  
    var deleteWebhookOptions = {
      "method": "DELETE",
      "headers": {
      "Authorization": "Bearer "+ userData[0][1],
      "contentType": "application/json",
      },
    };
    var deleteWebhook = UrlFetchApp.fetch('https://api.ciscospark.com/v1/webhooks/'+ webhookID[0][0], deleteWebhookOptions);
    // parse response, loop through and log results. 
      logsSheet.getRange(ssLines, 1).setValue('deleteWebhook: ');
      logsSheet.getRange(ssLines, 2).setValue(deleteWebhook);
      ssLines++;
    //store the webhook id 
    dataStorageSheet.getRange(1,1).setValue("");
    
      // This represents ALL the data on the recipientsSheet
  var recipientsRange = recipientsSheet.getDataRange();
  var recipientsValues = recipientsRange.getValues();

  // Create a baseline header for poll ending
      var postMessageOptions = {
      "method": "POST",
      "headers": {
      "Authorization": "Bearer "+ userData[0][1],
      "contentType": "application/json",
      },
      "payload": {
        "markdown" : userData[4][1] + ' (' + userData[5][1] + ') has ended the poll. <BR> _Note, this bot will now shutdown, contact ' + userData[4][1] + ' directly for more info._',
      }
    };
  
  //loop through all users and send them the poll ending info. 
  for (var i = 1; i < recipientsValues.length; i++) {
    postMessageOptions.payload.toPersonEmail = recipientsValues[i][0];
    var postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);
    // parse response, loop through and log results. 
    var jsonPostMessage = JSON.parse(postMessage);
    for (var property3 in jsonPostMessage) {
      if (jsonPostMessage.hasOwnProperty(property3)) {
        logsSheet.getRange(ssLines, 1).setValue('jsonPostMesage: '+ property3);
        logsSheet.getRange(ssLines, 2).setValue(jsonPostMessage[property3]);
        ssLines++;
      }
    }  
} 
    
    
    
  }else{
    SpreadsheetApp.getUi().alert('Poll not running!');
  }
  
}

// ---- Post handler ----
function doPost(e) {
  // open spradsheet and get tabs
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("logs");
  var userInputSheet = ss.getSheetByName("userInput");
  var recipientsSheet = ss.getSheetByName("recipients");
  
  // getSheetValues(startRow, startColumn, numRows, numColumns)
  var userData = userInputSheet.getSheetValues(1, 2, 1, 1);
  
  // keep track of used lines in the logs spreadsheet
  var ssLines = 1;

  // log the bearer token to the logs sheet
  logsSheet.getRange(ssLines, 1).setValue('Bearer token');
  logsSheet.getRange(ssLines, 2).setValue(userData[0][0]);
  ssLines++;
 
  // get the input data from the webhook and parse it into an object
  var jsonedInput = JSON.parse(e.postData.contents);
  var inputData = jsonedInput.data;

  //loop through the input data and log each key and value
  for (var property in inputData) {
  if (inputData.hasOwnProperty(property)) {
    logsSheet.getRange(ssLines, 1).setValue('inputData: '+ property);
    logsSheet.getRange(ssLines, 2).setValue(inputData[property]);
    ssLines++;
  }
}
/*
  // Make a test POST request with a JSON payload to a bin.
  var options1 = {
    'method' : 'post',
    'contentType': 'application/json',
    // Convert the JavaScript object to a JSON string.
    'payload' : JSON.stringify(inputData)
  };

  var response1 = UrlFetchApp.fetch('https://gsparrot.requestcatcher.com/test', options1);
    sheet.getRange(ssLines, 1).setValue('response 1: ');
    sheet.getRange(ssLines, 2).setValue(response1);
    ssLines++;
 */
  // Get message details via GET request from the inputData.id
  var getMessageDetailsOptions = {
  "method": "GET",
  "headers": {
    "Authorization": "Bearer " + userData[0][0],
  }
};
  var getMessageDetails = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/' + inputData.id, getMessageDetailsOptions);
  
  //JSON parse then loop through and log the results. 
  var jsonGMD = JSON.parse(getMessageDetails);
  for (var property2 in jsonGMD) {
  if (jsonGMD.hasOwnProperty(property2)) {
    logsSheet.getRange(ssLines, 1).setValue('jsonGetMessageDetails: '+ property2);
    logsSheet.getRange(ssLines, 2).setValue(jsonGMD[property2]);
    ssLines++;
  }
}
  
  // This represents ALL the data on the recipientsSheet
  var recipientsRange = recipientsSheet.getDataRange();
  var recipientsValues = recipientsRange.getValues();
  
  // varabile to hold the users row
  var userRow = 0;
  
  //find the users email in column A
  
    for(var i = 0; i<recipientsValues.length;i++){
    if(recipientsValues[i][0] == jsonGMD.personEmail){ //[0] because column A
      userRow = i+1; // row of the found user
      break;
       }
  }
  
  //is the message from a person (exclude messages from the bot or other bots)?
  var notBot = (jsonGMD.personEmail.indexOf("sparkbot.io") < 0 && jsonGMD.personEmail.indexOf("webex.bot") < 0);
  
  // Send back the message text to the user (if they are not a bot)
  if (notBot) { 
    
    recipientsSheet.getRange(userRow, 2).setValue(jsonGMD.personEmail);
    recipientsSheet.getRange(userRow, 3).setValue(jsonGMD.text);
    recipientsSheet.getRange(userRow, 4).setValue(jsonGMD.created);

    
    var postMessageOptions = {
      "method": "POST",
      "headers": {
      "Authorization": "Bearer "+ userData[0][0],
      "contentType": "application/json",
      },
      "payload": {
        "roomId" : jsonGMD.roomId,
        "markdown" : "_I got your answer:_ " + jsonGMD.text,
      }
    };

    var postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);
    // parse response, loop through and log results. 
    var jsonPostMessage = JSON.parse(postMessage);
    for (var property3 in jsonPostMessage) {
      if (jsonPostMessage.hasOwnProperty(property3)) {
        logsSheet.getRange(ssLines, 1).setValue('jsonPostMesage: '+ property3);
        logsSheet.getRange(ssLines, 2).setValue(jsonPostMessage[property3]);
        ssLines++;
      }
    }  
  }
return ContentService.createTextOutput(JSON.stringify(postMessage));
}
