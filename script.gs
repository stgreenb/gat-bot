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


function newChart() {
  // Generate a chart representing the data in the range of A1:B3.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responsesSheet = ss.getSheetByName("responses");

  var chart = responsesSheet.newChart()
     .setChartType(Charts.ChartType.BAR)
     .addRange(responsesSheet.getRange('A1:B3'))
     .setPosition(5, 5, 0, 0)
     .build();

  responsesSheet.insertChart(chart);
}

function startPoll(){
  // open spreadsheet and get tabs
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("logs");
  var userInputSheet = ss.getSheetByName("userInput");
  var dataStorageSheet = ss.getSheetByName("dataStorage");
  var recipientsSheet = ss.getSheetByName("recipients");
  var questionSheet = ss.getSheetByName("questions");
  
  // Read user's input
  // getSheetValues(startRow, startColumn, numRows, numColumns)
  
  var userData = userInputSheet.getSheetValues(1, 1, 4, 2);
  
  // Get the first question
  var q1 = questionSheet.getSheetValues(2,2,1,1);
  // Ensure that q1 is not empty. 
  if (q1 == ""){
    SpreadsheetApp.getUi().alert('Q1 empty!');
    return;
  }
  // Append the response choices to the question
  q1 = q1 + " <BR> " + questionSheet.getSheetValues(2,3,1,1);  
  // keep track of used lines in the logs spreadsheet
  var ssLines = 1;

  // log the userInput to logs sheet

for (var i = 0; i < 4; i++) { 
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
        "markdown" : userData[2][1] + ' (' + userData[3][1] + ') has started a poll. <BR>Please answer the following, Q1:<BR> **|' + q1 + '|**',
      }
    };
  
  //loop through all users and send them the poll info. 
  for (var i = 1; i < recipientsValues.length; i++) {
    postMessageOptions.payload.toPersonEmail = recipientsValues[i][1];
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
  
  var userData = userInputSheet.getSheetValues(1, 1, 4, 2);
  
  // keep track of used lines in the logs spreadsheet
  var ssLines = 1;

  // log the userInput to logs sheet

  for (var i = 0; i < 4; i++) { 
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
        "markdown" : userData[2][1] + ' (' + userData[3][1] + ') has ended the poll. <BR> _Note, this bot will now shutdown, contact ' + userData[2][1] + ' directly for more info._',
      }
    };
  
  //loop through all users and send them the poll ending info. 
  for (var i = 1; i < recipientsValues.length; i++) {
    postMessageOptions.payload.toPersonEmail = recipientsValues[i][1];
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



// ---- Post handler Inbound webhook----
function doPost(e) {
  // open spradsheet and get tabs
  var ss= SpreadsheetApp.getActiveSpreadsheet();
  var logsSheet = ss.getSheetByName("logs");
  var userInputSheet = ss.getSheetByName("userInput");
//  var recipientsSheet = ss.getSheetByName("recipients");
  var responsesSheet = ss.getSheetByName("responses");
  var questionSheet = ss.getSheetByName("questions");
  
  // getSheetValues(startRow, startColumn, numRows, numColumns)
  var userData = userInputSheet.getSheetValues(1, 2, 4, 1);
  var questions = questionSheet.getSheetValues(2,2,10,4);
  
  // keep track of used lines in the logs spreadsheet
  var ssLines = 1;

  // log the bearer token to the logs sheet
//  logsSheet.getRange(ssLines, 1).setValue('Bearer token');
//  logsSheet.getRange(ssLines, 2).setValue(userData[0][0]);
//  ssLines++;
 
  // get the input data from the webhook and parse it into an object
  var jsonedInput = JSON.parse(e.postData.contents);
  var inputData = jsonedInput.data;

  //loop through the input data and log each key and value
  for (var property in inputData) {
  if (inputData.hasOwnProperty(property)) {
//    logsSheet.getRange(ssLines, 1).setValue('inputData: '+ property);
//    logsSheet.getRange(ssLines, 2).setValue(inputData[property]);
//    ssLines++;
  }
}

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
//    logsSheet.getRange(ssLines, 1).setValue('jsonGetMessageDetails: '+ property2);
//    logsSheet.getRange(ssLines, 2).setValue(jsonGMD[property2]);
//    ssLines++;
  }
}
  //is the message from a person (exclude messages from the bot or other bots)?
  var notBot = (jsonGMD.personEmail.indexOf("sparkbot.io") < 0 && jsonGMD.personEmail.indexOf("webex.bot") < 0);
  if (notBot) { 

    
    
  // Determine if the user has responded previously 
  
  // This represents ALL the data on the recipientsSheet
  var responsesRange = responsesSheet.getDataRange();
  // Returns a two-dimensional array of values, indexed by row, then by column
  var responsesValues = responsesRange.getValues();
//  logsSheet.getRange(ssLines, 1).setValue('responsesValues: ');
//  logsSheet.getRange(ssLines, 2).setValue(responsesValues.toString());
//  ssLines++;
  
  // varabile to hold the users row (if remains 0, then user not in response sheet)
  var userRow = 0;
  // variable to hold the column of the users's last response (i.e # of questions responded)
  var userColumn = 1;
 
  
  //find the users email in column A
//  logsSheet.getRange(ssLines, 1).setValue('responsesValues.length');
//  logsSheet.getRange(ssLines, 2).setValue(responsesValues.length);
//  ssLines++;
  
    for(var i = 0; i<responsesValues.length;i++){
    if(responsesValues[i][0] == jsonGMD.personEmail){ //[0] because column A
      userRow = i+1; // row of the found user
      
      for (var j=1; j<11;j++){
        if (responsesValues[i][j] == ""){
          userColumn += j;
          break;
        }
      }
      
//      logsSheet.getRange(ssLines, 1).setValue('user found on Row: ');
//      logsSheet.getRange(ssLines, 2).setValue(userRow);
//      ssLines++;
//      logsSheet.getRange(ssLines, 1).setValue('column: ');
//      logsSheet.getRange(ssLines, 2).setValue(userColumn);
//      ssLines++;
      break;
       }
  }
     logsSheet.getRange(ssLines, 1).setValue('column: ');
     logsSheet.getRange(ssLines, 2).setValue(userColumn);
     ssLines++;
    
  //find a new row to put the data into
  var newRow = responsesValues.length+1;
//  logsSheet.getRange(ssLines, 1).setValue('newRow: ');
//  logsSheet.getRange(ssLines, 2).setValue(newRow);
//  ssLines++;
  
//create a header that can be used for webex team messages back to the user.  
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
    var postMessage;
    

    //if the user was not in the response table (they have not answered a question) add a new row for them, if not append their answer. 
    if (userRow == 0){
      // set up some variables to ensure the response is valid. Account for type and case differences. 
      var lowerValue = questions[0][2];
      var upperValue = questions[0][3];
      var responseStringLower = jsonGMD.text.toLowerCase();
      if (typeof(lowerValue)== "string"){
        lowerValue = lowerValue.toLowerCase();
      }
      if (typeof(upperValue)== "string"){
        upperValue = upperValue.toLowerCase();
      }
      // if the response is in the right range. 
      if (lowerValue <= responseStringLower && responseStringLower <= upperValue){
          responsesSheet.getRange(newRow, 1).setValue(jsonGMD.personEmail);
          responsesSheet.getRange(newRow, 2).setValue(jsonGMD.text);
          userColumn = 2; // set right column for next question verifcation
      } else {
        // if not, tell user and break out. 
        postMessageOptions.payload.markdown = "Valid responses are from "+ lowerValue + " to " + upperValue;
        postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);
        return ContentService.createTextOutput(JSON.stringify(postMessage));    
      }
      
    } else {
      // set up some variables to ensure the response is valid. Account for type and case differences. 
      var lowerValue = questions[userColumn-2][2]; //userColumn will be 3 for the 2nd response but the array starts at 0, so subtract 2
      var upperValue = questions[userColumn-2][3];
      var responseStringLower = jsonGMD.text.toLowerCase();
      if (typeof(lowerValue)== "string"){
        lowerValue = lowerValue.toLowerCase();
      }
      if (typeof(upperValue)== "string"){
        upperValue = upperValue.toLowerCase();
      }
      // if the response is in the right range. 
      if (lowerValue <= responseStringLower && responseStringLower <= upperValue){
         responsesSheet.getRange(userRow, userColumn).setValue(jsonGMD.text);
      } else{
         // if not, tell user and break out. 
        if (lowerValue == upperValue){
          postMessageOptions.payload.markdown = "Your input has already been saved. If you need to make a correction / addtion, contact " + userData[2][0] + ' (' + userData[3][0] + ')';
        } else
        { postMessageOptions.payload.markdown = "Valid responses are from "+ lowerValue + " to " + upperValue;
        }
        postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);
        return ContentService.createTextOutput(JSON.stringify(postMessage));    
    }
    }    

    postMessageOptions.payload.markdown = "_I got your answer:_ " + jsonGMD.text;
    postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);

    // if last question, thank user, else send next question. 
    
    //userColumn will be 2 for the first response, 3 for the second, so subtract 1
    var nextQuestion = userColumn-1;
    if (questions[nextQuestion][0] == ""){
      logsSheet.getRange(ssLines, 1).setValue('next question: ');
      logsSheet.getRange(ssLines, 2).setValue('none');
      ssLines++;   
      postMessageOptions.payload.markdown = "Thank you for participating. ";
      postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);      
    } else {
      logsSheet.getRange(ssLines, 1).setValue('next question: ');
      logsSheet.getRange(ssLines, 2).setValue(questions[nextQuestion][0]);
      ssLines++;
      postMessageOptions.payload.markdown = 'Q' + (userColumn) + ':<BR> **|' + questions[nextQuestion][0] + ' <BR> ' + questions[nextQuestion][1] + '|**';
      postMessage = UrlFetchApp.fetch('https://api.ciscospark.com/v1/messages/', postMessageOptions);         
    }
    
    // parse response, loop through and log results. 
/*    
    var jsonPostMessage = JSON.parse(postMessage);
    for (var property3 in jsonPostMessage) {
      if (jsonPostMessage.hasOwnProperty(property3)) {
        logsSheet.getRange(ssLines, 1).setValue('jsonPostMesage: '+ property3);
        logsSheet.getRange(ssLines, 2).setValue(jsonPostMessage[property3]);
        ssLines++;
      }
    }  
    */
  }
return ContentService.createTextOutput(JSON.stringify(postMessage));
}

/*
// ---- Fires when a new form is submitted ----
function onEdit(e) {
  
  if (e.range.getSheet().getName() ==  "recipients"){
    // open spradsheet and get tabs
    var ss= SpreadsheetApp.getActiveSpreadsheet();
    var logsSheet = ss.getSheetByName("logs");
    var dataStorageSheet = ss.getSheetByName("dataStorage");
    var userInputSheet = ss.getSheetByName("userInput");
    var dataStorageSheet = ss.getSheetByName("dataStorage");
    var questionSheet = ss.getSheetByName("questions");
  
    // Read user's input
    // getSheetValues(startRow, startColumn, numRows, numColumns)

    var questions = questionSheet.getSheetValues(2,2,10,3);
    
    var userData = userInputSheet.getSheetValues(1, 1, 6, 2);
    // keep track of used lines in the logs spreadsheet
    var ssLines = 1;

    // log the bearer token to the logs sheet
    logsSheet.getRange(ssLines, 1).setValue('Bearer token');
    logsSheet.getRange(ssLines, 2).setValue(userData[0][1]);
    ssLines++;
  
    // read current webhook id
    // getSheetValues(startRow, startColumn, numRows, numColumns)
    var webhookID = dataStorageSheet.getSheetValues(1,1,1,1);
    logsSheet.getRange(ssLines, 1).setValue("webhookID-length");
    logsSheet.getRange(ssLines, 2).setValue(webhookID[0][0].length);
    ssLines++;
    logsSheet.getRange(ssLines, 1).setValue("e");
    logsSheet.getRange(ssLines, 2).setValue(e);
    ssLines++;
    logsSheet.getRange(ssLines, 1).setValue("e.range");
    logsSheet.getRange(ssLines, 2).setValue(e.range);
    ssLines++;
    logsSheet.getRange(ssLines, 1).setValue("e.range.getSheet().getName()");
    logsSheet.getRange(ssLines, 2).setValue(e.range.getSheet().getName());
    ssLines++;

  // if user has been added there is a current webhook (ie poll running) then send them the current poll message
  if (webhookID[0][0].length > 1){ 
    // Create a baseline header
      var postMessageOptions = {
      "method": "POST",
      "headers": {
      "Authorization": "Bearer "+ userData[0][1],
      "contentType": "application/json",
      },
      "payload": {
        "markdown" : userData[2][1] + ' (' + userData[3][1] + ') already has a poll running. Please answer the following:<BR> **' + userData[3][1] + '**',
      }
    };

  }
}

*/
