//====================================================================||
//!!!!!!!!!!!!!!!!!!!JSON LOGGER SCRIPT!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!||
//====================================================================||

//var emailAddress = "emailtosalesforce@v-1kwer43cl0yzcmes1dmcusodlczba5ca75ti344n4xw4pxlcyq.6c-p09uae.cs63.le.sandbox.salesforce.com";  //email to SalesForce Email SandBox
var emailAddress ="emailtosalesforce@2kg0rvcepytvbe6hrv3b8takh6taxfca37katvqw18p8j3wmjt.8-ldaheao.na8.le.salesforce.com"; //email to SalesForce Email Live
//var emailAddress ="samstevenm@gmail.com"; //email to SalesForce Email Live
var subject = "Uni Pricer JSON"; //email subject
var data= "Email body has no content. Something went wrong."; //initialize data, this will be the email body
var refLink= "SalesForce_Reflink_not_specified."; //initialize refLink, if something goes wrong it
var personalRecordEmail= "samstevenm@gmail.com" 

function sendEmails() {
 
  var sheet = SpreadsheetApp.getActiveSheet(); //Get the active sheet
 
  var titleRow = 8;  //The Row with titles HARDCODE
  //var projectName = sheet.getRange(5, 2, 1, 1);  //The CELL with project name HARDCODE
  var lastRow = sheet.getLastRow();   //Find the last Row
  var lastColumn = sheet.getLastColumn();   //Find the last Column
  
  //var refLink = sheet.getRange(lastRow, 2).getValues().toString().split(".com/").pop(); //Get the REF Link from SalesForce URL in Last Row, Column B
  //var personalRecordEmail = sheet.getRange(lastRow, lastColumn).getValues().toString(); //Get the email address from the Last Row, Last Column
  
  var titles = sheet.getRange(titleRow, 1, 1, lastColumn); // Get title range (first row)
  var responses = sheet.getRange(lastRow, 1, 1, lastColumn); //Get response range (last row)
                    //getRange(row, column, numRows, numColumns) <---- getRange sample function
  
  var titlesResponses = ""; //initialize titles and responses
  
  var numRows = titles.getNumRows(); //count the rows of the title range, it should be 1 row
  var numCols = titles.getNumColumns(); //count the columns of the title range, it will vill vary based on lastColumn
  
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
      var title = titles.getCell(i,j).getValue().replace(/[^A-Z0-9]+/ig, "_"); //get rid of bad characters
      var response = responses.getCell(i,j).getValue().toString().trim().replace(/,/g, '').replace(/\s\s+/g, ',').replace(/\n/g, ',').split(","); //clean
      var response = JSON.stringify(response); //make it JSON

      if (response.length == 2) {(response = '_BLANK_')}; //If a box is left blank, show "UNANSWERED"
      var titleResponse = '\"' + title + '\" : ' + response + ','; //Concatenate title and response
      var titlesResponses = titlesResponses + titleResponse; //Append the ' "title": "response", 'to the rest
    }
  }
  
  var data = '[{';
  var data = data + titlesResponses; //Append titlesResponses to data
  var data = data + '"ref_link\" : [' +'\"' + refLink +'\"]} ]'; //Append RefLink, for automatic association in SalesForce
  var data = JSON.stringify(JSON.parse(data), undefined, 1); //make it pretty
  
  console.log(data);
 
  
  MailApp.sendEmail(
    {
    to: emailAddress,
    cc: personalRecordEmail,
    bcc:"",
    noReply: true,
    //replyTo: "vivepoe@lutron.com",
    subject: subject,
    body: data
    }
    );

  }
